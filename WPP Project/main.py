# Libraries
import pulp
import pandas as pd
import statistics as sta
from tkinter import *
from PIL import ImageTk, Image
from datetime import datetime
from pulp import LpStatus
from pulp import lpSum


def hugefunc():
    # User Final Inputs
    maxtime = int(entry1.get())
    wwhour = int(entry2.get())
    additional_pen = int(entry3.get())

    # Excel Read and Convert
    data_job = pd.read_excel(r'other_inputs.xlsx', sheet_name='input_3')
    data_locmat = pd.read_excel(r'other_inputs.xlsx', sheet_name='input_4')
    data_ava = pd.read_excel(r'other_inputs.xlsx', sheet_name='input_5')
    data_lastloc = pd.read_excel(r'other_inputs.xlsx', sheet_name='input_6')
    data_techmatrix = pd.read_excel(r'Yetkinlik_matris.xlsx', sheet_name='Saha Hiz Tek. Yetkinlik Matrisi')

    job_num = list(data_job['İş Emri Numarası'])
    job_def = list(data_job['İş Emri Açıklaması'])
    job_loc = list(data_job['Konum'])
    job_tech = list(data_job['Gereken Personel Sayısı'])
    job_skill = list(data_job['Yetkinlik İsteri'])
    job_multiplier = list(data_job['Önem Katsayısı'])
    job_hour = list(data_job['Min. İş Emri Çözüm Süresi'])
    last_loc = list(data_lastloc['Anlık Konum'])

    locmat = data_locmat.to_numpy()
    tech_data = data_techmatrix.to_numpy()
    tech_ava = data_ava.to_numpy()
    s_numcol = len(data_techmatrix.columns)

    tech_name = []
    tech_score = []
    score_name = []
    score_point = []
    tech_lic = []

    for i in range(1, int(s_numcol / 2)):
        if tech_data[96][2 * i + 1] == 'V':
            tech_lic += [1]
        else:
            tech_lic += [0]

    for i in range(1, int(s_numcol / 2)):
        tech_name += [tech_data[5][2 * i + 1]]

    for j in range(1, 21):
        score_name += [tech_data[6 + j][1]]

    # additional functions
    def ability(a, b):
        if a == 'hepsi':
            rookie_1 = []
            for i in range(1, 21):
                rookie_1 += [tech_data[6 + i][3 + 2 * b]]

            return sta.mean(rookie_1)

        else:
            rookie_2 = a.split(',')
            rookie_3 = []
            for i in rookie_2:
                rookie_3 += [tech_data[6 + int(i)][3 + 2 * b]]

            return sta.mean(rookie_3)

    def hour(a, b):
        if type(a) == int:
            rookie1 = job_loc[a - 1]
        else:
            rookie1 = a

        if type(b) == int:
            rookie2 = job_loc[b - 1]
        else:
            rookie2 = b
        rookie3 = 0
        rookie4 = 0
        for i in range(len(data_locmat.columns) - 1):
            if rookie1 == locmat[i + 1][0]:
                rookie3 = i
                break
            else:
                continue
        for j in range(len(data_locmat) - 1):
            if rookie2 == locmat[0][j + 1]:
                rookie4 = j
                break
            else:
                continue
        return float(locmat[rookie3 + 1][rookie4 + 1])

    def bin(a, b):
        if a + b == 2:
            return 1
        else:
            return 0

    # Sets
    set_tech = range(1, len(tech_name) + 1)
    set_job = range(1, len(job_num) + 1)
    set_week = range(1, 3)
    set_seq = range(1, 51),

    data_hour = [[0 for w in set_week] for t in set_tech]

    for t in set_tech:
        for w in set_week:
            if tech_ava[t - 1][w] == 'M':
                data_hour[t - 1][w - 1] = wwhour
            else:
                data_hour[t - 1][w - 1] = 0

    # Optimization

    opt = pulp.LpProblem("WPPOptimization", pulp.LpMinimize)

    # Decision Variables
    T = pulp.LpVariable.dicts("assignment", (set_tech, set_job, set_week), cat='Binary')
    A = pulp.LpVariable.dicts("additional_hour", (set_tech, set_week), cat='Continuous')
    L = pulp.LpVariable.dicts('right_bound', (set_tech, set_week), cat='Continuous')
    D = pulp.LpVariable.dicts('dummy_variable', (set_tech, set_job, set_job, set_week), cat='Binary')

    # Objective Function
    opt += lpSum(A[t][w] for t in set_tech for w in set_week) * additional_pen - \
           lpSum(
               T[t][j][w] * ability(job_skill[j - 1], t - 1) * job_multiplier[j - 1] for t in set_tech for j in set_job
               for w in set_week) + \
           lpSum(L[t][w] for t in set_tech for w in set_week)

    # Constraints

    for t in set_tech:
        for w in set_week:
            if data_hour[t - 1][w - 1] == 0:
                opt += A[t][w] == 0

    for j in set_job:
        opt += lpSum(T[t][j][w] * tech_lic[t - 1] for t in set_tech for w in set_week) >= 1

    for j in set_job:
        if job_multiplier[j - 1] >= 2:
            opt += lpSum(T[t][j][1] for t in set_tech) == job_tech[j - 1]
        else:
            opt += lpSum(T[t][j][2] for t in set_tech) == job_tech[j - 1]

    for t in set_tech:
        for w in set_week:
            opt += A[t][w] >= 0
            opt += A[t][w] <= 10
            opt += L[t][w] >= 0

    for t in set_tech:
        for w in set_week:
            opt += lpSum(T[t][j][w] * float(job_hour[j - 1]) for j in set_job) <= data_hour[t - 1][w - 1] + A[t][w]

    for t in set_tech:
        for w in set_week:
            for j1 in set_job:
                for j2 in set_job:
                    opt += D[t][j1][j2][w] <= T[t][j1][w]
                    opt += D[t][j1][j2][w] <= T[t][j2][w]
                    opt += D[t][j1][j2][w] >= T[t][j1][w] + T[t][j2][w] - 1

    for t in set_tech:
        for w in set_week:
            opt += 2*L[t][w] == lpSum(D[t][j1][j2][w] * hour(j1, j2) for j1 in set_job for j2 in set_job)

    opt.solve(pulp.CPLEX_CMD(timeLimit=maxtime, gapRel=0, msg=True))
    print("Status:", LpStatus[opt.status])

    # Excel write

    ct = str(datetime.now().strftime("%d.%m.%Y@%H-%M-%S")) + '.xlsx'
    output = pd.ExcelWriter('%s' % ct, engine='xlsxwriter')

    q1, q2, q3, q4, q5, q6, q7, q72, q8 = [], [], [], [], [], [], [], [], []

    for j in set_job:
        q4.clear()
        for w in set_week:
            for t in set_tech:
                if T[t][j][w].varValue == 1:
                    q4 += [tech_name[t - 1]]
        q1 += [str(' & '.join([str(i) for i in q4]))]

    for j in set_job:
        for w in set_week:
            if sum(T[t][j][w].varValue for t in set_tech) == job_tech[j - 1]:
                q2 += [w]

    job_num = pd.DataFrame(job_num)
    job_def = pd.DataFrame(job_def)
    job_loc = pd.DataFrame(job_loc)
    q1 = pd.DataFrame(q1)
    q2 = pd.DataFrame(q2)
    job_def = job_def.transpose()
    job_num = job_num.transpose()
    job_loc = job_loc.transpose()
    q1 = q1.transpose()
    q2 = q2.transpose()
    final_2 = pd.concat([job_num, job_def, job_loc, q1, q2])
    final_2 = final_2.transpose()
    final_2.columns = ['İş Emri Numarası', 'İş Emri Açıklaması', 'Konum', 'Atanan Teknisyenler', 'Hafta']
    final_2.to_excel(output, sheet_name='output_1')

    for t in set_tech:
        for w in set_week:
            q7 += [tech_name[t - 1]]
            q72 += [w]
            if A[t][w].varValue > 0:
                q8 += [wwhour + A[t][w].varValue]
            else:
                q8 += ['Less than %s hour' % wwhour]

    q7 = pd.DataFrame(q7)
    q72 = pd.DataFrame(q72)
    q8 = pd.DataFrame(q8)
    q7 = q7.transpose()
    q72 = q72.transpose()
    q8 = q8.transpose()
    final_4 = pd.concat([q7, q72, q8])
    final_4 = final_4.transpose()
    final_4.columns = ['Teknisyen', 'Hafta', 'Çalışma Saati']
    final_4.to_excel(output, sheet_name='output_2')

    for t in set_tech:
        for w in set_week:
            q5 += [L[t][w].name]
            q6 += [L[t][w].varValue]
    q5 = pd.DataFrame(q5)
    q6 = pd.DataFrame(q6)
    q5 = q5.transpose()
    q6 = q6.transpose()
    final_1 = pd.concat([q5, q6])
    final_1 = final_1.transpose()
    final_1.to_excel(output, sheet_name='output_3')

    output.save()

    print('Document has been generated')


# Graphical User Interface
huge = Tk()
huge.iconbitmap('icon_wpp.ico')
pilot = Canvas(huge, height=700, width=853)
pilot.pack()

img = ImageTk.PhotoImage(Image.open('image_wpp.jpg'))
imglabel1 = Label(image=img)
imglabel1.place(relx=0, rely=0.010)

input1 = Label(huge, text='Runtime limit')
input2 = Label(huge, text='Weekly working hour')
input3 = Label(huge, text='Additional hour penalty')
developed = Label(huge, text='Developed by Boran Deniz BAKIR')
developed.configure(font=('Times New Roman', 10, 'bold'))
developed.place(relx=0.78, rely=0.97)
input1.place(relx=0.35, rely=0.70)
input2.place(relx=0.35, rely=0.75)
input3.place(relx=0.35, rely=0.80)

entry1 = Entry(huge, width=4, fg='red')
entry1.place(relx=0.55, rely=0.70)
entry1.insert(0, '60')
entry2 = Entry(huge, width=4, fg='red')
entry2.place(relx=0.55, rely=0.75)
entry2.insert(0, '30')
entry3 = Entry(huge, width=4, fg='red')
entry3.place(relx=0.55, rely=0.80)
entry3.insert(0, '1')

mainbutton = Button(huge, text='Optimize!', command=hugefunc)
mainbutton.place(relx=0.450, rely=0.87)

huge.mainloop()
