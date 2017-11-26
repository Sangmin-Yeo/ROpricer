import numpy as np
import xlsxwriter

### TODO: rf / vol1 / vol2
rf = 0.036
vol1 = 0.4093
vol2 = 0.6285

rho = 1 / (1 + rf)

C_FE_0 = 89.17 * 1000  # KRW/kWh
C_RE_0 = 169.09 * 1000  # KRW/kWh
P_CE_0 = 10000  # KRW/tCO2
CEF_FE_0 = 0.4485  # tCO2/mWh
CEF_RE_0 = 0  # tCO2/mWh
RD_0 = 5000  # billion
A_0 = 5000  # billion

N = 14
# gWh
renewable0 = [20080, 24664, 31165, 34360, 38599, 44350, 54139, 58961, 66759, 80193, 74231, 77364, 83164, 90134]
renewable = np.array(renewable0)

u1 = np.e ** vol1
d1 = 1 / u1
u2 = np.e ** vol2
d2 = 1 / u2

p = ((np.e ** rf) - d1) / (u1 - d1)
q = ((np.e ** rf) - d2) / (u2 - d2)


def A_F(t):
    return A_0


def CEF_FE(t):
    return CEF_FE_0


def CEF_RE(t):
    return CEF_RE_0


def RE(t, k):
    return renewable[t - k]


def C_FE(t, i):
    return C_FE_0 * u1 ** (2 * i - t)


def P_CE(t, j):
    return P_CE_0 * u2 ** (2 * j - t)


def C_RE(t, r):
    return C_RE_0 * (1 - 0.02) ** r


def RD(t):
    return RD_0


def PI(t, i, j, r, k):
    pi = 0
    if pi_mat[t][i][j][4]:
        pi = pi_mat[t][i][j][3]
    else:
        if t == N - 1:
            pi = (CEF_FE(t) - CEF_RE(t)) * P_CE(t, j) * RE(t, k) + (C_FE(t, i) - C_RE(t, r)) * RE(t, k)
            pi_mat[t][i][j][3] = pi
            pi_mat[t][i][j][4] = True
        else:
            pi = (CEF_FE(t) - CEF_RE(t)) * P_CE(t, j) * RE(t, k) + (C_FE(t, i) - C_RE(t, r)) * RE(t, k) + \
                 rho * (p * (q * PI(t + 1, i + 1, j + 1, r, k) + (1 - q) * PI(t + 1, i + 1, j, r, k)) +
                        (1 - p) * (q * PI(t + 1, i, j + 1, r, k) + (1 - q) * PI(t + 1, i, j, r, k)))
            pi_mat[t][i][j][3] = pi
            pi_mat[t][i][j][4] = True
    return pi


def V(t, i, j, r):
    re = ''
    rv = 0
    if t + 1 == N:
        if result_mat[t][i][j][8]:
            A = result_mat[t][i][j][3]
            D = result_mat[t][i][j][4]
            rv = result_mat[t][i][j][6]
        else:
            A = -A_F(t)
            D = PI(t, i, j, r, 0)
            result_mat[t][i][j][3] = A
            result_mat[t][i][j][4] = D
            result_mat[t][i][j][8] = True
            if A > D:
                re = 'A'
                rv = A
            else:
                re = 'D'
                rv = D
            result_mat[t][i][j][6] = rv
            result_mat[t][i][j][7] = re
            print('V t:%d i:%d j:%d A:%d D:%d == %s' % (t, i, j, A, D, re))



    else:
        if result_mat[t][i][j][8]:
            A = result_mat[t][i][j][3]
            D = result_mat[t][i][j][4]
            R = result_mat[t][i][j][5]
            rv = result_mat[t][i][j][6]
            # print('%d %d %d %d' %(t,i,j,rv))
        else:
            A = -A_F(t)
            D = PI(t, i, j, r, t)
            R = rho * (p * (q * V(t + 1, i + 1, j + 1, r + 1) + (1 - q) * V(t + 1, i + 1, j, r + 1)) +
                       (1 - p) * (q * V(t + 1, i, j + 1, r + 1) + (1 - q) * V(t + 1, i, j, r + 1))) - RD(t)
            '''
            if t == 0 and i == 0 and j ==0:

                print("!!!")
                print(result_mat[1][0][0])
                print(result_mat[1][0][1])
                print(result_mat[1][1][0])
                print(result_mat[1][1][1])
                print(R)
                print(r)
                print(p)
                print(q)
            '''
            result_mat[t][i][j][3] = A
            result_mat[t][i][j][4] = D
            result_mat[t][i][j][5] = R

            ra = [A, D, R]
            return_v = ['A', 'D', 'R']
            re = return_v[np.array(ra).argmax()]
            rv = max(ra)
            result_mat[t][i][j][6] = rv
            result_mat[t][i][j][7] = re
            result_mat[t][i][j][8] = True

            print('V t:%d i:%d j:%d A:%d D:%d R:%d== %s' % (t, i, j, A, D, R, re))
            # print(result_mat[t][i][j][6])
    return rv


N = 14
result_mat = []
pi_mat = []
for t in range(N):
    temp_mat_1 = []
    pi_mat_1 = []
    for i in range(t + 1):
        temp_mat_2 = []
        pi_mat_2 = []
        for j in range(t + 1):
            temp_mat_3 = [t, i, j, 0, 0, 0, 0, '', False]  # ADR rv re
            pi_mat_3 = [t, i, j, 0, False]
            temp_mat_2.append(temp_mat_3)
            pi_mat_2.append(pi_mat_3)
        temp_mat_1.append(temp_mat_2)
        pi_mat_1.append(pi_mat_2)
    result_mat.append(temp_mat_1)
    pi_mat.append(pi_mat_1)

print(result_mat)

print(V(0, 0, 0, 0))

workbook = xlsxwriter.Workbook("result.xlsx")
for t in range(N):
    worksheet = workbook.add_worksheet(str(t))
    for i in range(len(result_mat[t])):
        for j in range(len(result_mat[t][i])):
            if i == 0:
                worksheet.write(i, j + 1, str(j))
            if j == 0:
                worksheet.write(i + 1, j, str(i))
            worksheet.write(i + 1, j + 1, result_mat[t][i][j][7])
workbook.close()

workbook = xlsxwriter.Workbook("result_details.xlsx")
for t in range(N):
    worksheet = workbook.add_worksheet(str(t))
    for i in range(len(result_mat[t])):
        for j in range(len(result_mat[t][i])):
            if i == 0:
                worksheet.write(i, j + 1, str(j))
            for k in range(3):
                if j == 0:
                    worksheet.write(i * 3 + 1, j, str(i) + " " + "A")
                    worksheet.write(i * 3 + 2, j, str(i) + " " + "D")
                    worksheet.write(i * 3 + 3, j, str(i) + " " + "R")
                worksheet.write(i * 3 + 1, j + 1, result_mat[t][i][j][3])
                worksheet.write(i * 3 + 2, j + 1, result_mat[t][i][j][4])
                worksheet.write(i * 3 + 3, j + 1, result_mat[t][i][j][5])
workbook.close()

print(result_mat[0][0][0][7])


# a = (CEF_FE(13) - CEF_RE(13)) * P_CE_0 * 90134 + (C_FE(13, 13) - C_RE(13, 13)) * 90134
# b = (CEF_FE(13) - CEF_RE(13)) * P_CE_0 * 90134 + (C_FE(13, 12) - C_RE(13, 13)) * 90134
# print(a, b, a / b)
# print(1635679 / 716036)
