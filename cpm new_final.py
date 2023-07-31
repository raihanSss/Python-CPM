#this is the task module
import numpy as np
import pandas as pd

#error messages
def errorCodeMsg():
    print("Error in input file : CODE ")
    quit()
    

def errorPredMsg():
    print("Error in input file : PREDECESSORS ")
    quit()

def errorDaysMsg():
    print("Error in input file : DAYS ")
    quit()

# Scans if the code in predecessors and successors are
# in the list of task codes:
def getTaskCode(mydata, code):
    x = 0
    flag = 0
    for i in mydata['CODE']:
        if(i == code):
            flag = 1
            break
        
        x+=1
    
    if(flag == 1):
        return x
    else:
        errorCodeMsg()



# Critical Path Method Forward Pass Function
# EF -> Earliest Finish
# ES -> Earliest Start

def forwardPass(mydata):
    ES = [0] * len(mydata)
    EF = [0] * len(mydata)

    for i in range(len(mydata)):
        predecessors = mydata['PREDECESSORS'][i]
        if predecessors is None:
            ES[i] = 0
        else:
            pred = predecessors.split(',')
            pred_indices = [mydata[mydata['CODE'] == p].index[0] for p in pred]
            ES[i] = max(EF[p] for p in pred_indices)
        EF[i] = ES[i] + mydata['DAYS'][i]

    mydata['ES'] = ES
    mydata['EF'] = EF

    return mydata


# Critical Path Method Backward Pass function
# LS -> Latest Start
# LF -> Latest Finish

def backwardPass(mydata):
    LS = [0] * len(mydata)
    LF = [0] * len(mydata)

    max_ef_index = mydata['EF'].idxmax()
    LF[max_ef_index] = mydata['EF'][max_ef_index]

    for i in range(len(mydata) - 1, -1, -1):
        successors = mydata['CODE'][i]
        if successors is None:
            LF[i] = LF[max_ef_index]
        else:
            succ = successors.split(',')
            succ_indices = [mydata[mydata['CODE'] == s].index[0] for s in succ]
            LF[i] = min(LS[s] for s in succ_indices)
        LS[i] = LF[i] - mydata['DAYS'][i]

    mydata['LS'] = LS
    mydata['LF'] = LF

    return mydata

#compute for SLACK and CRITICAL state
def slack(mydata):
    slack = [LF - EF for LF, EF in zip(mydata['LF'], mydata['EF'])]
    critical = ['Yes' if s == 0 else 'No' for s in slack]

    mydata['SLACK'] = slack
    mydata['CRITICAL'] = critical

    return mydata

#wrapper function:

def computeCPM(mydata):
    mydata = forwardPass(mydata)
    mydata = backwardPass(mydata)
    mydata = slack(mydata)
    return mydata

    # Simpan output ke Excel
def saveToExcel(mydata, output_file):
    excel_data = pd.DataFrame(mydata)
    excel_data.to_excel(output_file, index=False)

mydata = pd.DataFrame({
    'DESCR': ['Pemodelan_sistem', 'Kebutuhan_pengguna', 'Desain_grafis', 'Desain_pp', 'Desain_antarmuka', 'Validasi', 
              'Proses_bisnis', 'Kebutuhan_fungsi/fitur', 'Kebutuhan_non', 'Persetujuan_ttd', 
              'Penyusunan_waktu', 'RAB', 'Persetujuanttd', 
              'Penysunan_surat', 'Pemberian_no_surat', 'Persetujuan_ttd_terlibat',
              'Wireframe_prototype', 'Tampilan_fungsi', 'Interaksi_pengguna', 
              'Pemrograman', 
              'Pengujian_otomatis', 'pengecekan_interface',
              'pengujian_Software','Validasi_sistem', 'Evaluasi',
              'Solusi_si',
              'implementasi'],
              
    'CODE': ['A1','A2','A3','A4','A5','A6',
             'B1','B2','B3','B4',
             'C1','C2','C3',
             'D1','D2','D3',
             'E1','E2','E3',
             'F1',
             'G1','G2',
             'H1','H2','H3',
             'I1',
             'J1'],

   'PREDECESSORS': [None, None, 'A1,A2', 'A3', 'A3', 'A4,A5',
                    None, 'B1', 'B1', 'B2,B3', 
                    'B4,A6', 'C1', 'C2', 
                    'C3', 'C3', 'D1,D2', 
                    'D3', 'E1', 'E2', 
                    'E3', 
                    'F1', 'G1,G2', 
                    'H1', 'H2', 'H3',
                    'I1',
                    'J1'],

    'DAYS': [19, 27, 21, 15, 12, 5, 
             19, 27, 30, 11, 
             21, 13, 7, 
             8, 8, 5, 
             14, 10, 11, 
             45,
             36, 21,
             34, 50, 8,
             15,
             20]
})

mydata = forwardPass(mydata)
mydata = backwardPass(mydata)
mydata = slack(mydata)

output_file = 'outputcpm1.xlsx'
saveToExcel(mydata, output_file)

print("Data berhasil disimpan dalam file Excel:", output_file)