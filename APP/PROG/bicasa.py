from barcode.writer import ImageWriter
from barcode import Code39
from reportlab.pdfgen import canvas
import datetime
from fpdf import FPDF
import pandas as pd
from PyPDF2 import PdfFileMerger
from fpdf import FPDF
from os import path
from turtle import pen
from PyPDF2 import PdfFileWriter, PdfFileReader
import pandas as pd
from fpdf import FPDF
import docx
import pandas as pd
import os
from datetime import datetime
from datetime import date
from APP.STD_FILES.postcodefile import STD_GROUP_CODE

def Folder_creation_non_email(kc1):
    # ------------------Capture date----------------
    kc1=str(kc1)
    today =date.today()
    d1 = today.strftime("%d-%b-%Y")
    d2=d1[0:2]
    d3=d1[3:6]
    d1=d2+d3+kc1
    # ------------------Capture time----------------
    now = datetime.now()        
    start_time = now.strftime("%H:%M:%S")
    s1=start_time[0:2]
    s2=start_time[3:5]
    s3=start_time[6:8]
    start_time=s1+s2+s3

    # ---------------create first date folder------------
    # import os
    # directory = d
    # directory=d1
    # parent_dir = r"C:\Users\Lenovo\Desktop\NREQUIBICASA/"

    parent_dir = r"E:\IWOC\OUTPUT\BICASA\CYCLE/"
    directory = d1+"_"+start_time 
    path1 = os.path.join(parent_dir, directory) # date folder path
    os.mkdir(path1)
    # print("Directory '% s' created" % directory)

    directory = 'BIMB'
    parent_dir = path1
    path2 = os.path.join(path1, directory) # bkrcc folder path
    os.mkdir(path2)
    # print("Directory '% s' created" % directory)


    directory = 'PRINTING'
    parent_dir = path2
    path4 = os.path.join(path2, directory) # PRINTING folder path
    os.mkdir(path4)
    # print("Directory '% s' created" % directory)

    return path4

def Folder_creation_email(kc1):
    # ------------------Capture date----------------
    kc1=str(kc1)
    today =date.today()
    d1 = today.strftime("%d-%b-%Y")
    d2=d1[0:2]
    d3=d1[3:6]
    d1=d2+d3+kc1
    # ------------------Capture time----------------
    now = datetime.now()        
    start_time = now.strftime("%H:%M:%S")
    s1=start_time[0:2]
    s2=start_time[3:5]
    s3=start_time[6:8]
    start_time=s1+s2+s3

    # ---------------create first date folder------------
    # import os
    # directory = d
    # directory=d1
    # parent_dir = r"C:\Users\Lenovo\Desktop\NREQUIBICASA/"
    parent_dir = r"E:\IWOC\OUTPUT\BICASA\CYCLE/"
    directory = d1+"_"+start_time 
    path1 = os.path.join(parent_dir, directory) # date folder path
    os.mkdir(path1)
    # print("Directory '% s' created" % directory)

    directory = 'BIMB'
    parent_dir = path1
    path2 = os.path.join(path1, directory) # bkrcc folder path
    os.mkdir(path2)
    # print("Directory '% s' created" % directory)

    directory = 'DIGITAL'
    parent_dir = path2
    path3 = os.path.join(path2, directory) # DIGITAL folder path
    os.mkdir(path3)
    # print("Directory '% s' created" % directory)

    return path3

def BICASA(request,files1):
    move_files = []
    move_files.append(files1)
    # ---------------MESSAGE DATA COLLECTION--------------------------
    path11 = r'E:\IWOC\SOURCE\\'+ files1
    # path11 = r'Copy of SA20250930.txt'

    with open(path11) as file:
        k=file.readlines()
        total_lines=len(k)
#--------------------------------------------POSTCODE CAPTURED FOR MAPPING WITH GP---------------------------
    postcode = []  
    with open(path11) as file:
        for i in range(0,total_lines):
                line=file.readline()
                line=line.strip()
                if line.startswith('.') and len(line)==1:
                    l0=file.readline()
                    l1=file.readline()
                    l2=file.readline().strip()
                    l2 = l2.replace(' ','')
                    l2 = l2[0:5]
                    l3=file.readline().strip()
                    l3 = l3.replace(' ','')
                    l3 = l3[0:5]
                    l4 = file.readline().strip()
                    l4 = l4.replace(' ','')
                    l4 = l4[0:5]
                    l5 = file.readline().strip()
                    l5 = l5.replace(' ','')
                    l5 = l5[0:5]
                    l6 = file.readline().strip()
                    l6 = l6.replace(' ','')
                    l6 = l6[0:5]
                    l7 = file.readline().strip()
                    l7 = l7.replace(' ','')
                    l7 = l7[0:5]
                    l8 = file.readline().strip()
                    l8 = l8.replace(' ','')
                    l8 = l8[0:5]
                    if l2.startswith(('1','2','3','4','5','6','7','8','9','0')) and '/' not in l2 and len(l2) == 5 and  l2.isdigit() and '12312' not in l2:    
                        postcode.append(l2)
                    elif l3.startswith(('1','2','3','4','5','6','7','8','9','0')) and '/' not in l3 and len(l3) == 5  and l3.isdigit() and '12312' not in l3:    
                        postcode.append(l3)
                    elif l4.startswith(('1','2','3','4','5','6','7','8','9','0')) and '/' not in l4 and len(l4) == 5  and l4.isdigit() and '12312' not in l4:    
                        postcode.append(l4)
                    elif l5.startswith(('1','2','3','4','5','6','7','8','9','0')) and '/' not in l5 and len(l5) == 5  and l5.isdigit() and '12312' not in l5:    
                        postcode.append(l5)
                    elif l6.startswith(('1','2','3','4','5','6','7','8','9','0')) and '/' not in l6  and len(l6) == 5  and l6.isdigit() and '12312' not in l6:     
                        postcode.append(l6)
                    elif l7.startswith(('1','2','3','4','5','6','7','8','9','0'))  and '/' not in l7 and len(l7) == 5   and l7.isdigit() and '12312' not in l7:    
                        postcode.append(l7)
                    elif l8.startswith(('1','2','3','4','5','6','7','8','9','0')) and '/' not in l8  and len(l8) == 5  and l8.isdigit() and '12312' not in l8:    
                        postcode.append(l8)
                    else:
                        postcode.append('000')
#--------------------------------------------POSTCODE CAPTURED FOR MAPPING WITH GP COMPLETED---------------------------

        
    with open(path11) as file:
        c=-1
        date_start_line_number=[]
        total_debit=[]
        for line in file:
            c=c+1
            line=line.strip()
            if line.startswith('DATE    DESCRIPTION  '):
                date_start_line_number.append(c)
        date_start_line_number.append(total_lines)
        # print(len(date_start_line_number))

        MESSAGE=[]
        with open(path11) as file:
            lines=file.readlines()     
        kk=0
    # move_files = []
    # move_files.append(files1)
    kc1 = 1
    kk1 = Folder_creation_email(kc1)
    kc2 = 11
    kk2 = Folder_creation_non_email(kc2) 

 #----------------------- data captured from word file-----------------
    doc = docx.Document(r'E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\Copy of DD7405F.docx')
    all_docs = doc.paragraphs

    email_id = []
    Account_no = []

    for i in all_docs:
        a = i.text
        if '@' in a:
            email_id.append(a[20:52])
            Account_no.append(a[5:19])
    df1 = pd.DataFrame()
    df1['account_no'] = Account_no
    df1['email_id'] = email_id


    #---------------email data ends here-----
    with open(path11) as file:
        k=file.readlines()
        total_lines=len(k)
        total_accounts=0
        for i in k:
            i=i.strip()
            if 'ACCOUNT NO' in i:
                total_accounts=total_accounts+1
                file.readline()
                # print(i)
        print('Total Accounts',total_accounts)
        print('Total Lines',total_lines)

    a1=[]
    client_name=[]
    below_client = []
    statement_date=[]
    # add0=[] 
    add1=[]
    add2=[]
    add3=[]
    add4=[]
    account_no=[]
    branch=[]
    date_list=[]
    bal_bf =[]
    desc=[]
    debit = []
    total_debit=[]
    total_date_list = []
    len_list = []
    credit=[]
    total_credit=[]
    page_no = []
    monthly_average=[]
    total_bal = []
    bal = []
    with open(path11) as file:
        for i in file:
            if 'ACCOUNT NO  ' in i:
                account_no.append(i[105:120].strip())

    barcode_below = []
    with open(path11) as file:
        for i in range(0,total_lines):
            line=file.readline()
            line=line.strip()
            if line.startswith('.') and len(line) == 1:
                l1=file.readline().strip()
                barcode_below.append(l1[0:4])
    counter=0
    counter_col=[]
    with open(path11) as file:
        for i in range(0,total_lines):
            line=file.readline()
            line=line.strip()
            if line.startswith('.') and len(line) == 1:
                counter+=1
                counter_col.append(counter)
                l2=file.readline()
                a1.append(l2[90:103])
                l3=file.readline()
                l3=l3.strip()
                client_name.append(l3[0:40])
                l4=file.readline()
                # l4=l4.strip()
                statement_date.append(l4[112:130])
                below_client.append(l4[0:20])
                l5=file.readline()
                # l5=l5.strip()
                add1.append(l5[0:40])
                l5=file.readline()
                # l5=l5.strip()
                add2.append(l5[0:30])
                page_no.append(l5[118:120].strip())
                l6=file.readline()
                # l6=l6.strip()
                add3.append(l6[0:30])
                l7=file.readline()
                add4.append(l7[0:30])
                l8=file.readline()
                l9=file.readline()
                l9=l9[89:115]
                l9=l9.strip()
                branch.append(l9)
    
    # print(page_no)
    #-----------------------Total debit, credit----------------------------------------
    #------In source file we have keyword DATE DESCRIPTION which is start of transation, 
    # -----Here , we have captured start location of this keyword -----------------------------------
    # print(len(account_no),'****************************************')
    # print(page_no,len(page_no))
    with open(path11) as file:
            c=-1
            date_start_line_number=[]
            total_debit=[]
            for line in file:
                c=c+1
                line=line.strip()
                if line.startswith('DATE    DESCRIPTION'):
                    date_start_line_number.append(c)
            date_start_line_number.append(total_lines)
    with open(path11) as file:
            lines=file.readlines()
            file.seek(0)
            line=file.readline()
            num=0
            start_date_list=[]
            for line in file:
                line=line.strip()
                num+=1
                if '/' in line:
                    if line.startswith(('1','2','3','4','5','6','7','8','9','0')):
                        start_date_list.append(num)                
    with open(path11) as file:
            lines=file.readlines()
            msg_total_debit=[]
            msg_total_credit=[]
            monthly_avg=[]
            total_desc0=[]
            total_desc1=[]
            total_desc2=[]
            total_desc3=[]
            total_desc4=[]
            message_list=[]
            a=0
            b=date_start_line_number[1]
            # print(b,'date_start_line_number')
            c=2
            lenght=len(date_start_line_number)
            flag='False'
            while c<=lenght:
                flag1='False'
                flag2='False'
                flag3='False'
                for i in range(a,b):
                    line=lines[i].strip()

                    if 'PROFIT PAID' in line:
                        line1=lines[i+1].strip()
                        line1=line1.strip()
                        line2=lines[i+2].strip()
                        line2=line2.strip()
                        line3=lines[i+3].strip()
                        line3=line3.strip()
                        if 'TOTAL DEBIT' in line1:
                            msg_total_debit.append(line1[25:50])     
                            flag1='True'  
                        if 'TOTAL CREDIT' in line2:
                            msg_total_credit.append(line2[25:50])     
                            flag2='True'  
                        if 'MONTHLY AVERAGE' in line3:
                            monthly_avg.append(line3[25:50])
                            flag3='True'  

                    if 'MESSAGE' in line:
                        ll=file.readline().strip()
                        message_list.append('Testing')
                    else:
                        message_list.append('Test')

                if flag1=='False':
                    pass
                    msg_total_debit.append('Test')              
                if flag1=='False':
                    pass
                    msg_total_credit.append('Test')

                if flag1=='False':
                    monthly_avg.append('Test')              
                if c==lenght:
                    break
                a=b
                b=date_start_line_number[c]
                c=c+1
    # print(message_list,'-------------')
    #-----------------------------------------------------------------------------------
    with open(path11) as file:
            line=file.readline()
            copy = False
            date_list=[]
            desc0=[]
            desc1 = []
            desc2 = []
            desc3 = []
            desc4 = []
            bal_bf = []
                    
            debit = []
            for i in range(0,len(date_start_line_number)-1):
                
                a=date_start_line_number[i]
                b=date_start_line_number[i+1]
                date_line=lines[a]
                date_line=date_line.strip()
                for j in range(a,b):
                    line1 = lines[j]
                    line1 = line1.strip()
                    if j<b-1:
                        line2 = lines[j+1]
                        line2 = line2.strip()
                    if j<b-2:
                        line3 = lines[j+2]
                        line3 = line3.strip()
                    if j<b-3:
                        line4 = lines[j+3]
                        line4 = line4.strip()
                    if j<b-4:
                        line5 = lines[j+4]
                        line5 = line5.strip()

                    if '/' in line1 and line1.startswith(('1','2','3','4','5','6','7','8','9','0')) and 'STATEMENT DATE' not in line1  and '33u' not in line1:
                        desc0.append(line1[8:40].strip())
                        dl=line1[0:8].strip()
                        date_list.append(dl)
                        db=line1[66:72].strip()
                        debit.append(db)
                        ba=line1[106:119].strip()
                        bal.append(ba)
                        cr=line1[89:100].strip()
                        credit.append(cr)
                        
                        if '/' in line2 or line2.startswith('.') or 'MESSAGE' in line5:
                            desc1.append('desc_TEST')
                            desc2.append('desc_TEST')
                            desc3.append('desc_TEST')
                            desc4.append('desc_TEST')
                            continue
                        else:
                            if line2.startswith('TOTAL') or 'MESSAGE' in line2:
                                desc1.append('desc_TEST')
                            else:
                                desc1.append(line2[0:39])
                        if ('/' in line3 or 'T0TAL' in line3 or line3.startswith('.') or line3.startswith('MESSAGE')) and 'MAXXXmaxx' not in line3:
                            desc2.append('desc_TEST')
                            desc3.append('desc_TEST')
                            desc4.append('desc_TEST')
                            continue
                        else:
                            if line3.startswith('TOTAL'):
                                desc2.append('desc_TEST')
                            else:
                                desc2.append(line3[0:20])

                        if ('/' in line4 or 'T0TAL' in line4 or line4.startswith('.')) and 'QQQQQQQQQQQ' not in line4:
                            desc3.append('desc_TEST')
                            desc4.append('desc_TEST')
                            continue
                        else:
                            if line4.startswith('TOTAL') or line4.startswith('MONTHLY'):
                                desc3.append('desc_TEST')
                            else:
                                desc3.append(line4[0:20])
                        
                        if '/' in line5 or 'T0TAL' in line5 or line5.startswith('.') or line3.startswith('.') :
                            # print('000000')
                            desc4.append('desc_TEST')
                            continue
                        else:
                            if line5.startswith('TOTAL') or line5.startswith('MONTHLY') or line5.startswith('00') :
                                desc4.append('desc_TEST')
                            else:
                                desc4.append(line5[0:20])

                    else:
                        continue
                total_debit.append(debit)
                debit = []
                total_credit.append(credit)
                credit=[]
                total_date_list.append(date_list)
                date_list = []
                total_desc0.append(desc0)
                desc0=[]          
                total_desc1.append(desc1)
                desc1=[]          
                total_desc2.append(desc2)
                desc2=[]          
                total_desc3.append(desc3)
                desc3=[]   
                total_desc4.append(desc4)
                desc4=[]   
                total_bal.append(bal)
                bal = []       
            # print(total_desc2)
            for i in range(0,len(date_start_line_number)-1):
                aa=date_start_line_number[i]+2
                l1=lines[aa]
                l1 = l1.strip()
                if l1.startswith('BAL B/F'):
                    # print(l1,'L!----------------------------------')
                    linee=l1[80:120]
                    linee=linee.strip()
                    bal_bf.append(linee)
                else:
                    bal_bf.append('TEST')
            total_date_list.append(date_list)
            total_desc0.append(desc0)
            total_desc1.append(desc1)
            total_desc2.append(desc2)
            total_desc3.append(desc3)
            total_desc4.append(desc4)

            total_debit.append(debit)
            total_bal.append(bal)
            total_credit.append(credit)
            total_date_list.pop()
            total_bal.pop()
            total_debit.pop()
            total_desc0.pop()
            total_desc1.pop()
            total_desc2.pop()
            total_desc3.pop()
            total_desc4.pop()
            total_credit.pop()
            # print(total_debit)

            #         print(account_no[i],max_List1,'000')
            # print(total_desc0,len(total_desc0),'desc0')
            # print('--------------------------------------')
            # print(total_desc1,len(total_desc1),'desc1')
            # print('--------------------------------------')
            # print(total_desc2,len(total_desc2),'desc2')
            # print(total_desc3,len(total_desc3),'desc3')
            # print(total_desc4,len(total_desc4),'desc4')

            # print(total_date_list,'d',len(total_date_list))
            # print('--------------------------------------')
            # print(total_debit,len(total_debit),'debit')
            # print('--------------------------------------')
            # print('--------------------------------------')
            # print(total_bal,len(total_bal),'total_bf')
            # print(total_credit,len(total_credit),'credit')

    pdf = FPDF()

    #-------LEFT OMR data----------------#
    a1,b1,c1,d1 = [2,100,8,100] # 1st line
    a10,b10,c10,d10 = [2,118,8,118]  # 10 th line
    dgr1,dgr2,dgr3,dgr4=[2,102,8,102] #  DGR line
    abs,bbs,cbs,dbs = [2,104,8,104] #  bs1 line
    agm,bgm,cgm,dgm=[2,110,8,110]   # gm1 line
    pr1,pr2,pr3,pr4=[2,116,8,116]   #parity line
    parity_left_count=0
    parity_right_count=0
    #------------Right OMR data-----------------------
    rbs1,rbs2,rbs3,rbs4 = [200,220,206,220] # BM right omr 1st line
    reoc1,reoc2,reoc3,reoc4 = [200,218,206,218] # righ 2nd line EOC like DGR
    bsr1,bsr2,bsr3,bsr4 = [200,214,206,214] # right third BS1
    prr1,prr2,prr3,prr4=[200,208,206,208]   # right parity line

    
    mm=0
    new_code=1
    with open(path11) as file:
        c=-1
        date_start_line_number=[]
        for line in file:
            c=c+1
            line=line.strip()
            if line.startswith('DATE    DESCRIPTION'):
                date_start_line_number.append(c)
        date_start_line_number.append(total_lines)

    # print(total_credit,'11')
    postcode_received = STD_GROUP_CODE(postcode)
    df=pd.DataFrame()
    # df['a1']=a1
    df['client_name']=client_name
    df['below_client']=below_client
    df['statement_date']=statement_date
    df['add1']=add1
    df['add2']=add2
    df['add3']=add3
    df['add4']=add4
    df['account_no']=account_no
    df['branch']=branch
    df['barcode_below']=barcode_below
    df['bal_bf']=bal_bf
    df['total_date_list']=total_date_list
    df['total_desc0']=total_desc0
    df['total_desc1']=total_desc1
    df['total_desc2']=total_desc2
    df['total_desc3']=total_desc3
    df['total_desc4']=total_desc4

    df['msg_total_credit'] = msg_total_credit
    df['msg_total_debit'] = msg_total_debit 
    df['monthly_avg'] = monthly_avg
    df['total_debit'] = total_debit
    df['total_credit'] = total_credit
    df['total_bal'] = total_bal
    df['Page_no']=page_no
    result  = pd.merge(df1,df,how = 'inner', on = ['account_no'])
    email = result['email_id']
    email_account_no = list(result['account_no'])
    # print(email_account_no,'--------------------------------11')
    Total_email_accounts=len(email_account_no)
    Total_non_email_account=total_accounts-Total_email_accounts
    print('Total Accounts email BICASA-->',Total_email_accounts)
    print('Total Accounts without email BICASA-->',Total_non_email_account)

    df['Group_Code'] = postcode_received
    grp_code=list(df['Group_Code'])
    grp_code1=[]
    for i in grp_code:
        i=int(i)
        grp_code1.append(i+100)
    df['Group_Code1']=grp_code1

    df['counter']=counter_col
    df = df.sort_values(['Group_Code1','counter'],ignore_index=True,ascending=[True,True])
    # uniqueValues = (df['account_no'].append(df['account_no'])).unique()
    # print(uniqueValues,len(uniqueValues))
    final_df = df
    # print(final_df['msg_total_debit'])
    for i in final_df['total_date_list']:
        len_list.append(len(i))
    max_List1=max(len_list)
    for item in final_df['total_date_list']:               
        while len(item) < max_List1:   
            item.append('Test')
    for item in total_debit:               
        while len(item) < max_List1:   
            item.append('Test')
    for item in total_bal:               
        while len(item) < max_List1:   
            item.append('Test')
    for item in total_credit:               
        while len(item) < max_List1:   
            item.append('Test')

    print('Most Transactions of one client are --', max_List1)


    #-------------------------------------------------
    c=1
    c = c + 1
    unique_acc=final_df['account_no'][0]
    unique_acc_number_list=final_df['account_no'][0]
    unique_acc_list=[0]
    unique_acc_count=0
    unique_acc_list_for_message_without=[]
    unique_acc_list_for_message_with=[]
    
    total_repeating_non_email_acc=[]
    for i in range(0,total_accounts):
        if final_df['account_no'][i] not in email_account_no: 
            total_repeating_non_email_acc.append(i)
    # print('total_repeating_non_email_acc',total_repeating_non_email_acc)
    kt=0
# -------------------for non email account message list-------------
    unique_acc=final_df['account_no'][0]
    for i in range(0,total_accounts):
        if final_df['account_no'][i] not in email_account_no: 
            if unique_acc==final_df['account_no'][i]:
                kt=i                #kt=34
            else:
                unique_acc=final_df['account_no'][i]
                if kt not in unique_acc_list_for_message_without:
                    unique_acc_list_for_message_without.append(kt)
                    kt=i
    unique_acc_list_for_message_without.append(total_accounts-1)
    # print(unique_acc_list_for_message_without,len(unique_acc_list_for_message_without))

# --------------------------------------------------------------
    
# ------------------------------------------------------------------
# -------------------for email account message list-------------
    kt=0
    unique_acc=final_df['account_no'][0]
    for i in range(0,total_accounts):
        if final_df['account_no'][i] in email_account_no: 
            if unique_acc==final_df['account_no'][i]:
                kt=i                #kt=34
            else:
                unique_acc=final_df['account_no'][i]
                if kt not in unique_acc_list_for_message_with:
                    unique_acc_list_for_message_with.append(kt)
                    kt=i
    unique_acc_list_for_message_with.append(total_accounts-1)
    # print(unique_acc_list_for_message_with,len(unique_acc_list_for_message_with))


#-----------------------------function with nonemailids------------------------------
    count = 1
    total_page_count = 0
    for i in range(0,total_accounts):
        if final_df['account_no'][i] not in email_account_no: 
            total_page_count += 1 
    # ---------------------------------------------------------------
    acc = []
    total_repeating_non_email_acc=[]
    for i in range(0,total_accounts):
        if final_df['account_no'][i] not in email_account_no:
            acc.append(final_df['account_no'][i])
    
    d = {}
    for i in acc:
        if i in d:
            d[i] += 1
        else:
            d[i] = 1

    total_page_count = []
    for key,val in d.items():
        total_page_count.append(val)
    page_count = 1
    page_count_index = 0
    rrr=0
    ct = 0
    final_line_flag_list_non_email=[]
    final_line_flag_list_email=[]

# -----------------------------
    with open(path11) as file:
        c=-1
        account_start_line_number=[]
        for line in file:
            c=c+1
            line=line.strip()
            if line.startswith('.') and len(line)==1:
                account_start_line_number.append(c)
        account_start_line_number.append(total_lines)
        # print(account_start_line_number,len(account_start_line_number))

        message_list_printing=[]
        message_list_printing_acc_no=[]

        message_list_digital=[]
        message_list_digital_acc_no=[]
        message=[]
        for i in range(0,len(account_start_line_number)-1):
            a=account_start_line_number[i]
            b=account_start_line_number[i+1]
            # print(a,b)
            for j in range(a,b):
                if j<b-1:
                    line1 = lines[j+1]
                    line1 = line1.strip()
                    if j+8<b-1:
                        line88=lines[j+7].strip()
                        if 'ACCOUNT NO' in line88:
                            kk=line88[-14:]
                            if kk not in email_account_no:
                                for j in range(a,b):
                                    line1 = lines[j]
                                    line1 = line1.strip()
                                    if 'MESSAGE' in line1:
                                        message_list_printing_acc_no.append(kk)
                                        line2=lines[j+1].strip()
                                        if line2.startswith('.') and len(line2)==1:                        
                                            message.append('TEST')
                                        else:
                                            message.append(line2)
                                            if j+2>=total_lines:
                                                message_list_printing.append(message)
                                                message=[]
                                                break
                                            line3=lines[j+2].strip()
                                            if '.' in line3 and len(line3)==1:
                                                message.append('TEST')
                                            else:
                                                message.append(line3)
                                                line4=lines[j+3].strip()
                                                if '.' in line4 and len(line4)==1:
                                                    message.append('TEST')
                                                else:
                                                    message.append(line4)
                                                    line5=lines[j+4].strip()
                                                    if '.' in line5 and len(line5)==1:
                                                        message.append('TEST')
                                                    else:
                                                        message.append(line5)
                                                        line6=lines[j+5].strip()
                                                        if '.' in line6 and len(line6)==1:
                                                            message.append('TEST')
                                                        else:
                                                            message.append(line6)
                                                            line7=lines[j+6].strip()
                                                            if '.' in line7 and len(line7)==1:
                                                                message.append('TEST')
                                                            else:
                                                                message.append(line7)
                                                                line8=lines[j+7].strip()
                                                                if '.' in line8 and len(line8)==1:
                                                                    message.append('TEST')
                                                                else:
                                                                    message.append(line8)
                                                                    line9=lines[j+8].strip()
                                                                    if '.' in line9 and len(line9)==1:
                                                                        message.append('TEST')
                                                                    else:
                                                                        message.append(line9)
                                                                        line10=lines[j+9].strip()
                                                                        if '.' in line10 and len(line10)==1:
                                                                            message.append('TEST')
                                                                        else:
                                                                            message.append(line10)
                                                                            line11=lines[j+10].strip()
                                                                            if '.' in line11 and len(line11)==1:
                                                                                message.append('TEST')
                                                                            else:
                                                                                message.append(line11)
                                        message_list_printing.append(message)
                                        message=[]

                            else:
                                for j in range(a,b):
                                    line1 = lines[j]
                                    line1 = line1.strip()
                                    if 'MESSAGE' in line1:
                                        message_list_digital_acc_no.append(kk)
                                        line2=lines[j+1].strip()
                                        if line2.startswith('.') and len(line2)==1:                        
                                            message.append('TEST')
                                        else:
                                            message.append(line2)
                                            if j+2>=total_lines:
                                                message_list_digital.append(message)
                                                message=[]
                                                break
                                            line3=lines[j+2].strip()
                                            if '.' in line3 and len(line3)==1:
                                                message.append('TEST')
                                            else:
                                                message.append(line3)
                                                line4=lines[j+3].strip()
                                                if '.' in line4 and len(line4)==1:
                                                    message.append('TEST')
                                                else:
                                                    message.append(line4)
                                                    line5=lines[j+4].strip()
                                                    if '.' in line5 and len(line5)==1:
                                                        message.append('TEST')
                                                    else:
                                                        message.append(line5)
                                                        line6=lines[j+5].strip()
                                                        if '.' in line6 and len(line6)==1:
                                                            message.append('TEST')
                                                        else:
                                                            message.append(line6)
                                                            line7=lines[j+6].strip()
                                                            if '.' in line7 and len(line7)==1:
                                                                message.append('TEST')
                                                            else:
                                                                message.append(line7)
                                                                line8=lines[j+7].strip()
                                                                if '.' in line8 and len(line8)==1:
                                                                    message.append('TEST')
                                                                else:
                                                                    message.append(line8)
                                                                    line9=lines[j+8].strip()
                                                                    if '.' in line9 and len(line9)==1:
                                                                        message.append('TEST')
                                                                    else:
                                                                        message.append(line9)
                                                                        line10=lines[j+9].strip()
                                                                        if '.' in line10 and len(line10)==1:
                                                                            message.append('TEST')
                                                                        else:
                                                                            message.append(line10)
                                                                            line11=lines[j+10].strip()
                                                                            if '.' in line11 and len(line11)==1:
                                                                                message.append('TEST')
                                                                            else:
                                                                                message.append(line11)
                                        message_list_digital.append(message)
                                        message=[]

        max_size_message_printing=0
        for i in message_list_printing:
            if len(i)>max_size_message_printing:
                max_size_message_printing=len(i)
        for item in message_list_printing:               
            while len(item) < max_size_message_printing:   
                item.append('TEST')# print(message,len(message))
        print(max_size_message_printing,'------------')
        final_message_list_printing=dict(zip(message_list_printing_acc_no,message_list_printing))
        # print(final_message_list_printing,len(final_message_list_printing))

        max_size_message=0
        for i in message_list_digital:
            if len(i)>max_size_message:
                max_size_message=len(i)
        for item in message_list_digital:               
            while len(item) < max_size_message:   
                item.append('TEST')# print(message,len(message))
        final_message_list_digital=dict(zip(message_list_digital_acc_no,message_list_digital))
        # print(final_message_list_digital,len(final_message_list_digital))

    barcode_counter = 1
    lll=-1
    for i in range(0,total_accounts):
        if final_df['account_no'][i] not in email_account_no: 
            # print(final_df['Group_Code1'][i],final_df['Group_Code'][i],final_df['account_no'][i],final_df['Page_no'][i])
            if 'TEST' in final_df['bal_bf'][i]:
#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                line_flag_list=[]
                for j in range(0,max_List1):
                    if final_df['total_date_list'][i][j] != 'Test':
                        if final_df['total_desc3'][i][j]!='desc_TEST':
                            line_flag_list.append('4line')
                        elif final_df['total_desc2'][i][j]!='desc_TEST':
                            line_flag_list.append('3line')
                        elif final_df['total_desc1'][i][j]!='desc_TEST':
                            line_flag_list.append('2line')
                        elif final_df['total_desc0'][i][j]!='desc_TEST':
                            line_flag_list.append('1line')
                final_line_flag_list_non_email.append(line_flag_list)
            if 'TEST' not in final_df['bal_bf'][i]:
#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                line_flag_list=[]
                for j in range(0,max_List1):
                    if final_df['total_date_list'][i][j] != 'Test':
                        if final_df['total_desc3'][i][j]!='desc_TEST':
                            line_flag_list.append('4line')
                        elif final_df['total_desc2'][i][j]!='desc_TEST':
                            line_flag_list.append('3line')
                        elif final_df['total_desc1'][i][j]!='desc_TEST':
                            line_flag_list.append('2line')
                        elif final_df['total_desc0'][i][j]!='desc_TEST':
                            line_flag_list.append('1line')
                final_line_flag_list_non_email.append(line_flag_list)
    # print(final_line_flag_list_non_email,len(final_line_flag_list_non_email))
    
    for i in range(0,total_accounts):    
        if final_df['account_no'][i] in email_account_no: 
            if 'TEST' in final_df['bal_bf'][i]:
#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                line_flag_list=[]
                for j in range(0,max_List1):
                    if final_df['total_date_list'][i][j] != 'Test':
                        if final_df['total_desc3'][i][j]!='desc_TEST':
                            line_flag_list.append('4line')
                        elif final_df['total_desc2'][i][j]!='desc_TEST':
                            line_flag_list.append('3line')
                        elif final_df['total_desc1'][i][j]!='desc_TEST':
                            line_flag_list.append('2line')
                        elif final_df['total_desc0'][i][j]!='desc_TEST':
                            line_flag_list.append('1line')
                final_line_flag_list_email.append(line_flag_list)
            if 'TEST' not in final_df['bal_bf'][i]:
#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                line_flag_list=[]
                for j in range(0,max_List1):
                    if final_df['total_date_list'][i][j] != 'Test':
                        if final_df['total_desc3'][i][j]!='desc_TEST':
                            line_flag_list.append('4line')
                        elif final_df['total_desc2'][i][j]!='desc_TEST':
                            line_flag_list.append('3line')
                        elif final_df['total_desc1'][i][j]!='desc_TEST':
                            line_flag_list.append('2line')
                        elif final_df['total_desc0'][i][j]!='desc_TEST':
                            line_flag_list.append('1line')
                final_line_flag_list_email.append(line_flag_list)
    # print(final_line_flag_list_email,len(final_line_flag_list_email))


    for i in range(0,total_accounts):
        if final_df['account_no'][i] not in email_account_no: 
            pdf.add_page()
            ct = int(ct)
            ct += 1

            pdf.image(r'E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\Left_logo.jpg',10,15,60)
            pdf.image(r'E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\Right_logo.jpg',160,15,33)
            pdf.set_auto_page_break(auto=False)
        
        # ------------------BARCODE--------------------
            if final_df['account_no'][i]!=unique_acc:
                now = datetime.now() 
                date = now.strftime("%m%d%y")
                barcode_no = '1103' +  str(date) +  str(final_df['account_no'][i]) 
                number = barcode_no
                my_code = Code39(number, writer=ImageWriter())
                my_code.save(r"E:\IWOC\EXTRA\bicasa\%s" %(new_code))
                from PIL import Image

                img = Image.open(r"E:\IWOC\EXTRA\bicasa\%s.png" %(new_code))
                imgCropped = img.crop(box=(4,8,960,200))
                imgCropped.save(r"E:\IWOC\EXTRA\bicasa\%s.png" %(new_code))
                pdf.image(r"E:\IWOC\EXTRA\bicasa\%s.png" %(new_code),13,31,45)
                pdf.set_font('Arial','',7)
                barcode_counter  = str(barcode_counter)
                if len(barcode_counter) == 1:
                    count_value = '00000' + str(barcode_counter)
                if len(barcode_counter) == 2:
                    count_value = '0000' + str(barcode_counter)
                if len(barcode_counter) == 3:
                    count_value = '000' + str(barcode_counter)
                if len(barcode_counter) == 4:
                    count_value = '00' + str(barcode_counter)
                if len(barcode_counter) == 5:
                    count_value = '0' + str(barcode_counter)
                if len(barcode_counter) == 6:
                    count_value =  str(count)
                pdf.cell(3)
                pdf.cell(10,62,txt = '(' + final_df['Group_Code'][i]  + ')' +'-'+ count_value +'-' + final_df['barcode_below'][i] )
                barcode_counter = int(barcode_counter)
                barcode_counter += 1
            
                pdf.ln(0.01)
                new_code+=1
                unique_acc=final_df['account_no'][i]
                # print(unique_acc)

        # --------------------------------
            pdf.set_font('Arial', '', 9)
            pdf.cell(140)
            count  = str(count)
            if len(count) == 1:
                count_value = '00000' + str(count)
            if len(count) == 2:
                count_value = '0000' + str(count)
            if len(count) == 3:
                count_value = '000' + str(count)
            if len(count) == 4:
                count_value = '00' + str(count)
            if len(count) == 5:
                count_value = '0' + str(count)
            if len(count) == 6:
                count_value =  str(count)
            pdf.cell(10,-5,txt = str(count_value) + '/' + files1)
            # pdf.cell(10,-5,txt = str(count_value) + '/' + 'files1')

            count = int(count)
            count += 1

            pdf.cell(-146)
            pdf.cell(10,70,txt = final_df['client_name'][i])
        
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,72,txt = final_df['below_client'][i])

            pdf.set_font('Arial','B',9)
            pdf.cell(90)
            pdf.cell(30,70,txt = 'TARIKH PENYATA')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,74,txt =final_df['add1'][i])

            pdf.set_font('Arial','I',9)
            pdf.cell(90)
            pdf.cell(30,74,txt = 'STATEMENT DATE')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,76,txt = final_df['add2'][i])

            pdf.set_font('Arial','B',9)
            pdf.cell(90)
            pdf.cell(30,78,txt = 'HALAMAN')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,78,txt = final_df['add3'][i])

            pdf.set_font('Arial','I',9)
            pdf.cell(90)
            pdf.cell(30,82,txt = 'PAGE')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,80,txt = final_df['add4'][i])

            pdf.set_font('Arial','B',9)
            pdf.cell(90)
            pdf.cell(10,86,txt = 'NOMBOR AKAUN')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            # pdf.cell(10,82,txt = final_df[add4[i]])

            pdf.set_font('Arial','I',9)
            pdf.cell(100)
            pdf.cell(10,90,txt = 'ACCOUNT NO  ')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            # pdf.cell(10,84,txt = '58200 Kuala Lumpur')

            pdf.set_font('Arial','B',9)
            pdf.cell(100)
            pdf.cell(10,94,txt = 'CAWANGAN')

            pdf.set_font('Arial','I',9)
            pdf.cell(-10)
            pdf.cell(10,102,txt = 'BRANCH ')

            #-----------------------------------------INSERTING VALUES OF NAME,ADDRESS,ACC,BRANCH ON 2 PAGE-----------

            pdf.set_font('Arial', 'B', 9)
            pdf.cell(50)
            pdf.cell(10,55,txt = final_df['statement_date'][i])

            pdf.ln(2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(160)
            pdf.cell(10,66,txt = str(page_count))
            pdf.cell(-3)
            pdf.cell(10,66,txt = 'of')
            pdf.cell(-5)
            if total_page_count[page_count_index] != page_count: 
                pdf.cell(10,66,txt = str(total_page_count[page_count_index]))
                page_count += 1
            else:
                pdf.cell(10,66,txt = str(total_page_count[page_count_index]))
                page_count_index += 1
                page_count = 1

            pdf.ln(2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(157)
            pdf.cell(10,78,txt = final_df['account_no'][i])

            pdf.ln(2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(142)
            pdf.cell(10,90,txt = final_df['branch'][i])

            #-----------------------INSERTING ACCOUNT STATEMENT AND SAVING ACCOUNT ON 2 PAGE----------------------
            pdf.ln(1)
            pdf.set_font('Arial','B',9)
            pdf.cell(4)
            pdf.cell(10,100,txt = 'PENYATA AKAUN / ')

            pdf.set_font('Arial','BI',9)
            pdf.cell(20)
            pdf.cell(10,100,txt = 'ACCOUNT STATEMENT ')

            pdf.ln(1)
            pdf.set_font('Arial','',9)
            pdf.cell(4)
            pdf.cell(10,106,txt = 'AKAUN SIMPANAN / ')

            pdf.set_font('Arial','',9)
            pdf.cell(20)
            pdf.cell(10,106,txt = ' SAVINGS ACCOUNT')

            pdf.ln(1)
            pdf.set_font('Arial','',8)
            pdf.cell(4)
            pdf.cell(10,112,txt = '(Dilindungi oleh PIDM setakat RM250,000 bagi setiap pendeposit. /')

            pdf.set_font('Arial','I',8)
            pdf.ln(1)
            pdf.cell(5)
            pdf.cell(10,118,txt = 'Protected by PIDM up to RM250,000 for each depositor.)')

    #-----------------------------------INSERTING 2ND PAGE TABLE-------------------------------------
        #----------Left Omr Insertion-------------------------------------
            pdf.set_line_width(0.8)
            pdf.line(a1,b1,c1,d1)               #1st  Omr Line
            pdf.line(a10,b10,c10,d10)           #10th Omr Line
            if (i)%4==0:
                # print(i,'DGR')
                pdf.line(dgr1,dgr2,dgr3,dgr4)
                parity_left_count+=1
            if (i-1)%4==0:
                # print(i,'BS1')
                pdf.line(abs,bbs,cbs,dbs )
                parity_left_count+=1
            if (i-2)%4==0:
                # print(i,'BS2')
                pdf.line(abs,bbs+2,cbs,dbs+2 )
                parity_left_count+=1
            if (i+1)%4==0 and i-2!=0:
                # print(i,'BS1 AND BS2')
                pdf.line(abs,bbs,cbs,dbs )
                pdf.line(abs,bbs+2,cbs,dbs+2 )
                parity_left_count+=2
            # print(account_no[i])
            if parity_left_count%2!=0:
                pdf.line(pr1,pr2,pr3,pr4)
            if unique_acc_count!=len(unique_acc_list):
                if i==unique_acc_list[unique_acc_count]:
                    # print(i,'----')
                    index=unique_acc_list.index(unique_acc_list[unique_acc_count])
                    unique_acc_count+=1
                    if (index-1)%4==0:
                        pdf.line(agm,bgm,cgm,dgm)
                        # print(index,'GM1')
                        parity_left_count+=1
                    if (index-2)%4==0:
                        pdf.line(agm,bgm+2,cgm,dgm+2)
                        parity_left_count+=1
                        # print(index,'GM2')
                    if (index+1)%4==0 and i-2!=0:
                        pdf.line(agm,bgm,cgm,dgm)
                        pdf.line(agm,bgm+2,cgm,dgm+2)
                        parity_left_count+=2
                        # print(index,'GM1 & GM2')
            # print(i,parity_left_count,'---')
            parity_left_count=0
            parity_right_count=0
        #------------------------right omr insertion--------------------

            pdf.line(rbs1,rbs2,rbs3,rbs4)               #1st  right Omr Line
            if i==0 or i%5==0:
                pdf.line(reoc1,reoc2,reoc3,reoc4)       #simillar to dgr
                # print(i,'PGR')
                parity_right_count+=1
                
            if (i-1)%5==0:
                pdf.line(bsr1,bsr2,bsr3,bsr4)           #bs1 line
                # print(i,'BS1')
                parity_right_count+=1
            if (i-2)%5==0:
                pdf.line(bsr1,bsr2-2,bsr3,bsr4-2)           #bs2 line
                # print(i,'BS2')
                parity_right_count+=1
            if (i-3)%5==0:
                pdf.line(bsr1,bsr2,bsr3,bsr4)           #bs1 and bs2 line
                pdf.line(bsr1,bsr2-2,bsr3,bsr4-2)           #bs1 and bs2 line
                # print(i,'BS1__BS2')
                parity_right_count+=2
            if (i-4)%5==0:
                pdf.line(bsr1,bsr2-4,bsr3,bsr4-4)           #bs2 line
                # print(i,'BS4')
                parity_right_count+=1
            if parity_right_count%2!=0:
                pdf.line(prr1,prr2,prr3,prr4)               # parity line right
                # print(i,'parity')
        # ---------------------------------------------------------------

            if 'TEST' in final_df['bal_bf'][i]:
                pdf.ln(1)
                pdf.set_fill_color(r=211,g=211,b=211)
                pdf.set_line_width(0.01)
                pdf.rect(x = 15 , y = 100 , w = 180 , h = 15, style = 'DF')

                pdf.set_line_width(0.3)
                pdf.line(40,100, 40, 115)                      # first  | line
                pdf.line(110,100, 110, 115)                    # second | line
                pdf.line(140,100, 140, 115)                    # third | line
                pdf.line(170,100,170,115)                      # fourth | line

                #-------------------------------------INSERTING 2ND PAGE TABLE COLUMN NAMES---------------------------------
                pdf.ln(1)
                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,134,txt = 'TARIKH')
                pdf.cell(6)
                pdf.cell(30,134,txt = 'KETERANGAN')
                pdf.cell(30)
                pdf.cell(30,134,txt = 'DEBIT')
                pdf.cell(4)
                pdf.cell(25,134,txt = 'KREDIT')
                pdf.cell(1)
                pdf.cell(25,134,txt = 'BAKI')


                pdf.ln(4)

                pdf.set_font('Arial', 'I', 9)
                pdf.cell(10)
                pdf.cell(30,136,txt = 'DATE')
                pdf.cell(6)
                pdf.cell(30,136,txt = 'DESCRIPTION')
                pdf.cell(30)
                pdf.cell(25,136,txt = 'DEBIT')
                pdf.cell(8)
                pdf.cell(25,136,txt = 'CREDIT')
                pdf.cell(1)
                pdf.cell(25,136,txt = 'BALANCE')

                pdf.ln(4)

                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,138,txt = '')
                pdf.cell(6)
                pdf.cell(30,138,txt = '')
                pdf.cell(35)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(5)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(1)
                pdf.cell(25,138,txt = 'RM')
                pdf.set_line_width(0.01)
                pdf.ln(2)

#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                h = 138
                p2=h
                cc=1
                dd=115

                m1=h
                m2=h
                flag11=False
                line_flag = 'a'
                start=0
                end=0

                lll+=1
                # print(lll,account_no[i],'no bal bf')
                if final_line_flag_list_non_email[lll]!=True:
                    for j in range(0,max_List1):
                        if final_df['total_date_list'][i][j] != 'Test':
                            # print(final_df['total_date_list'][i][j])
                            pdf.set_line_width(0.005)
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc2'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc1'][i][j]!='desc_TEST':
                                h=h+4
                            elif final_df['total_desc0'][i][j]!='desc_TEST':
                                h=h+5
                                max_size='s1'            
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(10)
                            pdf.cell(10,h,txt = final_df['total_date_list'][i][j])

                            pdf.cell(10)
                            pdf.cell(10,h,txt = str(final_df['total_desc0'][i][j]))

                            pdf.cell(80)
                            pdf.cell(10,h,txt = final_df['total_debit'][i][j])
                            pdf.cell(20)
                            pdf.cell(10,h,txt = final_df['total_credit'][i][j])

                            pdf.cell(4)
                            pdf.cell(10,h,txt = final_df['total_bal'][i][j])
                            pdf.ln(1)

                            if final_df['total_desc0'][i][j]!='desc_TEST':
                                cc=cc+1  
                                line_flag = 'a0'
                            if final_df['total_desc1'][i][j]!='desc_TEST':
                                pdf.ln(1)
                                pdf.cell(30)
                                h=h+4
                                pdf.set_font('Arial','B',8)
                                pdf.cell(20,h,txt =  final_df['total_desc1'][i][j])
                                cc=cc+1  
                                line_flag = 'a1'
                            if final_df['total_desc2'][i][j]!='desc_TEST':                                                
                                pdf.ln(1)                                              
                                pdf.cell(30)
                                h=h+4                    
                                pdf.cell(10,h,txt =  final_df['total_desc2'][i][j])   #  27,29,31,33,35                         
                                cc=cc+1  
                                line_flag = 'a2'
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  final_df['total_desc3'][i][j])  
                                cc=cc+1  
                                line_flag = 'a3'
                            if final_df['total_desc4'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  final_df['total_desc4'][i][j])
                                cc=cc+1  
                                line_flag = 'a3'
                            p=h
                            h=h+0.3
                            pdf.ln(1)       
                            pdf.set_font('Arial','',8)
                            if final_line_flag_list_non_email[lll][j]=='1line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-4,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_non_email[lll][j]=='2line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-3,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_non_email[lll][j]=='3line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-2,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_non_email[lll][j]=='4line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-3.5,txt = '__________________________________________________________________________________________________________________')   
                            # print(lll,j,final_df['account_no'][i])

# 148.6,167.5,214.60000000000014 
                    if h>140 and h<=170 or h==214.60000000000014 :
                        pdf.line(15,110, 15,h-22)                       # 1st vertical line
                        pdf.line(40,110, 40,h-22)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-22)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-22)                       # 4th vertical line
                        pdf.line(170,110, 170,h-22)                       # 5th vertical line
                        pdf.line(195,110, 195,h-22)                       # 6th vertical line
# 216.50000000000006                    
                    elif h>170 and h<=200 or h==216.50000000000006:
                        pdf.line(15,110, 15,h-30)                       # 1st vertical line
                        pdf.line(40,110, 40,h-30)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-30)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-30)                       # 4th vertical line
                        pdf.line(170,110, 170,h-30)                       # 5th vertical line
                        pdf.line(195,110, 195,h-30)                       # 6th vertical line
# 215.4
                    elif h>200 and h<=220:
                        pdf.line(15,110, 15,h-24)                       # 1st vertical line
                        pdf.line(40,110, 40,h-24)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-24)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-24)                       # 4th vertical line
                        pdf.line(170,110, 170,h-24)                       # 5th vertical line
                        pdf.line(195,110, 195,h-24)                       # 6th vertical line
# 222.8
                    elif h>220 and h<=225:
                        pdf.line(15,110, 15,h-18)                       # 1st vertical line
                        pdf.line(40,110, 40,h-18)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-18)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-18)                       # 4th vertical line
                        pdf.line(170,110, 170,h-18)                       # 5th vertical line
                        pdf.line(195,110, 195,h-18)                       # 6th vertical line
# 227.1
                    elif h>225 and h<=228:
                        pdf.line(15,110, 15,h-31)                       # 1st vertical line
                        pdf.line(40,110, 40,h-31)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-31)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-31)                       # 4th vertical line
                        pdf.line(170,110, 170,h-31)                       # 5th vertical line
                        pdf.line(195,110, 195,h-31)                       # 6th vertical line
# 227.1
                    elif h>228 and h<=241:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,110, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-28)                       # 4th vertical line
                        pdf.line(170,110, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line

# 243.1 ,241.10000000000008                                       
                    elif h==243.10000000000008 or h==241.10000000000008 :
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,110, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-32)                       # 4th vertical line
                        pdf.line(170,110, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)                       # 6th vertical line
# 246.3
                    elif h>241 and h<=250:
                        pdf.line(15,110, 15,h-25)                       # 1st vertical line
                        pdf.line(40,110, 40,h-25)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-25)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-25)                       # 4th vertical line
                        pdf.line(170,110, 170,h-25)                       # 5th vertical line
                        pdf.line(195,110, 195,h-25)                       # 6th vertical line
# this for only 252.7000000000001
                    elif h==252.7000000000001:
                        pdf.line(15,110, 15,h-31)                       # 1st vertical line
                        pdf.line(40,110, 40,h-31)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-31)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-31)                       # 4th vertical line
                        pdf.line(170,110, 170,h-31)                       # 5th vertical line
                        pdf.line(195,110, 195,h-31)                       # 6th vertical line

# this for only 255.0000000000001 
                    elif h==255.0000000000001 :
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,110, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-28)                       # 4th vertical line
                        pdf.line(170,110, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line

# this for only 256.0000000000001 
                    elif h==256.0000000000001 :
                        pdf.line(15,110, 15,h-29)                       # 1st vertical line
                        pdf.line(40,110, 40,h-29)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-29)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-29)                       # 4th vertical line
                        pdf.line(170,110, 170,h-29)                       # 5th vertical line
                        pdf.line(195,110, 195,h-29)                       # 6th vertical line

# this for only 256.0000000000001 and  h==252.4000000000001
                    elif  h==256.4000000000001 or h==252.4000000000001 or h==254.7000000000001:
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,110, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-32)                       # 4th vertical line
                        pdf.line(170,110, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)                       # 6th vertical line

# this is for 257.7000000000001
                    elif h==257.7000000000001 :
                        pdf.line(15,110, 15,h-33)                       # 1st vertical line
                        pdf.line(40,110, 40,h-33)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-33)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-33)                       # 4th vertical line
                        pdf.line(170,110, 170,h-33)                       # 5th vertical line
                        pdf.line(195,110, 195,h-33)                       # 6th vertical line
# this is for 259.4000000000001,260.4000000000001 
                    elif h==259.4000000000001 or h==260.4000000000001 :
                        pdf.line(15,110, 15,h-35)                       # 1st vertical line
                        pdf.line(40,110, 40,h-35)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-35)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-35)                       # 4th vertical line
                        pdf.line(170,110, 170,h-35)                       # 5th vertical line
                        pdf.line(195,110, 195,h-35)                       # 6th vertical line


# 251.3, 256.6, 
                    elif h>250 and h<=260:
                        pdf.line(15,110, 15,h-24)                       # 1st vertical line
                        pdf.line(40,110, 40,h-24)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-24)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-24)                       # 4th vertical line
                        pdf.line(170,110, 170,h-24)                       # 5th vertical line
                        pdf.line(195,110, 195,h-24)                       # 6th vertical line

# 266.1,271.4
                    elif h>260 and h<=272:
                        pdf.line(15,110, 15,h-38)                       # 1st vertical line
                        pdf.line(40,110, 40,h-38)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-38)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-38)                       # 4th vertical line
                        pdf.line(170,110, 170,h-38)                       # 5th vertical line
                        pdf.line(195,110, 195,h-38)                       # 6th vertical line

# 273.7,274,279                   
                    elif h>272 and h<=280:
                        pdf.line(15,110, 15,h-20)                       # 1st vertical line
                        pdf.line(40,110, 40,h-20)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-20)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-20)                       # 4th vertical line
                        pdf.line(170,110, 170,h-20)                       # 5th vertical line
                        pdf.line(195,110, 195,h-20)                       # 6th vertical line
# 289.2,294.5
                    elif h>280 and h<=300:
                        pdf.line(15,110, 15,h-19)                       # 1st vertical line
                        pdf.line(40,110, 40,h-19)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-19)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-19)                       # 4th vertical line
                        pdf.line(170,110, 170,h-19)                       # 5th vertical line
                        pdf.line(195,110, 195,h-19)                       # 6th vertical line
# 307.6
                    elif h>300 and h<=320:
                        pdf.line(15,110, 15,h-13)                       # 1st vertical line
                        pdf.line(40,110, 40,h-13)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-13)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-13)                       # 4th vertical line
                        pdf.line(170,110, 170,h-13)                       # 5th vertical line
                        pdf.line(195,110, 195,h-13)                       # 6th vertical line
                    elif h>320 and h<=380:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,110, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-37)                       # 4th vertical line
                        pdf.line(170,110, 170,h-37)                       # 5th vertical line
                        pdf.line(195,110, 195,h-37)                       # 6th vertical line

                    else:
                        pass
                max_page_no=final_df['Page_no'][i]
                if final_df['account_no'][i] in message_list_printing_acc_no:
                        k1=final_df['account_no'][i]
                        k2=final_message_list_printing[k1]
                        if k1 in d.keys():
                            go_page=d[k1]

                if final_df['msg_total_debit'][i] == 'Test' and  final_df['Page_no'][i]==str(go_page):
                    # print('1-----------')
                    pdf.ln(1)
                    pdf.set_line_width(0.01)  
                    pdf.line(15,m1,15,m1-33)                             # 1st vertical line of total table  
                    pdf.line(195,m1,195,m1-33)                           # 2nd vertical line of total tabel

                    pdf.set_font('Arial','B',8)
                    pdf.cell(80)
                    pdf.cell(10,h+30,txt = ' ')

                    pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-15,180)     
                    pdf.line(15,h,15,h+10)                             # 1st vertical line of total table  
                    pdf.line(195,h,195,h+10) 
                    pdf.line(15,h+10,195,h+10) 

                if  final_df['msg_total_debit'][i] != 'Test' and final_df['Page_no'][i]==str(go_page):
                    pdf.ln(1)
                    pdf.set_line_width(0.01)  
                    pdf.set_line_width(0.01)  
                    t=55
                    for k in range(0,max_size_message_printing):
                        if k2[k]!='TEST':
                            if k2[k]=='TEST':
                                break
                            pdf.cell(80)
                            pdf.cell(10,h+t,txt = str(k2[k]))
                            pdf.ln(1)
                            t=t+5
                    if h>140 and h<=150:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h,180) 
                        pdf.line(15,h+16,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+16,195,h-35) 
                        pdf.line(15,h+16,195,h+16)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>150 and h<=160:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h,180) 
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.line(15,h+15,195,h+15)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>160 and h<=165:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-4,180) 
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.line(15,h+15,195,h+15)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+14,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+14,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+19,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+19,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+24,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+24,txt = final_df['monthly_avg'][i])
                    if h == 165.90000000000003 :
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-24,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                        pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+14,195,h-30) 
                        pdf.line(15,h+14,195,h+14)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+11,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+11,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>166 and h<=170:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                        pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+14,195,h-30) 
                        pdf.line(15,h+14,195,h+14)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+11,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+11,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>170 and h<=175:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-10,180) 
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>175 and h<=180:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26.8,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-10,180) 
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>180 and h<=190:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-29.5,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-9,180) 
                        pdf.line(15,h+8,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+8,195,h-35) 
                        pdf.line(15,h+8,195,h+8)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>190 and h<=200:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-30,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-11,180) 
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>200 and h<=210:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>210 and h<=220:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-21,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>220 and h<=228:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-18,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-2,180) 
                        pdf.line(15,h+18,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+18,195,h-35) 
                        pdf.line(15,h+18,195,h+18)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>230 and h<=240:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-2,180) 
                        pdf.line(15,h+18,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+18,195,h-35) 
                        pdf.line(15,h+18,195,h+18)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>240 and h<=250:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-2,180) 
                        pdf.line(15,h+16,15,h-45)                             # 1st vertical line of total table  
                        pdf.line(195,h+16,195,h-45) 
                        pdf.line(15,h+16,195,h+16)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>250 and h<=260:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-50,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-30,180) 
                        pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                        pdf.line(195,h-10,195,h-50) 
                        pdf.line(15,h-10,195,h-10)

                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-50,180)
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-30,180) 
                        pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                        pdf.line(195,h-10,195,h-50) 
                        pdf.line(15,h-10,195,h-10)

                    pdf.ln(1)
                    pdf.set_font('Arial','',7.2)
                    pdf.cell(4)
                    pdf.cell(10,h+80,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                    pdf.ln(1)
                    pdf.cell(4)
                    pdf.cell(10,h+85 , txt = 'penyata ini akan dianggap betul. ')
                    pdf.set_font('Arial','I',7.2)
                    pdf.ln(1)
                    pdf.cell(4)
                    pdf.cell(10,h+90,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                    pdf.set_font('Arial','',7.2)
                    pdf.ln(1)
                    pdf.cell(4)
                    pdf.cell(10,h+95,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                    pdf.ln(1)
                    pdf.cell(4)
                    pdf.cell(10,h+100,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                    pdf.ln(1)
                    pdf.cell(4)
                    pdf.cell(10,h+105,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')
                   
                rrr+=1

            if 'TEST' not in final_df['bal_bf'][i]:
                pdf.ln(1)
                pdf.set_fill_color(r=211,g=211,b=211)
                pdf.set_line_width(0.01)
                pdf.rect(x = 15 , y = 100 , w = 180 , h = 15, style = 'DF')

                pdf.set_line_width(0.3)
                pdf.line(40,100, 40, 115)                      # first  | line
                pdf.line(110,100, 110, 115)                    # second | line
                pdf.line(140,100, 140, 115)                    # third | line
                pdf.line(170,100,170,115)                      # fourth | line

                #-------------------------------------INSERTING 2ND PAGE TABLE COLUMN NAMES---------------------------------

                pdf.ln(1)
                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,134,txt = 'TARIKH')
                pdf.cell(6)
                pdf.cell(30,134,txt = 'KETERANGAN')
                pdf.cell(30)
                pdf.cell(30,134,txt = 'DEBIT')
                pdf.cell(4)
                pdf.cell(25,134,txt = 'KREDIT')
                pdf.cell(1)
                pdf.cell(25,134,txt = 'BAKI')


                pdf.ln(4)

                pdf.set_font('Arial', 'I', 9)
                pdf.cell(10)
                pdf.cell(30,136,txt = 'DATE')
                pdf.cell(6)
                pdf.cell(30,136,txt = 'DESCRIPTION')
                pdf.cell(30)
                pdf.cell(25,136,txt = 'DEBIT')
                pdf.cell(8)
                pdf.cell(25,136,txt = 'CREDIT')
                pdf.cell(1)
                pdf.cell(25,136,txt = 'BALANCE')

                pdf.ln(4)

                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,138,txt = '')
                pdf.cell(6)
                pdf.cell(30,138,txt = '')
                pdf.cell(35)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(5)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(1)
                pdf.cell(25,138,txt = 'RM')
                pdf.set_line_width(0.01)
                pdf.ln(1)
                pdf.set_font('Arial', 'B', 8)
                pdf.cell(30)
                pdf.cell(10,145,txt='BAL B/F')
                pdf.cell(123)
                pdf.cell(10,145,txt= final_df['bal_bf'][i])
                pdf.line(15,119,195,119)
                pdf.ln(2)


#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                h = 147
                p2=h
                cc=1
                dd=115

                m1=h
                m2=h
                flag11=False
                line_flag = 'a'
                start=0
                end=0
                lll+=1
                # print(total_debit)
                if final_line_flag_list_non_email[lll]!=True:
                    for j in range(0,max_List1):
                        if final_df['total_date_list'][i][j] != 'Test':
                            # print(final_df['total_date_list'][i][j])
                            pdf.set_line_width(0.005)
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc2'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc1'][i][j]!='desc_TEST':
                                h=h+4
                            elif final_df['total_desc0'][i][j]!='desc_TEST':
                                h=h+5
                                max_size='s1'            
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(10)
                            pdf.cell(10,h,txt = final_df['total_date_list'][i][j])

                            pdf.cell(10)
                            pdf.cell(10,h,txt = str(final_df['total_desc0'][i][j]))

                            pdf.cell(80)
                            pdf.cell(10,h,txt = final_df['total_debit'][i][j])
                            pdf.cell(20)
                            pdf.cell(10,h,txt = final_df['total_credit'][i][j])

                            pdf.cell(4)
                            pdf.cell(10,h,txt = final_df['total_bal'][i][j]) 
                            pdf.ln(1)

                            if final_df['total_desc0'][i][j]!='desc_TEST':
                                cc=cc+1  
                                line_flag = 'a0'
                            if final_df['total_desc1'][i][j]!='desc_TEST':
                                pdf.ln(1)
                                pdf.cell(30)
                                h=h+4
                                pdf.set_font('Arial','B',8)
                                pdf.cell(20,h,txt =  str(final_df['total_desc1'][i][j]))
                                cc=cc+1  
                                line_flag = 'a1'
                            if final_df['total_desc2'][i][j]!='desc_TEST':                                                
                                pdf.ln(1)                                              
                                pdf.cell(30)
                                h=h+4                    
                                pdf.cell(10,h,txt =  str(final_df['total_desc2'][i][j]))   #  27,29,31,33,35                         
                                cc=cc+1  
                                line_flag = 'a2'
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  str(final_df['total_desc3'][i][j]))   
                                cc=cc+1  
                                line_flag = 'a3'
                            if final_df['total_desc4'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  str(final_df['total_desc4'][i][j]))   
                                cc=cc+1  
                                line_flag = 'a3'
                            p=h
                            h=h+0.3
                            pdf.ln(1)       
                            pdf.set_font('Arial','',8)
                            if final_line_flag_list_non_email[lll][j]=='1line':
                                pdf.cell(4)
                                pdf.cell(10,h-4,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_non_email[lll][j]=='2line':
                                pdf.cell(4)
                                pdf.cell(10,h-3,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_non_email[lll][j]=='3line':
                                pdf.cell(4)
                                pdf.cell(10,h-2,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_non_email[lll][j]=='4line':
                                pdf.cell(4)
                                pdf.cell(10,h-3.5,txt = '__________________________________________________________________________________________________________________')   

                    # if final_df['account_no'][i]=='14014022004137':
                    #     print(h)
# done 162.9
                    if h>150 and h<=200:
                        pdf.line(15,115, 15,h-26)                       # 1st vertical line
                        pdf.line(40,119, 40,h-26)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-26)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-26)                       # 4th vertical line
                        pdf.line(170,119, 170,h-26)                       # 5th vertical line
                        pdf.line(195,115, 195,h-26)                       # 6th vertical line
# done 205.3
                    elif h>200 and h<=220:
                        pdf.line(15,110, 15,h-23)                       # 1st vertical line
                        pdf.line(40,119, 40,h-23)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-23)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-23)                       # 4th vertical line
                        pdf.line(170,119, 170,h-23)                       # 5th vertical line
                        pdf.line(195,110, 195,h-23)                       # 6th vertical line
# done 220.7
                    elif h>220 and h<=230:
                        pdf.line(15,110, 15,h-27)                       # 1st vertical line
                        pdf.line(40,119, 40,h-27)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-27)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-27)                       # 4th vertical line
                        pdf.line(170,119, 170,h-27)                       # 5th vertical line
                        pdf.line(195,110, 195,h-27)                       # 6th vertical line
# done 235.4
                    elif h>230 and h<=250:
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,119, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-32)                       # 4th vertical line
                        pdf.line(170,119, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)                       # 6th vertical line
# done 260.3
                    elif h>250 and h<=261:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,119, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-28)                       # 4th vertical line
                        pdf.line(170,119, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line
# done 263.1
                    elif h>261 and h<=265:
                        pdf.line(15,110, 15,h-39)                       # 1st vertical line
                        pdf.line(40,119, 40,h-39)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-39)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-39)                       # 4th vertical line
                        pdf.line(170,119, 170,h-39)                       # 5th vertical line
                        pdf.line(195,119, 195,h-39)                       # 6th vertical line
# done 266.9
                    elif h>265 and h<=267:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,119, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-28)                       # 4th vertical line
                        pdf.line(170,119, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line
# done 268.6
                    elif h>267 and h<=269:
                        pdf.line(15,110, 15,h-30)                       # 1st vertical line
                        pdf.line(40,119, 40,h-30)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-30)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-30)                       # 4th vertical line
                        pdf.line(170,119, 170,h-30)                       # 5th vertical line
                        pdf.line(195,110, 195,h-30)                       # 6th vertical line


# done 269.7
                    elif h>269 and h<=275:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,119, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-37)                       # 4th vertical line
                        pdf.line(170,119, 170,h-37)                       # 5th vertical line
                        pdf.line(195,110, 195,h-37)                       # 6th vertical line
# done 
                    elif h>275 and h <=276:
                        pdf.line(15,110, 15,h-42)                       # 1st vertical line
                        pdf.line(40,119, 40,h-42)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-42)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-42)                       # 4th vertical line
                        pdf.line(170,119, 170,h-42)                       # 5th vertical line
                        pdf.line(195,110, 195,h-42)                       # 6th vertical line

# done                    
                    elif h>276 and h<=280:
                        pdf.line(15,110, 15,h-24)                       # 1st vertical line
                        pdf.line(40,119, 40,h-24)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-24)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-24)                       # 4th vertical line
                        pdf.line(170,119, 170,h-24)                       # 5th vertical line
                        pdf.line(195,110, 195,h-24)                       # 6th vertical line
# done 284.3
                    elif h>280 and h<=285:
                        pdf.line(15,110, 15,h-23)                       # 1st vertical line
                        pdf.line(40,119, 40,h-23)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-23)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-23)                       # 4th vertical line
                        pdf.line(170,119, 170,h-23)                       # 5th vertical line
                        pdf.line(195,110, 195,h-23)                       # 6th vertical line

                    elif h>285 and h<=286:
                        pdf.line(15,110, 15,h-47)                       # 1st vertical line
                        pdf.line(40,119, 40,h-47)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-47)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-47)                       # 4th vertical line
                        pdf.line(170,119, 170,h-47)                       # 5th vertical line
                        pdf.line(195,110, 195,h-47)                       # 6th vertical line
# done 292.9
                    elif h>286 and h<=300:
                        pdf.line(15,110, 15,h-23)                       # 1st vertical line
                        pdf.line(40,119, 40,h-23)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-23)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-23)                       # 4th vertical line
                        pdf.line(170,119, 170,h-23)                       # 5th vertical line
                        pdf.line(195,110, 195,h-23)                       # 6th vertical line
# done 
                    elif h>300 and h<=320:
                        pdf.line(15,110, 15,h-16)                       # 1st vertical line
                        pdf.line(40,119, 40,h-16)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-16)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-16)                       # 4th vertical line
                        pdf.line(170,119, 170,h-16)                       # 5th vertical line
                        pdf.line(195,110, 195,h-16)                       # 6th vertical line

                    elif h>320 and h<=380:
                        pdf.line(15,110, 15,h-40)                       # 1st vertical line
                        pdf.line(40,119, 40,h-40)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-40)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-40)                       # 4th vertical line
                        pdf.line(170,119, 170,h-40)                       # 5th vertical line
                        pdf.line(195,119, 195,h-40)                       # 6th vertical line

                    else:
                        pass
                k=0
                # if '13026020058027' in final_df['account_no'][i]:
                    # print(h,'------------=----------------------=balbf')
                if i in unique_acc_list_for_message_without:
                    if msg_total_debit[i] == 'Test' and flag11==False and len_list[i]<1:
                        pdf.ln(1)
                        pdf.set_line_width(0.01)  
                        pdf.line(15,m1,15,m1-33)                             # 1st vertical line of total table  
                        pdf.line(195,m1,195,m1-33)                           # 2nd vertical line of total tabel

                        pdf.set_font('Arial','B',8)
                        pdf.cell(80)
                        pdf.cell(10,h+30,txt = 'NO MESSAGE')

                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-20,180)     
                        pdf.line(15,h,15,h+10)                             # 1st vertical line of total table  
                        pdf.line(195,h,195,h+10) 
                        pdf.line(15,h+10,195,h+10) 
                                                   
                if final_df['account_no'][i] in message_list_printing_acc_no:
                    k1=final_df['account_no'][i]
                    k2=final_message_list_printing[k1]
                    if k1 in d.keys():
                        go_page=d[k1]
                        # print(k1,go_page,'balbf')
                    if k1=='14014021300532':
                        print(type(final_df['Page_no'][i]),type(go_page))
                    if final_df['msg_total_debit'][i] == 'Test' and  final_df['Page_no'][i]==go_page:
                        print('4-----------')
                        
                        pdf.ln(1)
                        pdf.set_line_width(0.01)  
                        pdf.line(15,m1,15,m1-33)                             # 1st vertical line of total table  
                        pdf.line(195,m1,195,m1-33)                           # 2nd vertical line of total tabel

                        pdf.set_font('Arial','B',8)
                        pdf.cell(80)
                        pdf.cell(10,h+30,txt = ' ')

                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-15,180)     
                        pdf.line(15,h,15,h+10)                             # 1st vertical line of total table  
                        pdf.line(195,h,195,h+10) 
                        pdf.line(15,h+10,195,h+10) 
                    if  final_df['msg_total_debit'][i] != 'Test' and final_df['Page_no'][i]==str(go_page):
                
                        pdf.ln(1)
                        pdf.set_line_width(0.01)  
                        pdf.set_line_width(0.01)  
                        t=55
                        for k in range(0,max_size_message_printing):
                            if k2[k]!='TEST':
                                if k2[k]=='TEST':
                                    break
                                pdf.cell(80)
                                pdf.cell(10,h+t,txt = str(k2[k]))
                                pdf.ln(1)
                                t=t+5

# ----------------------sumbox-------------
                        if h>140 and h<=150:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h,180) 
                            pdf.line(15,h+16,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+16,195,h-35) 
                            pdf.line(15,h+16,195,h+16)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                        elif h>150 and h<=160:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h,180) 
                            pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+15,195,h-35) 
                            pdf.line(15,h+15,195,h+15)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                        elif h>160 and h<=165:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-6,180) 
                            pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+15,195,h-35) 
                            pdf.line(15,h+15,195,h+15)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+13,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+13,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+21,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+21,txt = final_df['monthly_avg'][i])
                        elif h>165 and h<=170:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22.5,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                            pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                            pdf.line(195,h+14,195,h-30) 
                            pdf.line(15,h+14,195,h+14)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+23,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+23,txt = final_df['monthly_avg'][i])
                        elif h>170 and h<=175:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-10,180) 
                            pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-30) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                        elif h>175 and h<=180:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-10,180) 
                            pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-30) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                        elif h>180 and h<=190:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-7,180) 
                            pdf.line(15,h+8,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+8,195,h-35) 
                            pdf.line(15,h+8,195,h+8)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+11,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+11,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+16,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+16,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+21,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+21,txt = final_df['monthly_avg'][i])
                        elif h>190 and h<=195:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-8,180) 
                            pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-35) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                        elif h>195 and h<=200:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-8,180) 
                            pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-35) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                        elif h>200 and h<=210:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                            pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-35) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                        elif h>210 and h<=215:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-24,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                            pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-35) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                        elif h>215 and h<=220:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                            pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+10,195,h-35) 
                            pdf.line(15,h+10,195,h+10)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                        elif h>220 and h<=240:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-18,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-2,180) 
                            pdf.line(15,h+18,15,h-35)                             # 1st vertical line of total table  
                            pdf.line(195,h+18,195,h-35) 
                            pdf.line(15,h+18,195,h+18)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                        elif h>240 and h<=250:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-30,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                            pdf.line(15,h+8,15,h-45)                             # 1st vertical line of total table  
                            pdf.line(195,h+8,195,h-45) 
                            pdf.line(15,h+8,195,h+8)
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(20)
                            pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                            pdf.cell(30)
                            pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                            pdf.cell(30)
                            pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                            pdf.ln(1)
                            pdf.cell(20)
                            pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                            pdf.cell(30)
                            pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                        elif h>250 and h<=260:
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-50,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-30,180) 
                            pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                            pdf.line(195,h-10,195,h-50) 
                            pdf.line(15,h-10,195,h-10)

                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-50,180)
                            pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-30,180) 
                            pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                            pdf.line(195,h-10,195,h-50) 
                            pdf.line(15,h-10,195,h-10)
                        


                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+80,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+85 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+90,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+95,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+100,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+105,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')
                    rrr+=1

    #---------------------------filenames -----------------------------
    kc1 = 0
    # kc1 = 1
    pdf_merger = PdfFileMerger()
    date1 = datetime.now()
    d = date1.strftime('%d')
    m = date1.strftime('%m')
    y = date1.strftime('%y')
    y = y[2:4]

    if m == '01':
        month = 'JAN'
    if m == '02':
        month = 'FEB'
    if m == '03':
        month = 'MAR'
    if m == '04':
        month = 'APR'
    if m == '05':
        month = 'MAY'
    if m == '06':
        month = 'JUN'
    if m == '07':
        month = 'JUL'
    if m == '08':
        month = 'AUG'
    if m == '09':
        month = 'SEP'
    if m == '10':
        month = 'OCT'
    if m == '11':
        month = 'NOV'
    if m == '12':
        month = 'DEC'

    filename = os.path.abspath(r"%s\C-%s%s%s_0000%s.pdf"%(kk2,d,month,y,kc1))
    f = filename    
    # print(Total_email_accounts,'NON EMAIL ACCOUNTS PDF CREATED')
    pdf.output(f)    
    kc1 +=  1  
################################################--EMAIL---FUNCTIONALITY--#####################################################
#############################################################################################################
###############################################################################################################
# start of emial_account-----------------##########################--------------
    count = 1
    total_page_count = 0

    acc1 = []
    for i in range(0,total_accounts):
        if final_df['account_no'][i] in email_account_no:
            acc1.append(final_df['account_no'][i])    
    d_email = {}

    for i in acc1:
        if i in d_email:
            d_email[i] += 1
        else:
            d_email[i] = 1
    total_page_count_email = []
    for key,val in d_email.items():
        total_page_count_email.append(val)    
    # print(total_page_count_email,'----1-1-1--1-1-',len(total_page_count_email),'--1')
    
    for i in range(0,total_accounts):
        if final_df['account_no'][i] in email_account_no: 
            a=1
            acc=final_df['account_no'][i]               #this is first email account
            if a==1:
                break
    for i in range(0,total_accounts):
        if final_df['account_no'][i] in email_account_no: 
            a=1
            last_acc=final_df['account_no'][total_accounts-1]               #this is last account no
            if a==1:
                break
    lll=-1
    pdf = FPDF()    
    page_count = 1
    ccc = 0
    page_count_index=0
    email_page_count = 0

    for i in range(0,total_accounts):
        if final_df['account_no'][i] in email_account_no: 
            if acc!=final_df['account_no'][i] :
                ccc += 1
                email_page_count = 1
                pdf_merger = PdfFileMerger()
                date1 = datetime.now()
                d = date1.strftime('%d')
                m = date1.strftime('%m')
                y = date1.strftime('%y')
                y = y[2:4]

                if m == '01':
                    month = 'JAN'
                if m == '02':
                    month = 'FEB'
                if m == '03':
                    month = 'MAR'
                if m == '04':
                    month = 'APR'
                if m == '05':
                    month = 'MAY'
                if m == '06':
                    month = 'JUN'
                if m == '07':
                    month = 'JUL'
                if m == '08':
                    month = 'AUG'
                if m == '09':
                    month = 'SEP'
                if m == '10':
                    month = 'OCT'
                if m == '11':
                    month = 'NOV'
                if m == '12':
                    month = 'DEC'
                filename = os.path.abspath(r"%s\C-%s%s%s_0000%s.pdf"%(kk1,d,month,y,kc1))
                f = filename    
                pdf.output(f)   
                kc1 +=  1  
                pdf=FPDF()
                acc=final_df['account_no'][i]

            pdf.add_page()
            ct = int(ct)
            ct += 1

            pdf.image(r'E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\Left_logo.jpg',10,15,60)
            pdf.image(r'E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\Right_logo.jpg',160,15,33)
            pdf.set_auto_page_break(auto=False)

        # --------------------------------
            pdf.set_font('Arial', '', 9)
            pdf.cell(140)
            count  = str(count)
            if len(count) == 1:
                count_value = '00000' + str(count)
            if len(count) == 2:
                count_value = '0000' + str(count)
            if len(count) == 3:
                count_value = '000' + str(count)
            if len(count) == 4:
                count_value = '00' + str(count)
            if len(count) == 5:
                count_value = '0' + str(count)
            if len(count) == 6:
                count_value =  str(count)
            pdf.cell(10,-5,txt = str(count_value) + '/' + files1)
            # pdf.cell(10,-5,txt = str(count_value) + '/' + 'files1')

            count = int(count)
            count += 1

            pdf.cell(-146)
            pdf.cell(10,70,txt = final_df['client_name'][i])
        
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,72,txt = final_df['below_client'][i])

            pdf.set_font('Arial','B',9)
            pdf.cell(90)
            pdf.cell(30,70,txt = 'TARIKH PENYATA')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,74,txt =final_df['add1'][i])

            pdf.set_font('Arial','I',9)
            pdf.cell(90)
            pdf.cell(30,74,txt = 'STATEMENT DATE')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,76,txt = final_df['add2'][i])

            pdf.set_font('Arial','B',9)
            pdf.cell(90)
            pdf.cell(30,78,txt = 'HALAMAN')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,78,txt = final_df['add3'][i])

            pdf.set_font('Arial','I',9)
            pdf.cell(90)
            pdf.cell(30,82,txt = 'PAGE')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            pdf.cell(10,80,txt = final_df['add4'][i])

            pdf.set_font('Arial','B',9)
            pdf.cell(90)
            pdf.cell(10,86,txt = 'NOMBOR AKAUN')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            # pdf.cell(10,82,txt = final_df[add4[i]])

            pdf.set_font('Arial','I',9)
            pdf.cell(100)
            pdf.cell(10,90,txt = 'ACCOUNT NO  ')

            pdf.set_font('Arial','',9)
            pdf.ln(2)
            pdf.cell(4)
            # pdf.cell(10,84,txt = '58200 Kuala Lumpur')

            pdf.set_font('Arial','B',9)
            pdf.cell(100)
            pdf.cell(10,94,txt = 'CAWANGAN')

            pdf.set_font('Arial','I',9)
            pdf.cell(-10)
            pdf.cell(10,102,txt = 'BRANCH ')

            #-----------------------------------------INSERTING VALUES OF NAME,ADDRESS,ACC,BRANCH ON 2 PAGE-----------

            pdf.set_font('Arial', 'B', 9)
            pdf.cell(50)
            pdf.cell(10,55,txt = final_df['statement_date'][i])
            pdf.ln(2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(160)
            pdf.cell(10,66,txt = str(page_count))
            pdf.cell(-3)
            pdf.cell(10,66,txt = 'of')
            pdf.cell(-5)
            if total_page_count_email[page_count_index] != page_count: 
                pdf.cell(10,66,txt = str(total_page_count_email[page_count_index]))
                page_count += 1
            else:
                pdf.cell(10,66,txt = str(total_page_count_email[page_count_index]))
                page_count_index += 1
                page_count = 1

            pdf.ln(2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(157)
            pdf.cell(10,78,txt = final_df['account_no'][i])

            pdf.ln(2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(142)
            pdf.cell(10,90,txt = final_df['branch'][i])

            #-----------------------INSERTING ACCOUNT STATEMENT AND SAVING ACCOUNT ON 2 PAGE----------------------
            pdf.ln(1)
            pdf.set_font('Arial','B',9)
            pdf.cell(4)
            pdf.cell(10,100,txt = 'PENYATA AKAUN / ')

            pdf.set_font('Arial','BI',9)
            pdf.cell(20)
            pdf.cell(10,100,txt = 'ACCOUNT STATEMENT ')

            pdf.ln(1)
            pdf.set_font('Arial','',9)
            pdf.cell(4)
            pdf.cell(10,106,txt = 'AKAUN SIMPANAN / ')

            pdf.set_font('Arial','',9)
            pdf.cell(20)
            pdf.cell(10,106,txt = ' SAVINGS ACCOUNT')

            pdf.ln(1)
            pdf.set_font('Arial','',8)
            pdf.cell(4)
            pdf.cell(10,112,txt = '(Dilindungi oleh PIDM setakat RM250,000 bagi setiap pendeposit. /')

            pdf.set_font('Arial','I',8)
            pdf.ln(1)
            pdf.cell(5)
            pdf.cell(10,118,txt = 'Protected by PIDM up to RM250,000 for each depositor.)')

    #-----------------------------------INSERTING 2ND PAGE TABLE-------------------------------------
        #----------Left Omr Insertion-------------------------------------
            pdf.set_line_width(0.8)
            pdf.line(a1,b1,c1,d1)               #1st  Omr Line
            pdf.line(a10,b10,c10,d10)           #10th Omr Line
            if (i)%4==0:
                # print(i,'DGR')
                pdf.line(dgr1,dgr2,dgr3,dgr4)
                parity_left_count+=1
            if (i-1)%4==0:
                # print(i,'BS1')
                pdf.line(abs,bbs,cbs,dbs )
                parity_left_count+=1
            if (i-2)%4==0:
                # print(i,'BS2')
                pdf.line(abs,bbs+2,cbs,dbs+2 )
                parity_left_count+=1
            if (i+1)%4==0 and i-2!=0:
                # print(i,'BS1 AND BS2')
                pdf.line(abs,bbs,cbs,dbs )
                pdf.line(abs,bbs+2,cbs,dbs+2 )
                parity_left_count+=2
            # print(account_no[i])
            if parity_left_count%2!=0:
                pdf.line(pr1,pr2,pr3,pr4)
            if unique_acc_count!=len(unique_acc_list):
                if i==unique_acc_list[unique_acc_count]:
                    # print(i,'----')
                    index=unique_acc_list.index(unique_acc_list[unique_acc_count])
                    unique_acc_count+=1
                    if (index-1)%4==0:
                        pdf.line(agm,bgm,cgm,dgm)
                        # print(index,'GM1')
                        parity_left_count+=1
                    if (index-2)%4==0:
                        pdf.line(agm,bgm+2,cgm,dgm+2)
                        parity_left_count+=1
                        # print(index,'GM2')
                    if (index+1)%4==0 and i-2!=0:
                        pdf.line(agm,bgm,cgm,dgm)
                        pdf.line(agm,bgm+2,cgm,dgm+2)
                        parity_left_count+=2
                        # print(index,'GM1 & GM2')
            # print(i,parity_left_count,'---')
            parity_left_count=0
            parity_right_count=0
        #------------------------right omr insertion--------------------

            pdf.line(rbs1,rbs2,rbs3,rbs4)               #1st  right Omr Line
            if i==0 or i%5==0:
                pdf.line(reoc1,reoc2,reoc3,reoc4)       #simillar to dgr
                # print(i,'PGR')
                parity_right_count+=1
                
            if (i-1)%5==0:
                pdf.line(bsr1,bsr2,bsr3,bsr4)           #bs1 line
                # print(i,'BS1')
                parity_right_count+=1
            if (i-2)%5==0:
                pdf.line(bsr1,bsr2-2,bsr3,bsr4-2)           #bs2 line
                # print(i,'BS2')
                parity_right_count+=1
            if (i-3)%5==0:
                pdf.line(bsr1,bsr2,bsr3,bsr4)           #bs1 and bs2 line
                pdf.line(bsr1,bsr2-2,bsr3,bsr4-2)           #bs1 and bs2 line
                # print(i,'BS1__BS2')
                parity_right_count+=2
            if (i-4)%5==0:
                pdf.line(bsr1,bsr2-4,bsr3,bsr4-4)           #bs2 line
                # print(i,'BS4')
                parity_right_count+=1
            if parity_right_count%2!=0:
                pdf.line(prr1,prr2,prr3,prr4)               # parity line right
                # print(i,'parity')
        # ---------------------------------------------------------------

            if 'TEST' in final_df['bal_bf'][i]:

                pdf.ln(1)
                pdf.set_fill_color(r=211,g=211,b=211)
                pdf.set_line_width(0.01)
                pdf.rect(x = 15 , y = 100 , w = 180 , h = 15, style = 'DF')

                pdf.set_line_width(0.3)
                pdf.line(40,100, 40, 115)                      # first  | line
                pdf.line(110,100, 110, 115)                    # second | line
                pdf.line(140,100, 140, 115)                    # third | line
                pdf.line(170,100,170,115)                      # fourth | line

                #-------------------------------------INSERTING 2ND PAGE TABLE COLUMN NAMES---------------------------------
                pdf.ln(1)
                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,134,txt = 'TARIKH')
                pdf.cell(6)
                pdf.cell(30,134,txt = 'KETERANGAN')
                pdf.cell(30)
                pdf.cell(30,134,txt = 'DEBIT')
                pdf.cell(4)
                pdf.cell(25,134,txt = 'KREDIT')
                pdf.cell(1)
                pdf.cell(25,134,txt = 'BAKI')


                pdf.ln(4)

                pdf.set_font('Arial', 'I', 9)
                pdf.cell(10)
                pdf.cell(30,136,txt = 'DATE')
                pdf.cell(6)
                pdf.cell(30,136,txt = 'DESCRIPTION')
                pdf.cell(30)
                pdf.cell(25,136,txt = 'DEBIT')
                pdf.cell(8)
                pdf.cell(25,136,txt = 'CREDIT')
                pdf.cell(1)
                pdf.cell(25,136,txt = 'BALANCE')

                pdf.ln(4)

                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,138,txt = '')
                pdf.cell(6)
                pdf.cell(30,138,txt = '')
                pdf.cell(35)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(5)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(1)
                pdf.cell(25,138,txt = 'RM')
                pdf.set_line_width(0.01)
                pdf.ln(2)

#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                h = 138
                p2=h
                cc=1
                dd=115

                m1=h
                m2=h
                flag11=False
                line_flag = 'a'
                start=0
                end=0

                lll+=1
                # print(lll,account_no[i],'no bal bf')
                if final_line_flag_list_email[lll]!=True:
                    for j in range(0,max_List1):
                        if final_df['total_date_list'][i][j] != 'Test':
                            # print(final_df['total_date_list'][i][j])
                            pdf.set_line_width(0.005)
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc2'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc1'][i][j]!='desc_TEST':
                                h=h+4
                            elif final_df['total_desc0'][i][j]!='desc_TEST':
                                h=h+5
                                max_size='s1'            
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(10)
                            pdf.cell(10,h,txt = final_df['total_date_list'][i][j])

                            pdf.cell(10)
                            pdf.cell(10,h,txt = str(final_df['total_desc0'][i][j]))

                            pdf.cell(80)
                            pdf.cell(10,h,txt = final_df['total_debit'][i][j])
                            pdf.cell(20)
                            pdf.cell(10,h,txt = final_df['total_credit'][i][j])

                            pdf.cell(4)
                            pdf.cell(10,h,txt = final_df['total_bal'][i][j])
                            pdf.ln(1)

                            if final_df['total_desc0'][i][j]!='desc_TEST':
                                cc=cc+1  
                                line_flag = 'a0'
                            if final_df['total_desc1'][i][j]!='desc_TEST':
                                pdf.ln(1)
                                pdf.cell(30)
                                h=h+4
                                pdf.set_font('Arial','B',8)
                                pdf.cell(20,h,txt =  final_df['total_desc1'][i][j])
                                cc=cc+1  
                                line_flag = 'a1'
                            if final_df['total_desc2'][i][j]!='desc_TEST':                                                
                                pdf.ln(1)                                              
                                pdf.cell(30)
                                h=h+4                    
                                pdf.cell(10,h,txt =  final_df['total_desc2'][i][j])   #  27,29,31,33,35                         
                                cc=cc+1  
                                line_flag = 'a2'
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  final_df['total_desc3'][i][j])  
                                cc=cc+1  
                                line_flag = 'a3'
                            if final_df['total_desc4'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  final_df['total_desc4'][i][j])
                                cc=cc+1  
                                line_flag = 'a3'
                            p=h
                            h=h+0.3
                            pdf.ln(1)       
                            pdf.set_font('Arial','',8)
                            if final_line_flag_list_email[lll][j]=='1line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-4,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_email[lll][j]=='2line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-3,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_email[lll][j]=='3line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-2,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_email[lll][j]=='4line':
                                pdf.cell(4.5)
                                pdf.cell(10,h-3.5,txt = '__________________________________________________________________________________________________________________')   
                            # print(lll,j,final_df['account_no'][i])


                    
# 148.6,167.5,214.60000000000014 
                    if h>140 and h<=170 or h==214.60000000000014 :
                        pdf.line(15,110, 15,h-22)                       # 1st vertical line
                        pdf.line(40,110, 40,h-22)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-22)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-22)                       # 4th vertical line
                        pdf.line(170,110, 170,h-22)                       # 5th vertical line
                        pdf.line(195,110, 195,h-22)                       # 6th vertical line
# 216.50000000000006                    
                    elif h>170 and h<=200 or h==216.50000000000006:
                        pdf.line(15,110, 15,h-30)                       # 1st vertical line
                        pdf.line(40,110, 40,h-30)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-30)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-30)                       # 4th vertical line
                        pdf.line(170,110, 170,h-30)                       # 5th vertical line
                        pdf.line(195,110, 195,h-30)                       # 6th vertical line

                    elif h == 217.50000000000006:
                        pdf.line(15,110, 15,h-31)                       # 1st vertical line
                        pdf.line(40,110, 40,h-31)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-31)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-31)                       # 4th vertical line
                        pdf.line(170,110, 170,h-31)                       # 5th vertical line
                        pdf.line(195,110, 195,h-31)

# 215.4
                    elif h>200 and h<=220:
                        pdf.line(15,110, 15,h-24)                       # 1st vertical line
                        pdf.line(40,110, 40,h-24)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-24)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-24)                       # 4th vertical line
                        pdf.line(170,110, 170,h-24)                       # 5th vertical line
                        pdf.line(195,110, 195,h-24)                       # 6th vertical line
                    
                    elif h == 221.50000000000006:
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,110, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-32)                       # 4th vertical line
                        pdf.line(170,110, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)  

                    elif h == 223.80000000000007:
                        pdf.line(15,110, 15,h-30)                       # 1st vertical line
                        pdf.line(40,110, 40,h-30)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-30)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-30)                       # 4th vertical line
                        pdf.line(170,110, 170,h-30)                       # 5th vertical line
                        pdf.line(195,110, 195,h-30)

# 222.8
                    elif h == 224.4000000000001:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,110, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-28)                       # 4th vertical line
                        pdf.line(170,110, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)
       
                    elif h>220 and h<=225:
                        pdf.line(15,110, 15,h-18)                       # 1st vertical line
                        pdf.line(40,110, 40,h-18)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-18)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-18)                       # 4th vertical line
                        pdf.line(170,110, 170,h-18)                       # 5th vertical line
                        pdf.line(195,110, 195,h-18)                       # 6th vertical line
                    
                    elif h == 226.4000000000001:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,110, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-28)                       # 4th vertical line
                        pdf.line(170,110, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)

# 227.1
                    elif h>225 and h<=228:
                        pdf.line(15,110, 15,h-31)                       # 1st vertical line
                        pdf.line(40,110, 40,h-31)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-31)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-31)                       # 4th vertical line
                        pdf.line(170,110, 170,h-31)                       # 5th vertical line
                        pdf.line(195,110, 195,h-31)                       # 6th vertical line

#234.80000000000007
                    elif h == 234.80000000000007:
                        pdf.line(15,110, 15,h-34)                       # 1st vertical line
                        pdf.line(40,110, 40,h-34)                      # 2nd vertical line
                        pdf.line(110,110, 110,h-34)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-34)                       # 4th vertical line
                        pdf.line(170,110, 170,h-34)                       # 5th vertical line
                        pdf.line(195,110, 195,h-34) 

                    elif h>228 and h<=235:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,110, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-28)                       # 4th vertical line
                        pdf.line(170,110, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line
                    
                    elif h>235 and h<=241:
                        pdf.line(15,110, 15,h-33)                       # 1st vertical line
                        pdf.line(40,110, 40,h-33)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-33)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-33)                       # 4th vertical line
                        pdf.line(170,110, 170,h-33)                       # 5th vertical line
                        pdf.line(195,110, 195,h-33)                       # 6th vertical line

# 243.1 ,241.10000000000008                                       
                    elif h==243.10000000000008 or h==241.10000000000008 :
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,110, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-32)                       # 4th vertical line
                        pdf.line(170,110, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)                       # 6th vertical line

                    elif h== 247.80000000000007:
                        pdf.line(15,110, 15,h-36)                       # 1st vertical line
                        pdf.line(40,110, 40,h-36)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-36)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-36)                       # 4th vertical line
                        pdf.line(170,110, 170,h-36)                       # 5th vertical line
                        pdf.line(195,110, 195,h-36) 
# 246.3
                    elif h>241 and h<=250:
                        pdf.line(15,110, 15,h-25)                       # 1st vertical line
                        pdf.line(40,110, 40,h-25)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-25)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-25)                       # 4th vertical line
                        pdf.line(170,110, 170,h-25)                       # 5th vertical line
                        pdf.line(195,110, 195,h-25)                       # 6th vertical line
# this for only 252.7000000000001
                    elif h==252.7000000000001:
                        pdf.line(15,110, 15,h-31)                       # 1st vertical line
                        pdf.line(40,110, 40,h-31)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-31)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-31)                       # 4th vertical line
                        pdf.line(170,110, 170,h-31)                       # 5th vertical line
                        pdf.line(195,110, 195,h-31)                       # 6th vertical line
                         
                    elif h==253.10000000000008:
                        pdf.line(15,110, 15,h-36)                       # 1st vertical line
                        pdf.line(40,110, 40,h-36)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-36)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-36)                       # 4th vertical line
                        pdf.line(170,110, 170,h-36)                       # 5th vertical line
                        pdf.line(195,110, 195,h-36)    
                    
                    elif h == 254.10000000000008:
                        pdf.line(15,110, 15,h-36)                       # 1st vertical line
                        pdf.line(40,110, 40,h-36)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-36)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-36)                       # 4th vertical line
                        pdf.line(170,110, 170,h-36)                       # 5th vertical line
                        pdf.line(195,110, 195,h-36)   

                    

# this for only 255.0000000000001 
                    elif h==255.0000000000001 :
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,110, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-28)                       # 4th vertical line
                        pdf.line(170,110, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line

# this for only 256.0000000000001 
                    elif h==256.0000000000001 :
                        pdf.line(15,110, 15,h-29)                       # 1st vertical line
                        pdf.line(40,110, 40,h-29)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-29)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-29)                       # 4th vertical line
                        pdf.line(170,110, 170,h-29)                       # 5th vertical line
                        pdf.line(195,110, 195,h-29)                       # 6th vertical line

# this for only 256.0000000000001 and  h==252.4000000000001
                    elif  h==256.4000000000001 or h==252.4000000000001 or h==254.7000000000001:
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,110, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-32)                       # 4th vertical line
                        pdf.line(170,110, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)                       # 6th vertical line

                    elif h==257.0000000000001 :
                        pdf.line(15,110, 15,h-29)                       # 1st vertical line
                        pdf.line(40,110, 40,h-29)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-29)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-29)                       # 4th vertical line
                        pdf.line(170,110, 170,h-29)                       # 5th vertical line
                        pdf.line(195,110, 195,h-29)

# this is for 257.7000000000001
                    elif h==257.7000000000001 :
                        pdf.line(15,110, 15,h-33)                       # 1st vertical line
                        pdf.line(40,110, 40,h-33)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-33)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-33)                       # 4th vertical line
                        pdf.line(170,110, 170,h-33)                       # 5th vertical line
                        pdf.line(195,110, 195,h-33)                       # 6th vertical line

 #258.1000000000001     
                    elif h== 258.1000000000001 or h == 258.4000000000001:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,110, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-37)                       # 4th vertical line
                        pdf.line(170,110, 170,h-37)                       # 5th vertical line
                        pdf.line(195,110, 195,h-37)

# this is for 259.4000000000001,260.4000000000001 
                    elif h==259.4000000000001 or h==260.4000000000001 :
                        pdf.line(15,110, 15,h-35)                       # 1st vertical line
                        pdf.line(40,110, 40,h-35)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-35)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-35)                       # 4th vertical line
                        pdf.line(170,110, 170,h-35)                       # 5th vertical line
                        pdf.line(195,110, 195,h-35)                       # 6th vertical line


# 251.3, 256.6, 
                    elif h>250 and h<=260:
                        pdf.line(15,110, 15,h-24)                       # 1st vertical line
                        pdf.line(40,110, 40,h-24)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-24)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-24)                       # 4th vertical line
                        pdf.line(170,110, 170,h-24)                       # 5th vertical line
                        pdf.line(195,110, 195,h-24)                       # 6th vertical line
                    
                    elif h == 266.7000000000001:
                        pdf.line(15,110, 15,h-35)                       # 1st vertical line
                        pdf.line(40,110, 40,h-35)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-35)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-35)                       # 4th vertical line
                        pdf.line(170,110, 170,h-35)                       # 5th vertical line
                        pdf.line(195,110, 195,h-35)   
                    
                    elif h == 268.50000000000017:
                        pdf.line(15,110, 15,h-25)                       # 1st vertical line
                        pdf.line(40,110, 40,h-25)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-25)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-25)                       # 4th vertical line
                        pdf.line(170,110, 170,h-25)                       # 5th vertical line
                        pdf.line(195,110, 195,h-25)                       # 6th vertical line

                    
                    elif h == 269.3000000000001:
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,110, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-32)                       # 4th vertical line
                        pdf.line(170,110, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)


# 266.1,271.4
                    elif h>260 and h<=272:
                        pdf.line(15,110, 15,h-38)                       # 1st vertical line
                        pdf.line(40,110, 40,h-38)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-38)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-38)                       # 4th vertical line
                        pdf.line(170,110, 170,h-38)                       # 5th vertical line
                        pdf.line(195,110, 195,h-38)                       # 6th vertical line

# 273.7,274,279                   
                    elif h>272 and h<=280:
                        pdf.line(15,110, 15,h-20)                       # 1st vertical line
                        pdf.line(40,110, 40,h-20)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-20)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-20)                       # 4th vertical line
                        pdf.line(170,110, 170,h-20)                       # 5th vertical line
                        pdf.line(195,110, 195,h-20)                       # 6th vertical line
# 289.2,294.5
                    elif h>280 and h<=300:
                        pdf.line(15,110, 15,h-19)                       # 1st vertical line
                        pdf.line(40,110, 40,h-19)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-19)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-19)                       # 4th vertical line
                        pdf.line(170,110, 170,h-19)                       # 5th vertical line
                        pdf.line(195,110, 195,h-19)                       # 6th vertical line
# 307.6
                    elif h>300 and h<=320:
                        pdf.line(15,110, 15,h-13)                       # 1st vertical line
                        pdf.line(40,110, 40,h-13)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-13)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-13)                       # 4th vertical line
                        pdf.line(170,110, 170,h-13)                       # 5th vertical line
                        pdf.line(195,110, 195,h-13)                       # 6th vertical line
                    elif h>320 and h<=380:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,110, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,110, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,110, 140,h-37)                       # 4th vertical line
                        pdf.line(170,110, 170,h-37)                       # 5th vertical line
                        pdf.line(195,110, 195,h-37)                       # 6th vertical line

                    else:
                        pass
                # if '12038028117400' in final_df['account_no'][i]:
                    # print(h,'-------------------nonbalbf')
                max_page_no=final_df['Page_no'][i]
                if final_df['account_no'][i] in message_list_digital_acc_no:
                        k1=final_df['account_no'][i]
                        k2=final_message_list_digital[k1]
                        if k1 in d_email.keys():
                            go_page=d_email[k1]
                            # print(k1,go_page,'nobal')

                if final_df['msg_total_debit'][i] == 'Test' and  final_df['Page_no'][i]==str(go_page):
                    t=45
                    if final_message_list_digital[k1][0]=='TEST':
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-10,180)     
                        pdf.line(15,h-30,15,h+10)                             # 1st vertical line of total table  
                        pdf.line(195,h-30,195,h+10) 
                        pdf.line(15,h+10,195,h+10)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+105,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+110 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+115,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+120,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+125,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+130,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')
                        
                    if 'TEST' not in final_message_list_digital[k1]:
                        h=h+t
                        for k in range(0,max_size_message):
                            if k2[k]!='TEST':
                                if k2[k]=='TEST':
                                    break
                                pdf.cell(60)
                                pdf.cell(10,h,txt = str(k2[k]))
                                pdf.ln(1)
                                h=h+5


                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-110,180)     
                        pdf.line(15,h-50,15,h-130)                             # 1st vertical line of total table  
                        pdf.line(195,h-50,195,h-130) 
                        pdf.line(15,h-50,195,h-50)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+105,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+110 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+115,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+120,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+125,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+130,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')

                if  final_df['msg_total_debit'][i] != 'Test' and final_df['Page_no'][i]==str(go_page):
                    pdf.ln(1)
                    pdf.set_line_width(0.01)  

                    if h>140 and h<=150:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                        pdf.line(15,h+16,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+16,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>150 and h<=160:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-24,180)
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>160 and h<=165:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+21,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+21,txt = final_df['monthly_avg'][i])
                    if h>165 and h<=170:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22.5,180)
                        pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+14,195,h-30) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+23,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+23,txt = final_df['monthly_avg'][i])
                    if h>170 and h<=175:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>175 and h<=180:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26.8,180)
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>180 and h<=190:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                        pdf.line(15,h+8,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+8,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>190 and h<=200:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-28,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>200 and h<=210:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>210 and h<=220:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-21,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>220 and h<=240:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-18,180)
                        pdf.line(15,h+18,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+18,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>240 and h<=250:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-30,180)
                        pdf.line(15,h+8,15,h-45)                             # 1st vertical line of total table  
                        pdf.line(195,h+8,195,h-45) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>250 and h<=260:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-50,180)
                        pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                        pdf.line(195,h-10,195,h-50) 

                    if k2[k][0]=='TEST':
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                        pdf.line(195,h+15,195,h-35) 
                        pdf.line(15,h+15,195,h+15)
                        # pdf.line(15,h+15,195,h+15)
                    else:
                        t=55
                        pdf.ln(1)
                        for k in range(0,max_size_message):
                            if k2[k]!='TEST':
                                if k2[k]=='TEST':
                                    break
                                pdf.cell(60)
                                pdf.cell(10,h+t,txt = str(k2[k]))
                                pdf.ln(1)
                                t=t+5

                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h,180)     
                        pdf.line(15,h-50,15,h+43)                             # 1st vertical line of total table  
                        pdf.line(195,h-50,195,h+43) 
                        pdf.line(15,h+43,195,h+43)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+145,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+150 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+155,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+160,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+165,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+170,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')

                elif  final_df['msg_total_debit'][i] != 'Test':
                    pdf.ln(1)
                    pdf.set_line_width(0.01)  
                    pdf.set_line_width(0.01)  
                    t=55
                    for k in range(0,max_size_message):
                        if k2[k]!='TEST':
                            if k2[k]=='TEST':
                                break
                            pdf.cell(80)
                            pdf.cell(10,h+t,txt = str(k2[k]))
                            pdf.ln(1)
                            t=t+5
                    if h>140 and h<=150:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                        pdf.line(15,h+16,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+16,195,h-35) 
                        pdf.line(15,h+16,195,h+16)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>150 and h<=160:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.line(15,h+15,195,h+15)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>160 and h<=165:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22,180)
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.line(15,h+15,195,h+15)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+16,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+16,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+21,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+21,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+26,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+26,txt = final_df['monthly_avg'][i])
                    if h>165 and h<=170:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22,180)
                        pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+14,195,h-30) 
                        pdf.line(15,h+14,195,h+14)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+23,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+23,txt = final_df['monthly_avg'][i])
                    if h>170 and h<=175:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>175 and h<=180:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26.8,180)
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>180 and h<=190:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-29.5,180)
                        pdf.line(15,h+8,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+8,195,h-35) 
                        pdf.line(15,h+8,195,h+8)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>190 and h<=200:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-30,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>200 and h<=210:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>210 and h<=215:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-28,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+11,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>216 and h<=220:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-31,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+11,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>220 and h<=228:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-30,180)
                        pdf.line(15,h+12,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+12,195,h-30) 
                        pdf.line(15,h+12,195,h+12)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>228 and h<=230:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                        pdf.line(15,h+12,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+12,195,h-30) 
                        pdf.line(15,h+12,195,h+12)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>230 and h<=240:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-33,180)
                        pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+14,195,h-30) 
                        pdf.line(15,h+14,195,h+14)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>240 and h<=245:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-34,180)
                        pdf.line(15,h+10,15,h-45)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-45) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+14,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+14,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>245 and h<=250:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                        pdf.line(15,h+10,15,h-45)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-45) 
                        pdf.line(15,h+10,195,h+10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+14,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+14,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>250 and h<=260:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-33,180)
                        pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                        pdf.line(195,h-10,195,h-50) 
                        pdf.line(15,h-10,195,h-10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+14,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+14,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>255 and h<=260:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-35,180)
                        pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                        pdf.line(195,h-10,195,h-50) 
                        pdf.line(15,h-10,195,h-10)
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+14,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+14,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])

                        

                rrr+=1

            if 'TEST' not in final_df['bal_bf'][i]:
                pdf.ln(1)
                pdf.set_fill_color(r=211,g=211,b=211)
                pdf.set_line_width(0.01)
                pdf.rect(x = 15 , y = 100 , w = 180 , h = 15, style = 'DF')

                pdf.set_line_width(0.3)
                pdf.line(40,100, 40, 115)                      # first  | line
                pdf.line(110,100, 110, 115)                    # second | line
                pdf.line(140,100, 140, 115)                    # third | line
                pdf.line(170,100,170,115)                      # fourth | line

                #-------------------------------------INSERTING 2ND PAGE TABLE COLUMN NAMES---------------------------------

                pdf.ln(1)
                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,134,txt = 'TARIKH')
                pdf.cell(6)
                pdf.cell(30,134,txt = 'KETERANGAN')
                pdf.cell(30)
                pdf.cell(30,134,txt = 'DEBIT')
                pdf.cell(4)
                pdf.cell(25,134,txt = 'KREDIT')
                pdf.cell(1)
                pdf.cell(25,134,txt = 'BAKI')


                pdf.ln(4)

                pdf.set_font('Arial', 'I', 9)
                pdf.cell(10)
                pdf.cell(30,136,txt = 'DATE')
                pdf.cell(6)
                pdf.cell(30,136,txt = 'DESCRIPTION')
                pdf.cell(30)
                pdf.cell(25,136,txt = 'DEBIT')
                pdf.cell(8)
                pdf.cell(25,136,txt = 'CREDIT')
                pdf.cell(1)
                pdf.cell(25,136,txt = 'BALANCE')

                pdf.ln(4)

                pdf.set_font('Arial', 'B', 9)
                pdf.cell(10)
                pdf.cell(30,138,txt = '')
                pdf.cell(6)
                pdf.cell(30,138,txt = '')
                pdf.cell(35)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(5)
                pdf.cell(25,138,txt = 'RM')
                pdf.cell(1)
                pdf.cell(25,138,txt = 'RM')
                pdf.set_line_width(0.01)
                pdf.ln(1)
                pdf.set_font('Arial', 'B', 8)
                pdf.cell(30)
                pdf.cell(10,145,txt='BAL B/F')
                pdf.cell(123)
                pdf.cell(10,145,txt= final_df['bal_bf'][i])
                pdf.line(15,119,195,119)
                pdf.ln(2)


#--------------------------------------------INSERTING 1ST TABLE DATA OF 1ST PAGE-----------------
                h = 147
                p2=h
                cc=1
                dd=115

                m1=h
                m2=h
                flag11=False
                line_flag = 'a'
                start=0
                end=0
                lll+=1
                # print(total_debit)
                if final_line_flag_list_email[lll]!=True:
                    for j in range(0,max_List1):
                        if final_df['total_date_list'][i][j] != 'Test':
                            # print(final_df['total_date_list'][i][j])
                            pdf.set_line_width(0.005)
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc2'][i][j]!='desc_TEST':
                                h=h+2
                            elif final_df['total_desc1'][i][j]!='desc_TEST':
                                h=h+4
                            elif final_df['total_desc0'][i][j]!='desc_TEST':
                                h=h+5
                                max_size='s1'            
                            pdf.ln(1)
                            pdf.set_font('Arial','B',8)
                            pdf.cell(10)
                            pdf.cell(10,h,txt = final_df['total_date_list'][i][j])

                            pdf.cell(10)
                            pdf.cell(10,h,txt = str(final_df['total_desc0'][i][j]))

                            pdf.cell(80)
                            pdf.cell(10,h,txt = final_df['total_debit'][i][j])
                            pdf.cell(20)
                            pdf.cell(10,h,txt = final_df['total_credit'][i][j])

                            pdf.cell(4)
                            pdf.cell(10,h,txt = final_df['total_bal'][i][j]) 
                            pdf.ln(1)

                            if final_df['total_desc0'][i][j]!='desc_TEST':
                                cc=cc+1  
                                line_flag = 'a0'
                            if final_df['total_desc1'][i][j]!='desc_TEST':
                                pdf.ln(1)
                                pdf.cell(30)
                                h=h+4
                                pdf.set_font('Arial','B',8)
                                pdf.cell(20,h,txt =  str(final_df['total_desc1'][i][j]))
                                cc=cc+1  
                                line_flag = 'a1'
                            if final_df['total_desc2'][i][j]!='desc_TEST':                                                
                                pdf.ln(1)                                              
                                pdf.cell(30)
                                h=h+4                    
                                pdf.cell(10,h,txt =  str(final_df['total_desc2'][i][j]))   #  27,29,31,33,35                         
                                cc=cc+1  
                                line_flag = 'a2'
                            if final_df['total_desc3'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  str(final_df['total_desc3'][i][j]))   
                                cc=cc+1  
                                line_flag = 'a3'
                            if final_df['total_desc4'][i][j]!='desc_TEST':
                                pdf.ln(1)        
                                pdf.cell(30)
                                h=h+4
                                pdf.cell(10,h,txt =  str(final_df['total_desc4'][i][j]))   
                                cc=cc+1  
                                line_flag = 'a3'
                            p=h
                            h=h+0.3
                            pdf.ln(1)       
                            pdf.set_font('Arial','',8)
                            if final_line_flag_list_email[lll][j]=='1line':
                                pdf.cell(4)
                                pdf.cell(10,h-4,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_email[lll][j]=='2line':
                                pdf.cell(4)
                                pdf.cell(10,h-3,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_email[lll][j]=='3line':
                                pdf.cell(4)
                                pdf.cell(10,h-2,txt = '__________________________________________________________________________________________________________________')   
                            if final_line_flag_list_email[lll][j]=='4line':
                                pdf.cell(4)
                                pdf.cell(10,h-3.5,txt = '__________________________________________________________________________________________________________________')   

                    # if '14014021101661' in final_df['account_no'][i]:
                    #     print(h,'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<,')
# done 162.9
                    if h>150 and h<=200:
                        pdf.line(15,115, 15,h-26)                       # 1st vertical line
                        pdf.line(40,119, 40,h-26)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-26)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-26)                       # 4th vertical line
                        pdf.line(170,119, 170,h-26)                       # 5th vertical line
                        pdf.line(195,115, 195,h-26)                       # 6th vertical line
                    
                    elif h == 217.80000000000007:
                        pdf.line(15,110, 15,h-34)                       # 1st vertical line
                        pdf.line(40,119, 40,h-34)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-34)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-34)                       # 4th vertical line
                        pdf.line(170,119, 170,h-34)                       # 5th vertical line
                        pdf.line(195,110, 195,h-34)  
# done 205.3
                    elif h>200 and h<=220:
                        pdf.line(15,110, 15,h-23)                       # 1st vertical line
                        pdf.line(40,119, 40,h-23)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-23)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-23)                       # 4th vertical line
                        pdf.line(170,119, 170,h-23)                       # 5th vertical line
                        pdf.line(195,110, 195,h-23)                       # 6th vertical line
# done 220.7
                    elif h>220 and h<=230:
                        pdf.line(15,110, 15,h-27)                       # 1st vertical line
                        pdf.line(40,119, 40,h-27)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-27)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-27)                       # 4th vertical line
                        pdf.line(170,119, 170,h-27)                       # 5th vertical line
                        pdf.line(195,110, 195,h-27)                       # 6th vertical line
# done 235.4
                    elif h>230 and h<=250:
                        pdf.line(15,110, 15,h-32)                       # 1st vertical line
                        pdf.line(40,119, 40,h-32)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-32)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-32)                       # 4th vertical line
                        pdf.line(170,119, 170,h-32)                       # 5th vertical line
                        pdf.line(195,110, 195,h-32)                       # 6th vertical line

                    elif h == 252.10000000000008:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,119, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-37)                       # 4th vertical line
                        pdf.line(170,119, 170,h-37)                       # 5th vertical line
                        pdf.line(195,110, 195,h-37)  


                    elif h>250 and h<=261:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,119, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-28)                       # 4th vertical line
                        pdf.line(170,119, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line
# done 263.1
                    elif h>261 and h<=265:
                        pdf.line(15,110, 15,h-39)                       # 1st vertical line
                        pdf.line(40,119, 40,h-39)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-39)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-39)                       # 4th vertical line
                        pdf.line(170,119, 170,h-39)                       # 5th vertical line
                        pdf.line(195,119, 195,h-39)                       # 6th vertical line
                    
                    elif h == 265.4000000000001:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,119, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-37)                       # 4th vertical line
                        pdf.line(170,119, 170,h-37)                       # 5th vertical line
                        pdf.line(195,119, 195,h-37) 
                    
                    elif h == 267.4000000000001:
                        pdf.line(15,110, 15,h-39)                       # 1st vertical line
                        pdf.line(40,119, 40,h-39)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-39)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-39)                       # 4th vertical line
                        pdf.line(170,119, 170,h-39)                       # 5th vertical line
                        pdf.line(195,119, 195,h-39) 
# 265.1000000000001                    
                    elif h == 265.1000000000001:
                        pdf.line(15,110, 15,h-40)                       # 1st vertical line
                        pdf.line(40,119, 40,h-40)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-40)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-40)                       # 4th vertical line
                        pdf.line(170,119, 170,h-40)                       # 5th vertical line
                        pdf.line(195,110, 195,h-40)

# done 266.9
                    elif h>265 and h<=267:
                        pdf.line(15,110, 15,h-28)                       # 1st vertical line
                        pdf.line(40,119, 40,h-28)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-28)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-28)                       # 4th vertical line
                        pdf.line(170,119, 170,h-28)                       # 5th vertical line
                        pdf.line(195,110, 195,h-28)                       # 6th vertical line
# done 268.6
                    elif h>267 and h<=269:
                        pdf.line(15,110, 15,h-30)                       # 1st vertical line
                        pdf.line(40,119, 40,h-30)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-30)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-30)                       # 4th vertical line
                        pdf.line(170,119, 170,h-30)                       # 5th vertical line
                        pdf.line(195,110, 195,h-30)                       # 6th vertical line


# done 269.7
                    elif h>269 and h<=275:
                        pdf.line(15,110, 15,h-37)                       # 1st vertical line
                        pdf.line(40,119, 40,h-37)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-37)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-37)                       # 4th vertical line
                        pdf.line(170,119, 170,h-37)                       # 5th vertical line
                        pdf.line(195,110, 195,h-37)                       # 6th vertical line
# done 
                    elif h>275 and h <=276:
                        pdf.line(15,110, 15,h-42)                       # 1st vertical line
                        pdf.line(40,119, 40,h-42)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-42)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-42)                       # 4th vertical line
                        pdf.line(170,119, 170,h-42)                       # 5th vertical line
                        pdf.line(195,110, 195,h-42)                       # 6th vertical line

# done                    
                    elif h>276 and h<=280:
                        pdf.line(15,110, 15,h-24)                       # 1st vertical line
                        pdf.line(40,119, 40,h-24)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-24)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-24)                       # 4th vertical line
                        pdf.line(170,119, 170,h-24)                       # 5th vertical line
                        pdf.line(195,110, 195,h-24)                       # 6th vertical line
# done 284.3
                    elif h>280 and h<=285:
                        pdf.line(15,110, 15,h-23)                       # 1st vertical line
                        pdf.line(40,119, 40,h-23)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-23)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-23)                       # 4th vertical line
                        pdf.line(170,119, 170,h-23)                       # 5th vertical line
                        pdf.line(195,110, 195,h-23)                       # 6th vertical line

                    elif h>285 and h<=286:
                        pdf.line(15,110, 15,h-47)                       # 1st vertical line
                        pdf.line(40,119, 40,h-47)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-47)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-47)                       # 4th vertical line
                        pdf.line(170,119, 170,h-47)                       # 5th vertical line
                        pdf.line(195,110, 195,h-47)                       # 6th vertical line
# done 292.9
                    elif h>286 and h<=300:
                        pdf.line(15,110, 15,h-23)                       # 1st vertical line
                        pdf.line(40,119, 40,h-23)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-23)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-23)                       # 4th vertical line
                        pdf.line(170,119, 170,h-23)                       # 5th vertical line
                        pdf.line(195,110, 195,h-23)                       # 6th vertical line
# done 
                    elif h>300 and h<=320:
                        pdf.line(15,110, 15,h-16)                       # 1st vertical line
                        pdf.line(40,119, 40,h-16)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-16)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-16)                       # 4th vertical line
                        pdf.line(170,119, 170,h-16)                       # 5th vertical line
                        pdf.line(195,110, 195,h-16)                       # 6th vertical line

                    elif h>320 and h<=380:
                        pdf.line(15,110, 15,h-40)                       # 1st vertical line
                        pdf.line(40,119, 40,h-40)                       # 2nd vertical line
                        pdf.line(110,119, 110,h-40)                       # 3rd vertical line
                        pdf.line(140,119, 140,h-40)                       # 4th vertical line
                        pdf.line(170,119, 170,h-40)                       # 5th vertical line
                        pdf.line(195,119, 195,h-40)                       # 6th vertical line

                    else:
                        pass
                
                # if '12038022006490' in final_df['account_no'][i]:
                    # print(h,'-----------------balbf')
                max_page_no=final_df['Page_no'][i]
                if final_df['account_no'][i] in message_list_digital_acc_no:
                        k1=final_df['account_no'][i]
                        k2=final_message_list_digital[k1]
                        if k1 in d_email.keys():
                            go_page=d_email[k1]
                            # print(k1,go_page,'nobal')

                if final_df['msg_total_debit'][i] == 'Test' and  final_df['Page_no'][i]==str(go_page):
                    t=45
                    if final_message_list_digital[k1][0]=='TEST':
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-10,180)     
                        pdf.line(15,h-30,15,h+10)                             # 1st vertical line of total table  
                        pdf.line(195,h-30,195,h+10) 
                        pdf.line(15,h+10,195,h+10)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+105,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+110 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+115,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+120,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+125,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+130,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')
                        
                    if 'TEST' not in final_message_list_digital[k1]:
                        h=h+t
                        for k in range(0,max_size_message):
                            if k2[k]!='TEST':
                                if k2[k]=='TEST':
                                    break
                                pdf.cell(60)
                                pdf.cell(10,h,txt = str(k2[k]))
                                pdf.ln(1)
                                h=h+5


                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-110,180)     
                        pdf.line(15,h-50,15,h-130)                             # 1st vertical line of total table  
                        pdf.line(195,h-50,195,h-130) 
                        pdf.line(15,h-50,195,h-50)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+105,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+110 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+115,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+120,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+125,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+130,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')
                
                if  final_df['msg_total_debit'][i] != 'Test' and final_df['Page_no'][i]==str(go_page):
                    pdf.ln(1)
                    pdf.set_line_width(0.01)  

                    if h>140 and h<=150:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-23,180)
                        pdf.line(15,h+16,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+16,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>150 and h<=160:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-25,180)
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>160 and h<=165:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                        pdf.line(15,h+15,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+15,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+21,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+21,txt = final_df['monthly_avg'][i])
                    if h>165 and h<=170:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-22.5,180)
                        pdf.line(15,h+14,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+14,195,h-30) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+18,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+23,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+23,txt = final_df['monthly_avg'][i])
                    if h>170 and h<=175:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26,180)
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>175 and h<=180:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-26.8,180)
                        pdf.line(15,h+10,15,h-30)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-30) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>180 and h<=190:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                        pdf.line(15,h+8,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+8,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>190 and h<=200:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-28,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>200 and h<=210:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-27,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+12,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+12,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+17,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+17,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+22,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+22,txt = final_df['monthly_avg'][i])
                    if h>210 and h<=220:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-32,180)
                        pdf.line(15,h+10,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+10,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+10,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+10,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+15,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+15,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+20,txt = final_df['monthly_avg'][i])
                    if h>220 and h<=240:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-18,180)
                        pdf.line(15,h+18,15,h-35)                             # 1st vertical line of total table  
                        pdf.line(195,h+18,195,h-35) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+8,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+8,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+13,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+13,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+18,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+18,txt = final_df['monthly_avg'][i])
                    if h>240 and h<=250:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-30,180)
                        pdf.line(15,h+8,15,h-45)                             # 1st vertical line of total table  
                        pdf.line(195,h+8,195,h-45) 
                        pdf.ln(1)
                        pdf.set_font('Arial','B',8)
                        pdf.cell(20)
                        pdf.cell(10,h+20,txt = 'TOTAL DEBIT')
                        pdf.cell(30)
                        pdf.cell(20,h+20,txt = final_df['msg_total_debit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+25,txt = 'TOTAL CREDIT')
                        pdf.cell(30)
                        pdf.cell(20,h+25,txt = final_df['msg_total_credit'][i])
                        pdf.ln(1)
                        pdf.cell(20)
                        pdf.cell(10,h+30,txt = 'MONTHLY AVERAGE')
                        pdf.cell(30)
                        pdf.cell(40,h+30,txt = final_df['monthly_avg'][i])
                    if h>250 and h<=260:
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\SumBox1.png',15.09,h-50,180)
                        pdf.line(15,h-10,15,h-50)                             # 1st vertical line of total table  
                        pdf.line(195,h-10,195,h-50) 

                    t=55
                    pdf.ln(1)
                    if k2[k]=='TEST':
                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-5,180) 
                        pdf.line(195,h+15,195,h-35) 
                        pdf.line(15,h+15,195,h+15)
                        pdf.line(15,h+15,195,h+15)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+80,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+85 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+90,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+95,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+100,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+105,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')


                    else:
                        for k in range(0,max_size_message):
                            if k2[k]!='TEST':
                                if k2[k]=='TEST':
                                    break
                                pdf.cell(80)
                                pdf.cell(10,h+t,txt = str(k2[k]))
                                pdf.ln(1)
                                t=t+5

                        pdf.image('E:\IWOC\APP SCRIPT\BICASA\CYCLE\BICASA\DATA\MsgBox1.png',15,h-110,180)     
                        pdf.line(15,h-50,15,h-130)                             # 1st vertical line of total table  
                        pdf.line(195,h-50,195,h-130) 
                        pdf.line(15,h-50,195,h-50)

                        pdf.ln(1)
                        pdf.set_font('Arial','',7.2)
                        pdf.cell(4)
                        pdf.cell(10,h+135,txt ='Sekiranya anda mendapati sebarang perbezaan, sila maklumkan kepada pihak Bank di dalam tempoh 14 hari daripada tarikh penyata ini. Jika tiada perbezaan')   
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+140 , txt = 'penyata ini akan dianggap betul. ')
                        pdf.set_font('Arial','I',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+145,txt= 'If you note any discrepancies, please advise the Bank within 14 days from the date of this statement, otherwise this account statement is considered to be correct. ')
                        pdf.set_font('Arial','',7.2)
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+150,txt = 'Untuk pertanyaan, ajukan kepada / For enquiries, please channel to: ')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+155,txt = 'Jabatan Khidmat Pelanggan (Customer Care Department), Tingkat 17, Menara Bank Islam, No 22, Jalan Perak, 50450 Kuala Lumpur')
                        pdf.ln(1)
                        pdf.cell(4)
                        pdf.cell(10,h+160,txt = 'Tel: 03-26 900 900 / Faks: 03-2782 1337. Emel / Email: contactcenter@bankislam.com.my')

        if final_df['account_no'][i]==last_acc:
        # print(account_no[i],'---------')
            pdf_merger = PdfFileMerger()
            date1 = datetime.now()
            d = date1.strftime('%d')
            m = date1.strftime('%m')
            y = date1.strftime('%y')
            y = y[2:4]

            if m == '01':
                month = 'JAN'
            if m == '02':
                month = 'FEB'
            if m == '03':
                month = 'MAR'
            if m == '04':
                month = 'APR'
            if m == '05':
                month = 'MAY'
            if m == '06':
                month = 'JUN'
            if m == '07':
                month = 'JUL'
            if m == '08':
                month = 'AUG'
            if m == '09':
                month = 'SEP'
            if m == '10':
                month = 'OCT'
            if m == '11':
                month = 'NOV'
            if m == '12':
                month = 'DEC'
            filename = os.path.abspath(r"%s\C-%s%s%s_0000%s.pdf"%(kk1,d,month,y,kc1))
            f = filename    
            pdf.output(f)   
            kc1 +=  1  
            # pdf=FPDF()
            break

    
    path20 = r'E:\IWOC\SOURCE' 
    target = r'E:\IWOC\TEMP' 

    for i in move_files:
        src_path = os.path.join(path20, i)
        dst_path = os.path.join(target, i)
        os.rename(src_path, dst_path)
    print('Files Moved................')