from datetime import date, datetime
import time
from fpdf import FPDF

from APP.models import Scheduler

def MQ9(request,files1):
    move_files = []
    move_files.append(files1)
    cppp = 1
    def Folder_creation(cppp):
        # ------------------Capture date----------------
        cppp=str(cppp)
        today11=date.today()
        d1 = today11.strftime("%d-%b-%Y")
        d2=d1[0:2]
        d3=d1[3:6]
        d1=d2+d3+cppp
        # ------------------Capture time----------------
        now = datetime.now()        
        start_time = now.strftime("%H:%M:%S")
        s1=start_time[0:2]
        s2=start_time[3:5]
        s3=start_time[6:8]
        start_time=s1+s2+s3

        # ---------------create first date folder------------

        # directory=d1
        # parent_dir = "E:\IWOC\OUTPUT\Bank Simpanan Nasional\CYCLE\\"
        outp = Scheduler.objects.filter(PROGRAM_NAME='MQ9').values()
        res = []
        otp = []
        for idx, sub in enumerate(outp, start = 0):
            if idx == 0:
                res.append(list(sub.values()))
        new = [i[-1] for i in res]
        import os
        for z in new:
            directory=d1
            # parent_dir = "E:\IWOC\OUTPUT\Bank Simpanan Nasional\CYCLE\\"
            parent_dir = z
            directory = d1+"_"+start_time 
            path1 = os.path.join(parent_dir, directory) # date folder path
            os.mkdir(path1)
            # print("Directory '% s' created" % directory)

            directory = 'MQ9'
            parent_dir = path1
            path2 = os.path.join(path1, directory) # bkrcc folder path
            os.mkdir(path2)
            # print("Directory '% s' created" % directory)

            directory = 'PRINTING'
            parent_dir = path2
            path3 = os.path.join(path2, directory) # DIGITAL folder path
            os.mkdir(path3)
            # print("Directory '% s' created" % directory)

            # directory = 'PRINTING'
            # parent_dir = path2
            # path4 = os.path.join(path2, directory) # PRINTING folder path
            # os.mkdir(path4)
            # print("Directory '% s' created" % directory)

            return path3

    path11 = r'E:\IWOC\SOURCE\\'+ files1
    with open(path11) as file:
        job_id = []
        job_name = []
        user_id = []
        sysout_class = []
        output_group = []
        title = []
        destination = []
        name = []
        room = []
        building = []
        departement = []
        address = []
        for line in file:
            if line.startswith('* JOBID:') and not line.endswith('* JOBID:'):
                job_id.append(line[17:30])
                line = file.readline()
                job_name.append(line[17:30])
                line = file.readline()
                user_id.append(line[17:30])
                line = file.readline()
                sysout_class.append(line[17:30])
                line = file.readline()
                output_group.append(line[17:30])
                line = file.readline()
                title.append(line[17:30])
                line = file.readline()
                line = file.readline()
                destination.append(line[17:30])
                line = file.readline()
                name.append(line[17:30])
                line = file.readline()
                room.append(line[17:30])
                line = file.readline()
                building.append(line[17:30])
                line = file.readline()
                departement.append(line[17:30])
                line = file.readline()
                address.append(line[17:30])
            
    with open(path11) as file:
        for line in file:
            line = file.readlines()
            total_line = len(line)
        # print(total_line)

    pdf = FPDF(orientation='L')
    pdf.set_auto_page_break(False)
    pdf.add_page()

    pdf.image(r'E:\IWOC\APP SCRIPT\MQ9\CYCLE\mq9\DATA\Capture.PNG',10,10,120)

    pdf.ln(1)
    pdf.set_font('Courier','B',6)
    pdf.cell(-6)
    pdf.cell(10,214,txt = 'JOBID:')
    pdf.cell(10)
    pdf.cell(10,214,txt = str(job_id[0]))

    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,220,txt = 'JOB NAME:')
    pdf.cell(10)
    pdf.cell(10,220,txt = str(job_name[0]))

    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,226,txt = 'USER ID:')
    pdf.cell(10)
    pdf.cell(10,226,txt = str(user_id[0]))

    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,232,txt = 'SYSYOUT CLASS:')
    pdf.cell(10)
    pdf.cell(10,232,txt = sysout_class[0])

    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,238,txt = 'OUTPUT GROUP')
    pdf.cell(10)
    pdf.cell(10,238,txt = str(output_group[0]))

    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,244,txt = 'TITLE:')
    pdf.cell(10)
    pdf.cell(10,244,txt = str(title[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,256,txt = 'DESTINATION:')
    pdf.cell(10)
    pdf.cell(10,256,txt = str(destination[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,262,txt = 'NAME:')
    pdf.cell(10)
    pdf.cell(10,262,txt = str(name[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,268,txt = 'ROOM:')
    pdf.cell(10)
    pdf.cell(10,268,txt = str(room[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,274,txt = 'BUILDING:')
    pdf.cell(10)
    pdf.cell(10,274,txt = str(building[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,280,txt = ' DEPARTMENT:')
    pdf.cell(10)
    pdf.cell(10,280,txt = str(departement[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,286,txt = ' ADDRESS:')
    pdf.cell(10)
    pdf.cell(10,286,txt = str(address[0]))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,298,txt = ' PRINT TIME:')
    pdf.cell(10)


    named_tuple = time.localtime() # get struct_time
    time_string = time.strftime("%H:%M:%S", named_tuple)
    pdf.cell(10,298,txt = str(time_string))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,304,txt = ' PRINT DATE:')
    pdf.cell(10)
    named_tuple = time.localtime() # get struct_time
    time_string = time.strftime("%d %b %Y", named_tuple)
    pdf.cell(10,304,txt = str(time_string))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,310,txt = ' PRINT TIME:')
    pdf.cell(10)
    named_tuple = time.localtime() # get struct_time
    time_string = time.strftime("%H:%M:%S", named_tuple)
    pdf.cell(10,310,txt = str(time_string))


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,316,txt = ' PRINTER NAME:')
    pdf.cell(10)
    pdf.cell(10,316,txt = 'PRT14')


    pdf.ln(1)
    pdf.cell(-6)
    pdf.cell(10,322,txt = ' SYSTEM:')
    pdf.cell(10)
    pdf.cell(10,322,txt = 'MBB1')


    pdf.ln(-16)
    pdf.set_font('Courier','B',6)
    pdf.cell(-10)
    pdf.cell(1,200,txt = '**START*******START*******START*******START*******START*******START*******START*******START****')
    pdf.ln(1)
    pdf.cell(-10)
    pdf.cell(1,380,txt = '**START*******START*******START*******START*******START*******START*******START*******START****')
    pdf.ln(190)
    pdf.cell(-198)
    pdf.rotate(90)
    pdf.cell(10,380,txt = '*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *' )
    pdf.ln(-0.1)
    pdf.cell(-68)
    pdf.rotate(90)
    pdf.cell(10,360,txt = '*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *' )


    pdf.add_page()

    with open(path11) as file:
        page_1=[]
        page_2=[]
        page_3=[]
        for line in file:
            line = line.strip()
            if line.startswith('MAYBANK') and line.endswith('PAGE:   1'):
                page_1.append(line)
                line=file.readline()
                while (not line.startswith('MAYBANK') and not line.endswith('PAGE:   2')):
                    page_1.append(line)
                    line=file.readline()
            if 'PAGE:   2' in line:
            # if line.startswith('MAYBANK') and line.endswith('PAGE:   2'):
                # print('-------------------')
                page_2.append(line)
                line=file.readline()
                while (not line.startswith('MAYBANK') and not line.endswith('PAGE:   3')):
                    page_2.append(line)
    
                    line=file.readline()
            if 'PAGE:   3' in line:
            # if line.startswith('MAYBANK') and line.endswith('PAGE:   3'):
                page_3.append(line)
                line=file.readline()
                while ( not '*** END OF REPORT ***' in line):
                    page_3.append(line)
                    line=file.readline()
        # print(len(page_1))
        # print(len(page_2))
        for i in page_3:
            # print(i)
            pass
        if len(page_1)>len(page_2):
            max=len(page_1)
        else:
            max=len(page_2)
        
        while (not len(page_1)==max):
            page_1.append('TEST')

        
        
        # print(len(page_1))
        # print(len(page_2))
        # print(len(page_3))

        # print('----------------------111111111111111111111cls')
        h = -4
        for i in range(0,max):
            pdf.set_font('Courier','B',5)
            pdf.cell(-6)
            pdf.cell(2,h,txt = str(page_1[i]))
            pdf.cell(140)
            pdf.cell(2,h,txt = str(page_2[i]))
            h = h + 4
            pdf.ln(1)

    h = -4
    pdf.add_page()
    for i in range(0,len(page_3)):
        pdf.set_font('Courier','B',5)
        pdf.cell(-6)
        pdf.cell(2,h,txt = str(page_3[i]))
        h =  h  + 4
        pdf.ln(1)
        pdf.cell(10)
        pdf.image(r'E:\IWOC\APP SCRIPT\MQ9\CYCLE\mq9\DATA\Capture.PNG',140,8,120)

    pdf.ln(1)
    pdf.set_font('Courier','B',6)
    pdf.cell(102)
    pdf.cell(10,180,txt = 'JOBID:')
    pdf.cell(10)
    pdf.cell(10,180,txt = str(job_id[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,186,txt = 'JOB NAME:')
    pdf.cell(10)
    pdf.cell(10,186,txt = str(job_name[0]))

    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,192,txt = 'USER ID:')
    pdf.cell(10)
    pdf.cell(10,192,txt = str(user_id[0]))

    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,198,txt = 'SYSYOUT CLASS:')
    pdf.cell(10)
    pdf.cell(10,198,txt = str(sysout_class[0]))

    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,204,txt = 'OUTPUT GROUP')
    pdf.cell(10)
    pdf.cell(10,204,txt = str(output_group[0]))

    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,210,txt = 'TITLE:')
    pdf.cell(10)
    pdf.cell(10,210,txt = str(title[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,222,txt = 'DESTINATION:')
    pdf.cell(10)
    pdf.cell(10,222,txt = str(destination[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,228,txt = 'NAME:')
    pdf.cell(10)
    pdf.cell(10,228,txt = str(name[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,234,txt = 'ROOM:')
    pdf.cell(10)
    pdf.cell(10,234,txt = str(room[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,240,txt = 'BUILDING:')
    pdf.cell(10)
    pdf.cell(10,240,txt = str(building[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,246,txt = ' DEPARTMENT:')
    pdf.cell(10)
    pdf.cell(10,246,txt = str(departement[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,252,txt = ' ADDRESS:')
    pdf.cell(10)
    pdf.cell(10,252,txt = str(address[0]))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,264,txt = ' PRINT TIME:')
    pdf.cell(10)
    named_tuple = time.localtime() # get struct_time
    time_string = time.strftime("%H:%M:%S", named_tuple)
    pdf.cell(10,264,txt = str(time_string))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,270,txt = ' PRINT DATE:')
    pdf.cell(10)
    named_tuple = time.localtime() # get struct_time
    time_string = time.strftime("%d %b %Y", named_tuple)
    pdf.cell(10,270,txt = str(time_string))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,276,txt = ' PRINT TIME:')
    pdf.cell(10)
    named_tuple = time.localtime() # get struct_time
    time_string = time.strftime("%H:%M:%S", named_tuple)
    pdf.cell(10,276,txt = str(time_string))


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,282,txt = ' PRINTER NAME:')
    pdf.cell(10)
    pdf.cell(10,282,txt = 'PRT14')


    pdf.ln(1)
    pdf.cell(102)
    pdf.cell(10,288,txt = ' SYSTEM:')
    pdf.cell(10)
    pdf.cell(10,288,txt = 'MBB1')


    pdf.ln(-10)
    pdf.set_font('Courier','B',6)
    pdf.cell(100)
    pdf.cell(1,150,txt = '**START*******START*******START*******START*******START*******START*******START*******START****')
    pdf.ln(1)
    pdf.cell(100)
    pdf.cell(1,340,txt = '**START*******START*******START*******START*******START*******START*******START*******START****')
    pdf.ln(166)
    pdf.cell(52)
    pdf.rotate(90)
    pdf.cell(10,340,txt = '*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *' )
    pdf.ln(-0.1)
    pdf.cell(-69)
    pdf.rotate(90)
    pdf.cell(10,340,txt = '*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *' )

    kk=Folder_creation(cppp)
    bk = kk

    pdf.output(r"%s\MQ1.pdf"%kk)

    # pdf.output('E:\IWOC\OUTPUT\MQ9\MQ1.pdf')

    import os
    path20 = r'E:\IWOC\SOURCE' 
    target = r'E:\IWOC\TEMP' 

    for i in move_files:
        src_path = os.path.join(path20, i)
        dst_path = os.path.join(target, i)
        os.rename(src_path, dst_path)
    print('Files Moved................')