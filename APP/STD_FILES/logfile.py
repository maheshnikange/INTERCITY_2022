from datetime import datetime
from django.db.models import Max

##programm name varun log file name add kara or else add Program name in History table

def Log_File(today,start_time,end_time,user,ffname,files1,total_clients,totalpages,first_record,last_record):
# def Log_File(path):
    # queryset_data = []
    # allfilename = []
    # aa = History.objects.all().order_by('-JobIDX').values()
    # aa = History.objects.values().last()
    # queryset_data.append(aa)
    # print(queryset_data,'----------------------------------------queryset')
    # for index,val in enumerate(queryset_data):
        # allfilename.append(list(val.values()))
    # filenamepdf = []
    # txtfilenamepdf = []
    # totalaccpdf = []
    # totalpagespdf = []
    # firstrecordpdf = []
    # lastrecordpdf = []
    # programnamepdf = []
    # for x in allfilename:
        # new = [i[13] for i in allfilename]
        # filenamepdf.append(new)
    #     new = [i[6] for i in allfilename]
        # txtfilenamepdf.append(new)
    #     new = [i[9] for i in allfilename]
    #     totalaccpdf.append(new)
    #     new = [i[-4] for i in allfilename]
    #     totalpagespdf.append(new)
    #     new = [i[-2] for i in allfilename]
    #     lastrecordpdf.append(new)
    #     new = [i[16] for i in allfilename]
    #     firstrecordpdf.append(new)
    #     new = [i[-1] for i in allfilename]
    #     programnamepdf.append(new)
        # filenamepdf = [element for nestedlist in filenamepdf for element in nestedlist]
    #     txtfilenamepdf = [element for nestedlist in txtfilenamepdf for element in nestedlist]
    #     totalaccpdf = [element for nestedlist in totalaccpdf for element in nestedlist]
    #     totalpagespdf = [element for nestedlist in totalpagespdf for element in nestedlist]
    #     lastrecordpdf = [element for nestedlist in lastrecordpdf for element in nestedlist]
    #     firstrecordpdf = [element for nestedlist in firstrecordpdf for element in nestedlist] 
    #     programnamepdf = [element for nestedlist in programnamepdf for element in nestedlist]

        # filenamepdf = ''.join([str(ele) for ele in filenamepdf])
    #     txtfilenamepdf = ' '.join([str(elem) for elem in txtfilenamepdf])
    #     totalaccpdf = ' '.join([str(elem) for elem in totalaccpdf])
    #     totalpagespdf = ' '.join([str(elem) for elem in totalpagespdf])
    #     lastrecordpdf = ' '.join([str(elem) for elem in lastrecordpdf])
    #     firstrecordpdf = ' '.join([str(elem) for elem in firstrecordpdf])
    #     programnamepdf = ' '.join([str(elem) for elem in programnamepdf ])

        

        time1 = datetime.now() 
        start_time=time1.strftime("%H:%M:%S")
        date1 = time1.strftime("%d/%b/%Y")
        file=open(f"E:\\IWOC\\LOG_FILES\\{ffname}.txt","a")
        file.write("INTERCITY  MPC (M) Sdn. Bhd.\n")
        file.write("FIRST REMINDER\n")
        file.write("=============================================================\n")
        file.write("Program ID           : FIRST REMINDER.PDS\n")
        file.write("Paper Type           :\n")
        file.write("File Name (PDF)      : ")
        file.write('PGI.pdf')
        file.write("\n")
        file.write("Data Name (TXT)      : ")
        file.write(str(files1))
        file.write("\n")
        file.write("Total Accounts       :        ")
        file.write(str(total_clients))
        file.write("\n")
        file.write("Total Pages          :        ")
        file.write(str(total_clients))
        file.write("\n")
        file.write("Total Impression     :        ")
        file.write(str(totalpages))
        file.write("\n")

        file.write("First Record         : ")
        file.write(str(first_record))
        file.write("\n")
        file.write("Last Record          : ")
        # last_record=df['ILOM_OLD_SEQUENCE'].iat[-1]
        file.write(str(last_record))
        file.write("\n")
        file.write("=============================================================\n")
        file.write("CUSTOMER NAME        : BSN                                   \n")
        file.write("=============================================================\n")
        file.write("Date & Time processed: ")


        file.write(str(date1))
        file.write("    ")
        file.write(str(start_time))
        file.write("    " )
        time1 = datetime.now()
        end_time=time1.strftime("%H:%M:%S")
        file.write("    ")
        file.write(str(end_time) )
        file.write('\n\n\n\n')
        file.close()
        return file


#