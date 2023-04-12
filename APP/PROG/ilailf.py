from fpdf import FPDF
import os

def ILAILF(request,files1):
    move_files = []
    move_files.append(files1)
    path11=r'E:\IWOC\SOURCE\\'+ files1

    with open(path11) as file:
        line = file.readlines()
        total_clients = len(line)
        print(total_clients)

    with open(path11) as file:
        Name = []
        Add1 = []
        Add2 = []
        Add3 = []
        Add4 = []
        Date = []
        Accout_No = []
        Profit_Amount = []
        Effective_Rate =[]
        Financing_Balance = []
        Rollover_First_Date = []
        Rollover_Second_Date = []
        Number_of_Days = []
        Profit_Payment_Due_Date = []
        for line in file:
            Name.append(line[0:24])
            Add1.append(line[40:80].strip())
            Add2.append(line[80:119].strip())
            Add3.append(line[120:156].strip())
            Add4.append(line[160:196].strip())
            Date.append(line[200:211])
            Accout_No.append(line[210:223])
            Profit_Amount.append(line[224:239].strip())
            Effective_Rate.append(line[240:245].strip())
            Financing_Balance.append(line[248:261].strip())
            Rollover_First_Date.append(line[261:272])
            Rollover_Second_Date.append(line[285:296])
            Number_of_Days.append(line[273:275])
            Profit_Payment_Due_Date.append(line[275:286])

    count = 1
    pdf = FPDF()
    for i in range(0,total_clients):
        pdf.add_page()
        pdf.set_auto_page_break(auto=False)
        pdf.set_line_width(6)
        pdf.set_draw_color(r=0, g=0, b=0)
        pdf.line(x1=5, y1=4, x2=205, y2=4)         #1st horizontal line
        pdf.line(x1=5, y1=292, x2=205, y2=292)     #2nd horizontal line 
        pdf.line(5,6,5,290)                         # 1st vertical line
        pdf.line(205,6,205,290)                     # 2nd vertical line    

        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MiB-02.jpg',30,10,50)
        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(166)
        pdf.cell(10,27,txt = 'Sulit & Persendirian',align='R')
        pdf.ln(1)
        pdf.cell(10)
        pdf.set_font('Arial','I',9)
        pdf.cell(158)
        pdf.cell(10,31,txt = 'Private & Confidential',align='R')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,80,txt = 'Tarikh : ')
        pdf.cell(2)
        pdf.cell(10,80,txt = str(Date[i]))
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,84,txt = 'Date ')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,94,txt = 'Tuan/Puan,')
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,98,txt = 'Sir/Madam,')

        pdf.set_font('Arial','B',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,108,txt = 'NOTIS BAYARAN JUMLAH KEUNTUNGAN')
        pdf.set_font('Arial','BI',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,112,txt = 'NOTICE ON PAYMENT OF PROFIT')


        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,122,txt = 'No. Akaun')
        pdf.cell(50)
        pdf.cell(10,122,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,122,txt = str(Accout_No[i]))
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,126,txt = 'Account No.')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,136,txt = 'Anggaran Amaun Keuntungan')
        pdf.cell(50)
        pdf.cell(10,136,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,136,txt = 'RM')
        pdf.cell(-4)
        pdf.cell(10,136,txt = Profit_Amount[i])
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,140,txt = 'Forecast Profile Amount')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,150,txt = 'Kadar Keuntungan Efektif')
        pdf.cell(50)
        pdf.cell(10,150,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,150,txt = str(Effective_Rate[i]) + '% tahunan')
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,154,txt = 'Effective Rate')
        pdf.cell(65)
        pdf.cell(10,154,txt= 'p.a')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,164,txt = 'Baki Pembiayaan')
        pdf.cell(50)
        pdf.cell(10,164,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,164,txt = 'RM')
        pdf.cell(-2)
        pdf.cell(10,164,txt = Financing_Balance[i])
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,168,txt = 'Financing Balance')


        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,178,txt = 'Tarikh Kadar Lanjutan')
        pdf.cell(50)
        pdf.cell(10,178,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,178,txt = 'Dari')
        pdf.cell(-2)
        pdf.cell(10,178,txt = Rollover_First_Date[i])
        pdf.cell(7)
        pdf.cell(10,178,txt = 'ke')
        pdf.cell(-5)
        pdf.cell(10,178,txt = Rollover_Second_Date[i])
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,182,txt = 'Rollover Date')
        pdf.cell(54)
        pdf.cell(10,182,txt = 'From')
        pdf.cell(13)
        pdf.cell(10,182,txt = 'to')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,192,txt = 'Bilangan Hari')
        pdf.cell(50)
        pdf.cell(10,192,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,192,txt = str(Number_of_Days[i]) + ' hari')
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,196,txt = 'Number of Days')
        pdf.cell(59)
        pdf.cell(10,196,txt = 'days')


        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,206,txt = 'Tarikh Matang Bayaran Keuntungan')
        pdf.cell(50)
        pdf.cell(10,206,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,206,txt = Profit_Payment_Due_Date[i])
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,210,txt = 'Profit Payment Due Date')

        pdf.set_line_width(0.01)
        pdf.line(13,142,195,142)

        pdf.ln(1)
        pdf.set_font('Arial','',8.5)
        pdf.cell(2)
        pdf.cell(10,225 ,txt = 'Kami ingin memaklumkan bahawa jumlah bayaran keuntungan adalah separti yang dinyatakan diatas. Pembayaran boleh dibuat di mana- ')
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,229,txt = 'mana cawangan Maybank.')
        pdf.ln(1)
        pdf.set_font('Arial','I',8.5)
        pdf.cell(2)
        pdf.cell(10,233,txt = 'We wish to inform that the total profit due and payable as specified above. Payment can be made at any of our Maybank branches.')


        pdf.ln(1)
        pdf.set_font('Arial','',8.65)
        pdf.cell(2)
        pdf.cell(10,244 ,txt = 'Jika bayaran melalui pemindahan antara bank, sila nyatakan nombor akaun dan nama syarikat tuan. Jika bayaran dibuat melalui arahan ')
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,248,txt = 'tetap, pastikan akaun mempunyai baki  yang mencukupi.')
        pdf.ln(1)
        pdf.set_font('Arial','I',8.65)
        pdf.cell(2)
        pdf.cell(10,252,txt = "For payment via Interbank Transfer, please indicate the above said account number and Compnay's" ' name accordingly. For payment via')
        pdf.set_line_width(0.01)
        pdf.line(105,165,155,165)
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,256,txt = 'Standing Instruction. please ensure there is sufficient funds in your account for auto-debiting.')


        pdf.ln(1)
        pdf.set_font('Arial','',8.5)
        pdf.cell(2)
        pdf.cell(10,266 ,txt = 'Nota: Anggaran amaun keuntungan di atas adalah berdasarkan tiada transaksi tambahan ke atas akaun dan jadar keuntungan efektif yang')
        pdf.ln(1)
        pdf.cell(2)
        pdf.set_font('Arial','',8.6)
        pdf.cell(10,270,txt = 'dikenakan atas akaun adalah sama. Jika tarikh matang bayaran keuntungan jatuh pada hujung minggu atau cuti umum, sila buat bayaran')
        pdf.ln(1)
        pdf.set_font('Arial','',8.6)
        pdf.cell(2)
        pdf.cell(10,274,txt = 'sebelum tarikh tersebut.')
        pdf.ln(1)
        pdf.set_font('Arial','I',8.5)
        pdf.cell(2)
        pdf.cell(10,278,txt = "Note: The  above  forecast  profit  amount  show  is  on  the  expectations  that  there  are  no  further  transaction  to  the  account  and  the")
        pdf.ln(1)
        pdf.cell(2)
        pdf.set_font('Arial','I',8.1)
        pdf.cell(10,282,txt = 'effective  rate  charged  to  the  account  is  the  same  throughout  the  tenor. Should  the  due  date  fall  on  a  weekend  or  public  holiday, kindly')
        pdf.ln(1)
        pdf.cell(2)
        pdf.cell(10,286,txt = 'arrange to pay prior to the due date.')



        pdf.ln(1)
        pdf.set_font('Arial','',8.5)
        pdf.cell(2)
        pdf.cell(10,296 ,txt = 'Sila maklumkan kepada pihak Bank dengan SEGERA sekiranya terdapat apa-apa perubahan alamat dan nombor telefon.')
        pdf.ln(1)
        pdf.set_font('Arial','I',8.5)
        pdf.cell(2)
        pdf.cell(10,300,txt = "Kindly notify the Bank IMMEDIATELY if there is any change of address and telephone number.")


        pdf.ln(1)
        pdf.set_font('Arial','',8.5)
        pdf.cell(2)
        pdf.cell(10,310 ,txt = 'Terima kasih kerana menggunkan perkhidmatan perbankan kami.')
        pdf.ln(1)
        pdf.set_font('Arial','I',8.5)
        pdf.cell(2)
        pdf.cell(10,314,txt = "Thank you for banking with us.")


        pdf.ln(1)
        pdf.set_font('Arial','',8.5)
        pdf.cell(2)
        pdf.cell(10,324 ,txt = 'Ini adalah cetakan komputer. Tandatangan tidak diperlukan')
        pdf.ln(1)
        pdf.set_font('Arial','I',8.5)
        pdf.cell(2)
        pdf.cell(10,328,txt = "This is a computer generated letter. No signature required.")

        pdf.ln(1)
        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MiB-03.jpg',10,265,15)
        pdf.ln(1)
        pdf.cell(160)
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
            pdf.cell(10,419, txt  = str(count_value))
        pdf.cell(-25)
        pdf.cell(10,450,txt = files1 + '-' +str(count_value))

        pdf.add_page()
        count = int(count)
        pdf.set_font('Arial','',8)
        pdf.cell(-4)
        pdf.cell(2,332,txt = '00' + str(count) )
        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MiB-01 - Copy.jpg',2, 5,210)
        pdf.ln(1)
        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MBB_ASB-02 - Copy.jpg',25, 180,50)
        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MiB-03.jpg',110,165,20)
        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MiB-04.jpg',16,165,50)
        pdf.image(r'E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\MBB_ASB-02 - Copy (2).jpg',140,160,65)

        pdf.cell(80)
        pdf.cell(10,416,txt = '0000'+ str(count))

        count += 1

        pdf.ln(1)
        pdf.set_font('Arial','',10)
        pdf.cell(80)
        pdf.cell(10,420,txt = Name[i])

        pdf.ln(1)
        pdf.cell(80)
        pdf.cell(10,426,txt = Add1[i])

        pdf.ln(1)
        pdf.cell(80)    
        pdf.cell(10,432,txt = Add2[i])

        pdf.ln(1)
        pdf.cell(80)
        pdf.cell(10,438,txt = Add3[i])

        pdf.ln(1)
        pdf.cell(80)
        pdf.cell(10,444,txt = Add4[i])



    pdf.output('E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\ML0184.pdf')


    
    path20 = r'E:\IWOC\SOURCE' 
    target = r'E:\IWOC\TEMP' 

    for i in move_files:

        src_path = os.path.join(path20, i)
        dst_path = os.path.join(target, i)
        os.rename(src_path, dst_path)
    print('Files Moved................')


import PyPDF2

pdf_in = open(r"E:\IWOC\APP SCRIPT\ILA_ILF\CYCLE\ML0184I\DATA\ML0184.pdf", 'rb')
pdf_reader = PyPDF2.PdfFileReader(pdf_in)
pdf_writer = PyPDF2.PdfFileWriter()

for pagenum in range(pdf_reader.numPages):
    page = pdf_reader.getPage(pagenum)
    if pagenum % 2:
        page.rotateClockwise(180)
    pdf_writer.addPage(page)

pdf_out = open(r'E:\IWOC\OUTPUT\ILAILF\ML111.pdf', 'wb')
pdf_writer.write(pdf_out)
pdf_out.close()
pdf_in.close()   

