import os
from fpdf import FPDF
import PyPDF2
import datetime
from fpdf import FPDF
from datetime import datetime
from datetime import datetime
from barcode import EAN13,Code39
from barcode.writer import ImageWriter
from fpdf import FPDF
from PIL import Image

def BBA(request,files1):
    move_files = []
    move_files.append(files1)
    path11=r'E:\IWOC\SOURCE\\'+ files1


    with open(path11) as file:
        line = file.readlines()
        total_clients = len(line)
        print(total_clients)

    with open(path11) as file:
        IBS = []
        date = []
        Account_no = []
        Rev_eff_rate_from = []
        Rev_eff_rate_to = []
        Default_rate = []
        Effective_rate_of_revision = []
        Client_name = []
        Barcode_below = []
        Add1 = []
        Add2 = []
        Add3 = []
        Add4 = []
        Add5 = []
        for line in file:
            IBS.append(line[354:380].strip())
            date.append(line[15:25])
            Account_no.append(line[0:15])
            Rev_eff_rate_from.append(line[265:279])
            Rev_eff_rate_to.append(line[279:294])
            Default_rate.append(line[335:349])
            Effective_rate_of_revision.append(line[325:335])
            Client_name.append(line[25:57].strip())
            Barcode_below.append(line[349:354])
            Add1.append(line[65:87].strip())
            Add2.append(line[105:143].strip())
            Add3.append(line[145:181].strip())
            Add4.append(line[185:222].strip())
            Add5.append(line[225:254].strip())
        
        # for i in date:
        dates = (datetime.strptime(ts, '%d/%m/%Y') for ts in date)
        date_strings = [datetime.strftime(d, '%d%m%y') for d in dates]

        def barcode(Account_no,count):
            barcode_no = '0131' +  str(date_strings[count]) +  '02' + str(Account_no[count]) 
            # print(barcode_no,'-------------')
            # print("date :",date) 
        
            number = barcode_no
            my_code = Code39(number, writer=ImageWriter())
            new_code  =  str(count)
            my_code.save(r"E:\IWOC\EXTRA\bba_barcode\%s" %(new_code))

            from PIL import Image

            img = Image.open(r"E:\IWOC\EXTRA\bba_barcode\%s.png" %(new_code))
            imgCropped = img.crop(box=(4,10,1100,200))
            imgCropped.save(r"E:\IWOC\EXTRA\bba_barcode\%s.png" %(new_code))
        
    count = 0
    c = 1
    d = 1
    pdf = FPDF()
    for i in range(0,total_clients):
        barcode(Account_no,count)
        pdf.add_page()
        pdf.set_auto_page_break(auto=False)
        pdf.set_line_width(5)
        pdf.set_draw_color(r=0, g=0, b=0)
        pdf.line(x1=5, y1=4, x2=205, y2=4)         #1st horizontal line
        pdf.line(x1=5, y1=292, x2=205, y2=292)     #2nd horizontal line 
        pdf.line(5,6,5,290)                         # 1st vertical line
        pdf.line(205,6,205,290)                     # 2nd vertical line    

        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MiB-02.jpg',30,15,45)
        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(85)
        pdf.cell(10,50,txt = IBS[i])
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,65,txt = 'Tarikh')
        pdf.cell(2)
        pdf.cell(10,65,txt = ':')
        pdf.cell(-6)
        pdf.cell(10,65,txt = date[i])
        pdf.set_font('Arial','B',9)
        pdf.cell(130)
        pdf.cell(10,65,txt = 'Sulit & Persendirian',align='R')
        pdf.ln(1)
        pdf.set_font('Arial','I',9)
        pdf.cell(10)
        pdf.cell(10,70,txt = 'Date')
        pdf.set_font('Arial','BI',9)
        pdf.cell(149)
        pdf.cell(10,70,txt = 'Private & Confidential',align='R')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,85,txt = 'Tuan / Puan,')

        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,95,txt = 'Sir / Madam,')

        pdf.set_font('Arial','B',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,102,txt = 'PINDAAN TERMA & SYARAT KEMUDAHAN PEMBIYAAN')
        pdf.set_font('Arial','BI',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,107,txt = 'REVISION OF TERMS AND CONDITIONS OF FINANCING FACILITY')


        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,120,txt = 'No. Akaun')
        pdf.set_font('Arial','B',9)
        pdf.cell(42)
        pdf.cell(10,120,txt = ':')
        pdf.set_font('Arial','',9)
        pdf.cell(-4)
        pdf.cell(10,120,txt = Account_no[i])
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,125,txt = 'Account No.')



        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,138,txt = 'Pindaan Kadar Keuntungan Efektif')
        pdf.set_font('Arial','B',9)
        pdf.cell(42)
        pdf.cell(10,138,txt = ':')
        pdf.set_font('Arial','',9)
        pdf.cell(-4)
        pdf.cell(10,138,txt = 'Dari')
        pdf.cell(-3)
        pdf.cell(10,138,txt = Rev_eff_rate_from[i]+'%')
        pdf.cell(28)
        pdf.cell(10,138,txt = 'Kepada')
        pdf.cell(3)
        pdf.cell(10,138,txt = Rev_eff_rate_to[i] + '%')
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,143,txt = 'Revision of Effective Profit Rate')
        pdf.cell(48)
        pdf.cell(10,143,txt = 'From')
        pdf.cell(35)
        pdf.cell(10,143,txt = 'To')

        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,155,txt = 'Kadar Kemungkiran')
        pdf.set_font('Arial','B',9)
        pdf.cell(42)
        pdf.cell(10,155,txt = ':')
        pdf.set_font('Arial','',9)
        pdf.cell(-4)
        pdf.cell(10,155,txt = Default_rate[i])
        pdf.ln(1) 
        pdf.set_font('Arial','I',9)
        pdf.cell(10)
        pdf.cell(10,160,txt = 'Default Rate ')


        pdf.set_font('Arial','',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,173,txt = 'Tarikh Pindaan Berkuatkuasa')
        pdf.set_font('Arial','B',9)
        pdf.cell(42)
        pdf.cell(10,173,txt = ':')
        pdf.set_font('Arial','',9)
        pdf.cell(-4)
        pdf.cell(10,173,txt = Effective_rate_of_revision[i])
        pdf.ln(1) 
        pdf.set_font('Arial','I',9)
        pdf.cell(10)
        pdf.cell(10,178,txt = 'Effective Date of Revision')


        pdf.set_line_width(0.01)
        pdf.line(21,118,190,118)

        pdf.ln(1)
        pdf.set_font('Arial','',9.3)
        pdf.cell(10)
        pdf.cell(10,190,txt = 'Kami   ingin   memaklumkan   bahawa   berikutan   dengan   penamatan   perkhidmatan   anda   dengan   Bank   dan')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,194,txt = 'berdasarkan  kepada  terma  dan  syarat  kemudahan  pembiayaan  / dokumen  sekuriti  di  antara  anda  dan  Bank,')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,198,txt = 'terma  dan  syarat  kemudahan  pembiayaan  anda  telah  dipinda  seperti  yang  dinyatakan  di  atas. ')
        pdf.set_font('Arial','I',9.1)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,202,txt = 'We  wish  to  advise  that  following  the  cessation  of  your  employment  with  the  Bank  and  in  accordance  with  the')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,206,txt = 'terms  and  conditions  of  the  facility  /  security  documents  between  you  and  the  Bank,  the  terms  and  conditions')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,210,txt = 'of your facility have been revised as stated above. ')


        pdf.ln(1)
        pdf.set_font('Arial','',9.3)
        pdf.cell(10)
        pdf.cell(10,223, txt = 'Bagi   pembiayaan   yang   dikeluarkan   secara   progresif,  bayaran   keuntungan   bulanan   perlu  dijelaskan   pada')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,227,txt = 'akhir  bulan  manakala   bagi  pembiayaan   yang  telah   dikeluarkan  sepenuhnya,  ansuran  bulanan  perlu   dibayar')
        pdf.ln(1)
        pdf.set_font('Arial','',9.3)
        pdf.cell(10)
        pdf.cell(10,231,txt='pada   1   haribulan   pada   setiap   bulan.   Sekiranya   anda   gagal   menjelaskan   bayaran    keuntungan  sebelum')
        pdf.ln(1)
        pdf.set_font('Arial','',9.3)
        pdf.cell(10)
        pdf.cell(10,235,txt = 'permulaan   bayaran    ansuran   bulanan   atau  bayaran  ansuran   bulanan   sebanyak  tiga   (3)   kali,  pihak   Bank')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,239,txt = "berhak   menukar   kadar   keuntungan   efektif  tersebut  kepada " +   Default_rate[i]   +  " setahun    atau    kadar   lain   yang")
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,243,txt = 'mungkin   ditetapkan   oleh   Bank   dari   semasa   ke  semasa.  Segala   terma   dan  syarat   lain  bagi   kemudahan')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,247,txt = 'pembiayaan  ini  adalah  tidak  berubah.')
        pdf.set_font('Arial','I',9)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,251,txt = 'For   facility  under   progressive  release,  payment  of   the   monthly   profit  must   be  serviced  by   end  of  the  month')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,255,txt = 'while  for  fully  released  facility,  payment  of  the  monthly  instalment  is  due  on  the  1st  day  of  each  month.  In  the')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,259,txt = 'event   you  default  on  three  (3)   payments  of   the  monthly  profit   pending  the  commencement   of  the   instalment')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,263,txt = 'payment,  or  three   (3)  payments   of   the  monthly  instalment,  the   Bank  shall  be  entitled   to  convert  the  effective')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,268,txt = "profit   rate   to  " + Default_rate[i] + ' per   annum  or  such  other   rate  the  Bank  may  prescribe  from  time  to  time.  All')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,272,txt = 'other  terms  and  conditions  of  your  financing  facility  shall  remain  unchanged. ')


        pdf.set_font('Arial','',9.3)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,285,txt = 'Pihak  Bank   berhak   mengubah   bilangan  dan   jumlah   ansuran   anda   tetapi  perubahan  tersebut   tidak   akan')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,289,txt = 'menyebabkan  jumlah  ansuran  melebihi  Harga  Jualan  Bank,  di  mana  berkenaan.')
        pdf.set_font('Arial','I',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,293,txt = 'The  Bank  has  the  discretion  to  vary  your  number  and  amount  of  instalment  however  such  variation  shall  not')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,297,txt = "result  in  the  total  instalment  amount  exceeding  the  Bank's  Sale  Price, where applicable.")


        pdf.set_font('Arial','',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,310,txt = 'Bagi  membolehkan  pihak   Bank   menilai  semula  kemudahan  pembiayaan  anda,  sila  hubungi   cawangan   Bank')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,314,txt = 'di  mana  akaun  anda  diselenggarakan  untuk  mengemaskinikan  butir  pekerjaan  dan  pendapatan  anda.')
        pdf.set_font('Arial','I',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,318,txt = 'To   enable   the   Bank   to   reassess   your  financing  facility,  please   visit   your  home   branch  and  provide  your')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,322,txt = 'current   employment   and   income   details')

        pdf.set_font('Arial','',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,335,txt = 'Perubahan   yang   tersebut   di   atas   adalah   tanpa   menjejaskan   hak   pihak   Bank  untuk  mengkaji  semula  dan')
        pdf.set_font('Arial','',9.4)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,339,txt = 'mengubah   margin   pembiayaan  dan/  atau  menamatkan   kemudahan   tersebut    selaras    dengan    penamatan')
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,343,txt = 'perkhidmatan  anda  dengan  pihak  Bank.')
        pdf.set_font('Arial','I',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,347,txt = "The  above  revision  is  without  prejudice  to   the  Bank's  rights  to  review   and  revise   the  margin  of  financing  of")
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,352,txt = 'the  facility  and/ or  to  terminate  the  facility  following  the  cessation  of  your  employment  with  the  Bank.')

        pdf.set_font('Arial','',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,365,txt = 'Terima Kasih.')
        pdf.set_font('Arial','I',9.2)
        pdf.ln(1)
        pdf.cell(10)
        pdf.cell(10,369,txt = "Thank you.")

        pdf.ln(1)
        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MiB-03.jpg',10,245,15)


        pdf.set_font('Arial','',9.2)
        pdf.ln(1)
        pdf.cell(20)
        pdf.cell(10,383,txt = 'Ini adalah cetakan komputer, tandatangan tidak diperlukan.')
        pdf.set_font('Arial','I',9.2)
        pdf.ln(1)
        pdf.cell(20)
        pdf.cell(10,388,txt = "This is computer generated, no signature is required.")


        pdf.set_font('Arial','',9.2)
        pdf.ln(1)
        pdf.cell(140)
        d  = str(d)
        if len(d) == 1:
            count_value = '00000' + str(d)
        if len(d) == 2:
            count_value = '0000' + str(d)
        if len(d) == 3:
            count_value = '000' + str(d)
        if len(d) == 4:
            count_value = '00' + str(d)
        if len(d) == 5:
            count_value = '0' + str(d)
        if len(d) == 6:
            count_value =  str(d)
            pdf.cell(10,430, txt  = str(count_value))
        pdf.set_font('Arial','',7)
        pdf.cell(-10)
        pdf.cell(10,430,txt = files1 +'/' + str(count_value))
        d = int(d)
        d += 1

        ############################second page


        pdf.add_page()
        count = int(count)
        pdf.set_font('Arial','',8)
        pdf.cell(-4)
        pdf.cell(2,332,txt = '00' + str(c) )
        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MiB-01 - Copy.jpg',2, 5,210)
        pdf.ln(1)
        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MBB_ASB-02 - Copy.jpg',15, 178,50)
        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MiB-03.jpg',110,165,15)
        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MiB-04.jpg',16,165,50)
        pdf.image(r'E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\MBB_ASB-02 - Copy (2).jpg',140,160,65)

        pdf.cell(60)
        pdf.cell(10,416,txt = '0000'+ str(c))

        new_code = str(count)
        pdf.image(r'E:\IWOC\EXTRA\bba_barcode\%s.png'%(new_code),15,190,40) 
        count = int(count)
        count += 1
        
        pdf.ln(1)
        pdf.cell(5)
        pdf.cell(10,374,txt=Barcode_below[i])

        
        h = 420
        pdf.ln(1)
        pdf.set_font('Arial','',10)
        pdf.cell(60)
        pdf.cell(10,h,txt = Client_name[i])

        if len(Add1[i]) > 1:
            pdf.ln(1)
            pdf.cell(60)
            h += 6
            pdf.cell(10,h,txt = Add1[i])


        pdf.ln(1)
        pdf.cell(60)   
        h += 6 
        pdf.cell(10,h ,txt = Add2[i])
        
        c += 1
        
        pdf.ln(1)
        pdf.cell(60)
        h += 6
        pdf.cell(10,h,txt = Add3[i])

        pdf.ln(1)
        pdf.cell(60)
        h += 6
        pdf.cell(10,h,txt = Add4[i])

        pdf.ln(1)
        pdf.cell(60)
        h += 6
        pdf.cell(10,h,txt = Add5[i])


    pdf.output('E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\BBA.pdf')

    path20 = r'E:\IWOC\SOURCE' 
    target = r'E:\IWOC\TEMP' 

    for i in move_files:

        src_path = os.path.join(path20, i)
        dst_path = os.path.join(target, i)
        os.rename(src_path, dst_path)
    print('Files Moved................')


import PyPDF2

pdf_in = open(r"E:\IWOC\APP SCRIPT\BBA\CYCLE\ML0184I\DATA\BBA.pdf", 'rb')
pdf_reader = PyPDF2.PdfFileReader(pdf_in)
pdf_writer = PyPDF2.PdfFileWriter()

for pagenum in range(pdf_reader.numPages):
    page = pdf_reader.getPage(pagenum)
    if pagenum % 2:
        page.rotateClockwise(180)
    pdf_writer.addPage(page)

pdf_out = open(r'E:\IWOC\OUTPUT\BBA\BBA111.pdf', 'wb')
pdf_writer.write(pdf_out)
pdf_out.close()
pdf_in.close() 