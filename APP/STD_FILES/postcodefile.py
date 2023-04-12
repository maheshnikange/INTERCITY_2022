def STD_GROUP_CODE(List_of_Postcodes):
    group_code=[]
    for i in List_of_Postcodes:
        i = int(i)
        if i>=1000 and i<=2999:
            group_code.append('001')
        elif i>=5000 and i<=6999:
            group_code.append('002')
        elif i>=7000 and i<=7999:
            group_code.append('003')
        elif i>=8000 and i<=8099:
            group_code.append('004')
        elif (i>=8100 and i<=8899) or (i>=9100 and i<=9599) or (i>=9700 and i<=9899) or i==34950:
            group_code.append('005')
        elif i>=9000 and i<=9099 or (i>=9600 and i<=9699):
            group_code.append('006')
        elif i>=10000 and i<=10999 or (i>=11000 and i<=1150) or i==11800 or (i>=13100 and i<=13310) or i==13050 or (i>=14100 and i<=14390):
            group_code.append('007')
        elif i==11600 or i==11700:
            group_code.append('008')
        elif i>=11900 and i<=11999:
            group_code.append('009')
        elif i>=12000 and i<=12300 or (i>=12700 and i<=12990) or (i>=13000 and i<=13020) or i==13400 or i==13800:
            group_code.append('010')

        # -----------11-20---------------------
        elif i == 13500 or i == 13600 or i == 13700:
            group_code.append('011')
        elif i>=14000 and i<= 14020 or i == 14400:
            group_code.append('012')
        elif i >= 30000 and i <= 30999:
            group_code.append('013')
        elif i == 31450 or i == 31500 or i == 31650:
            group_code.append('014')
        elif (i >= 31000 and i <= 31100) or i == 31200 or i == 31300 or i == 31350 or (i >= 31550 and i <=31610) or (i >=31700 and i <=32900) or (i >= 33100 and i <= 33500) or (i >= 34100 and i <= 34500) or i == 34800 or i == 34900 or (i >= 35000 and i <= 35999) or (i >= 39000 and i <= 39999):
            group_code.append('015')
        elif i == 31400 or i == 31150 or i ==  31250:
            group_code.append('016')
        elif  (i >= 33000 and i <= 33099) or i == 33600 or i == 33700:
            group_code.append('017') 
        elif (i >= 34000 and i <= 34099) or i == 34600 or i == 34650  or i == 34700 or i == 34750 or i == 34850:
            group_code.append('018')
        elif (i >= 36000 and i <= 36999):
            group_code.append('019')        
        elif i == 40150:
            group_code.append('020')

        # -----------21-30---------------------
        elif i == 40400 or (i>= 40460 and i <= 40690) or i == 42450:
            group_code.append('021')
        elif i == 40000 or i == 40100 or i == 40200 or i == 40300 or i == 40450:
            group_code.append('022')
        elif (i >= 40700 and i <= 40990):
            group_code.append('023')
        elif i == 41000  or i == 41100 or i == 41200 or i == 41250 or (i >= 41500 and i <= 41990):
            group_code.append('024')
        elif i == 40170 or i == 41050 or i == 41150 or i == 41300 or i == 41400 or i == 42100:
            group_code.append('025')
        elif (i >= 42000 and i <=42009) or i == 42920:
            group_code.append('026')
        elif i == 42600 or i == 42700 or (i >=45000 and i <= 45030) or i == 45100:
            group_code.append('027')
        elif i == 42500 or i == 42800 or i == 45800 or (i >= 44000 and  i <= 44300) or i == 48100 or i == 42940 or i == 42960 or i == 45200 or i == 45300 or i == 45400 or i == 45500 or i == 45600 or i == 45700:
            group_code.append('028')
        elif i == 43900 or i == 43950 or i == 64000:
            group_code.append('029')
        elif (i >= 43000 and i <= 43009) or (i >= 43600 and i <= 43650) or (i >= 62000 and i <= 62306):
            group_code.append('030')

        # -----------31-40---------------------
        elif i == 43100 or (i >= 43200 and i <= 43207):
            group_code.append('031')
        elif (i >= 43300 and i <= 43499):
            group_code.append('032')
        elif i == 43500 or i == 43700:
            group_code.append('033')
        elif (i >= 46000 and i <= 46999):
            group_code.append('034')
        elif i == 40160 or i == 47000 or i == 47810 or i == 48050:
            group_code.append('035')
        elif (i >= 47100  and i <= 47190) or i == 42610:
            group_code.append('036')
        elif i == 47200 or (i >= 47300 and i <= 47308):
            group_code.append('037')
        elif i == 47400 or i == 47410 or i == 47800 or i == 47820 or i == 47830:
            group_code.append('038')
        elif (i >= 47500 and i <= 47699):
            group_code.append('039')
        elif i == 48000 or i == 48010 or i == 48020 or (i >= 48200 and i <= 48399):
            group_code.append('040')


        # -----------41-50---------------------
        elif (i >= 62400 and i <= 62688) or i == 62692:
            group_code.append('041')
        elif (i >=42200 and i <= 42300):
            group_code.append('042')
        elif (i >= 43800 and i <= 43809) or (i >= 63000 and i <= 63300):
            group_code.append('043')
        elif (i >= 50000 and i <= 50400):
            group_code.append('044')
        elif (i >= 50500 and i <= 50699):
            group_code.append('045')
        elif (i >= 50450 and i <= 50490):
            group_code.append('046')
        elif (i >= 50700 and i <= 50990):
            group_code.append('047')
        elif (i >= 51000 and i <= 51200) or ( i >= 51700 and i <= 51990):
            group_code.append('048')
        elif i == 52000 or ( i >= 52100 and i <= 52200):
            group_code.append('049')
        elif (i >= 53000 and i <= 53300) or (i >= 53700 and i <= 53990):
            group_code.append('050')


        # -----------51-60---------------------
        elif (i >= 54000 and i <= 54200):
            group_code.append('051')
        elif (i >= 55000 and i <= 55300) or (i >= 55700 and i <= 55990):
            group_code.append('052')
        elif (i >= 56000 and i <= 56100):
            group_code.append('053')
        elif (i >= 58000 and i <= 58200) or (i >= 58700 and i <= 58990):
            group_code.append('054')
        elif (i >= 590000 and i <= 59200)  or (i >= 59700 and i <= 59900):
            group_code.append('055')
        elif (i >= 57000 and i <= 57100) or (i >= 57700 and i <= 57990):
            group_code.append('056')
        elif i == 60000:
            group_code.append('057')    
        elif i == 68000:
            group_code.append('058')
        elif i == 68100:
            group_code.append('059')
        elif i == 70000 or (i >= 70100 and i <= 70200) or (i >= 70500  and i <= 70990) or i == 71450 or i == 71770:
            group_code.append('060')


        # -----------61-70---------------------                 
        elif i == 70300 or(i >= 70400 and i <= 70450):
            group_code.append('061')
        elif (i >=71000 and i <=71009) or i == 71010 or i == 72100 or i == 71760 or (i >= 71800 and i <= 71809) or i == 71900 or (i >= 71950 and i <= 71960) or (i >= 72000 and i <= 72009) or (i >=71500 and i <= 71550) or i == 72500 or (i>=73100 and i <= 73500):
            group_code.append('062')
        elif (i >= 71050 and i <= 71409) or (i >= 71600 and i <= 71759) or (i >= 72120 and i <= 72400) or (i >= 73200 and i <= 73400):
            group_code.append('063')
        elif (i >= 75000  and i <= 75260):
            group_code.append('064')
        elif i == 75460 or i == 76100 or i == 76400 or i == 76450:
            group_code.append('065')
        elif i == 77500 or i == 76200 or i == 76300 or i == 78000 or i == 78100 or i == 78200 or i == 78300:
            group_code.append('066')
        elif (i >= 80000  and i <= 80999):
            group_code.append('067')
        elif (i >= 81100 and i <=  81109):
            group_code.append('068')
        elif (i >= 81200  and i <= 81299):
            group_code.append('069')
        elif i == 81300 or i == 81310 or i == 81550 or i == 81560 or i == 81110 or i == 79000 or i == 79100 or i == 79150 or i == 79200 or i== 79250:
            group_code.append('070')

        # -----------71-80--------------------- 
        elif (i >= 81700  and i <= 81799):
            group_code.append('071')
        elif (i >= 81000 and i <= 81099) or (i >= 81400 and i <= 81499):
            group_code.append('072')
        elif (i >= 81600 and i <= 81699) or i == 81800 or (i >= 81900 and i <= 81990):
            group_code.append('073')
        elif(i >= 82000  and i <= 82099) or i == 81500:
            group_code.append('074')
        elif (i >= 83000 and i <=  83999 ) or i ==  86400:
            group_code.append('075')
        elif (i >= 84000  and i  <=  84999):
            group_code.append('076')
        elif (i >= 85000 and i <= 85099) or i == 86500:
            group_code.append('077')
        elif (i >= 86000 and i <= 86399) or (i >= 86600 and i <= 86999) or i == 81850:
            group_code.append('078')
        elif(i >= 25000 and i <= 25989) or (i >= 26000 and i <= 26040) or i == 26060:
            group_code.append('079')
        elif i == 25990 or i == 26050 or (i >= 26070 and i <= 26090) or (i >= 26100 and i <= 26999):
            group_code.append('080')

        # -----------81-90--------------------- 
        elif (i >= 27000 and i <= 27699) or (i >= 49000 and i <= 49099):
            group_code.append('081')
        elif (i >= 28000 and i <= 28899) or (i >= 96000 and i <= 69099):
            group_code.append('082')
        elif (i >= 20000 and i <= 21009) or (i >= 21070 and i <= 21090) or (i>=21100 and i <= 21109) or(i >= 24100 and i >= 24109) or (i >= 24200 and i <= 24209):
            group_code.append('083')
        elif (i >= 21010 and i <= 21030) or i == 21060 or (i >= 21200 and i <= 21309) or i == 21450 or i == 21500 or (i >= 22000 and i <= 22020) or (i >= 22100 and i <= 22120) or (i >= 22200 and i <= 22309) or (i >= 23000 and i <= 23050) or (i >= 23100 and i <= 23400):
            group_code.append('084')
        elif i == 21040 or i == 21400 or (i >= 21600 and i <= 21820) or (i >= 24000 and i <= 24009) or (i >= 24050 and i <= 24060) or i == 24300:
            group_code.append('085')
        elif (i >= 15000 and i <= 15999) or i == 16010 or i == 16150:
            group_code.append('086')
        elif (i >= 16020 and i <= 16090) or i == 16100 or (i >= 16200 and i <= 16899) or (i >= 17000 and i <= 17999) or (i >= 18000 and i <= 18999):
            group_code.append('087')
        elif (i >= 87000 and i <= 87999):
            group_code.append('088')
        elif (i >= 88000 and i <= 88899):
            group_code.append('089')
        elif (i >= 89000 and i <=89999):
            group_code.append('090')

        # -----------91-100--------------------- 
        elif (i >= 90000 and i <= 90999):
            group_code.append('091')
        elif (i >= 91000 and i <= 91099) or (i >=  91200 and i <=91399):
            group_code.append('092')
        elif (i >= 91100 and i <= 91199):
            group_code.append('093')
        elif (i >= 93000 and i <= 93049) or (i >=93100 and i <=93300):
            group_code.append('094')
        elif (i >= 93050 and i <= 93099) or (i >=  93350 and i <= 93450):
            group_code.append('095')
        elif(i>= 94000 and i <= 95999):
            group_code.append('096')
        elif (i >= 96000 and i <=  96999):
            group_code.append('097')
        elif (i >= 97000 and i <= 97999):
            group_code.append('098')
        elif (i >= 98000 and i <= 98999 ):
            group_code.append('099')
        else:
            group_code.append('000')
    return group_code