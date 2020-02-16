import xlrd
f = ['Центральный административный округ', 'Северный административный округ', 'Северо-Восточный административный округ','Восточный административный округ','Юго-Восточный административный округ','Южный административный округ','Юго-Западный административный округ','Западный административный округ','Северо-Западный административный округ','Зеленоградский административный округ(','Новомосковский административный округ','Троицкий административный округ']
n=0
k=0
for i in range(12):
    n=n+1
    print(f[i],"№",n)
print('Введите ближайший индекс округа')
k=int(input())
for i in range(13):
    if k==i+1:
        okr=f[i]
active = input('Какой отдых более предпочтителен: активный/пассивный?')
distance = input('Важно ли расстояние: да/нет?')
price = input('Какой отдых более предпочтителен: платно/бесплатно?')
otdix=input('Какой вид отдыха вы предпочитаете?')
def func():
    a = []
    for i in range(1, len(vals)):
        a.append(vals[0 + i][1])
        a.append(vals[0 + i][4])
        a.append(vals[0 + i][29])
    if distance == 'да':
        for i in range(len(a)):
            if okr == a[i]:
                if price == a[i + 1]:
                    print(f[i], a[i - 1])
    else:
        for i in range(len(a)):
            print(a[i - 1])
if active == 'активный':
    x = input('Какая область вас интересует?')
    if x == 'катание на лошадях':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/katanie_na_lochadyax.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()
    if x == 'стрельбища':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/sterlb.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()
    if x == 'тир':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/strelkov.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()
    if x == 'музыкальные площадки':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/DancePlosh.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()
    if otdix =='водный':
        if x == 'аквапарк':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/acva.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        elif x == 'бассейн':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/bassein_critie.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
    if otdix=='игры с мячем':
        if x=='футбол':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/Football.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        elif x == 'зона отдыха':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/ZonaOtdiha.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
    if otdix =='зимнее игры':
        if x == 'каток':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/katok.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        if x == 'сноуборд в парках':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/snoubordvparkah.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
    if otdix =='спортивные залы или городки':
        if x == 'тренажёрный городок':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/trengor.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        elif x=='спортивные площадки':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/Sport_plosh(1).xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
    if otdix =='отдых для детей':
        if x=='детские площадки':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/detskie_plochadkie.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
    if otdix == 'зимний спорт':
        if x=='снежные горки':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/GorkiSnesh.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        if x=='Горнолыжный':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/gornolignie.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        if x=='ледянные горки':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/ledyanie_gorki.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
        if x=='сноуборды':
            rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/snoubordvparkah.xls', formatting_info=True)
            sheet = rb.sheet_by_index(0)
            val = sheet.row_values(1)[1]
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            func()
elif active=='пассивный':
    if x == 'зона отдыха':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/ZonaOtdiha.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()
    if x == 'детские площадки':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/detskie_plochadkie.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()
    if x == 'разивающие детские центры':
        rb = xlrd.open_workbook('C:/Users/User/Desktop/Olimpuda/detskie_plochadkie.xls', formatting_info=True)
        sheet = rb.sheet_by_index(0)
        val = sheet.row_values(1)[1]
        vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        func()

