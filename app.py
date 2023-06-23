import requests
import xlwt
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

print("Program Type Seçin:")
print("1 - Güncel Bülten(Aktif)")
print("2 - Sonraki Bülten(Hazırlanıyor!)")
print("3 - Gelecek Maçlar(Hazırlanıyor!)")
print("-"*100)

program_type = input("Seçiminizi yapın (1-3): ")

url = 'https://sportprogram.iddaa.com/SportProgram'

params = {
    'ProgramType': program_type,
    'SportId': '1',
    'MukList': '1_1,2_88,2_100,2_101_2.5,2_89'
}

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7',
    'Cache-Control': 'max-age=0',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1'
}

response = requests.get(url, params=params, headers=headers, verify=False)

if response.status_code == 200:
    sonuc = response.json()

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Sheet1')

    # Set headers
    headers = ['Lig', 'Tarih1', 'Saat', 'Ev Sahibi', 'Deplasman', 'MS1', 'MSX', 'MS2', 'IY1', 'IYX', 'IY2', '2,5 A/U', 'KGV', 'KGY']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    for row, sup in enumerate(sonuc['data']['spg'][0]['eventGroup'][0]['eventResponse'], start=1):
        # Extract data
        lig = sup['cn']
        tarih1 = sup['ede']
        tarih2 = sup['edh'] #sup['e']
        ev = sup['en'].split(' - ')[0]
        dep = sup['en'].split(' - ')[1]
        ms1 = sup['m'][0]['o'][0]['odd']
        msx = sup['m'][0]['o'][1]['odd']
        ms2 = sup['m'][0]['o'][2]['odd']
        IY1 = sup['m'][1]['o'][0]['odd']
        IYx = sup['m'][1]['o'][1]['odd']
        IY2 = sup['m'][1]['o'][2]['odd']

        # Handikaplı Maç Sonucu
        handikap_ms = '-'
        for option in sup['m'][2]['o']:
            if option['ona'] == 'H':
                handikap_ms = option['odd']

        # 2.5 Gol A/Ü
        alt = '-'
        ust = '-'
        for option in sup['m'][3]['o']:
            if option['ona'] == 'Alt':
                if option['ona']:
                    alt = option['odd']
            elif option['ona'] == 'Üst':
                if option['ona']:
                    ust = option['odd']

        
        alt = alt or ''
        ust = ust or ''
        
        gol_au = ''
        if alt and ust:
            gol_au = f"{alt} / {ust}"

        KGV = sup['m'][4]['o'][0]['odd']
        KGY = sup['m'][4]['o'][1]['odd']

        data = [lig, tarih1, tarih2, ev, dep, ms1, msx, ms2, IY1, IYx, IY2, gol_au, KGV, KGY]

        for col, value in enumerate(data, start=0):
            worksheet.write(row, col, value)

        workbook.save('iddaa_bulten_data.xls')
    print("Tebrikler : İşlemleriniz Tamamlandı!")
else:
    print('Error:', response.status_code)
