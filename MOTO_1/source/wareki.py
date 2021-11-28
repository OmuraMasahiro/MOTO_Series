import datetime

def wareki_conv(dateYMD):
    # 令和の開始日
    reiwa_year = datetime.date(2019,5,1)
    #平成の開始日
    heisei_year = datetime.date(1989,1,8)
    #昭和の開始日
    showa_year = datetime.date(1926,12,25)
    #大正開始日
    taisho_year = datetime.date(1912,7,30)
    wareki1='令和'
    wareki2=1
    if dateYMD >= reiwa_year:
        wareki2 = dateYMD.year % 100 - 18 #100で割ることで下二桁出す。
    elif dateYMD >= heisei_year:
        wareki1='平成'
        wareki2 = dateYMD.year % 100 + 12 #100で割ることで下二桁出す。
        wareki2=str(wareki2)[-2:]
    elif dateYMD >= showa_year:
        wareki1='昭和'
        wareki2 = dateYMD.year % 100 - 25 #100で割ることで下二桁出す。
    elif dateYMD >= taisho_year:
        wareki1='大正'
        wareki2 = dateYMD.year % 100 - 11 #100で割ることで下二桁出す。
    else:
        print("令和、平成、昭和ではない和暦となります。")
    return wareki1,str(wareki2)
def conv_num(wareki):
    if wareki=='大正':
        return '2'
    elif wareki=='昭和':
        return '3'
    elif wareki=='平成':
        return '4'
    else:
        return '5'

#w1,w2=wareki_conv(now)
#print(w1,str(w2),'年',now.month,'月',now.day,'日')
