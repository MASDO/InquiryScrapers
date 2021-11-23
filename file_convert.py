import pandas as pd
path = "C:\\Users\\M_manteghipoor.SB24\\Desktop\\saman_samat"
# '//RD7-115//Users//p_mohebmaleki//Desktop//share'
# os.path.join(os.path.join(os.environ['USERPROFILE'], 'Desktop'))
# print(path)


Inquiry_path_load = path + '\\saman_samat.csv'
if __name__ == '__main__':
    samat = pd.read_csv(Inquiry_path_load)
    i = 0
    k = 0
    for j in range(0, 15):
        s = samat.iloc[i:i+10000, :].copy(deep=True)
        s = pd.DataFrame(s)
        name = "samat_{}.xlsx".format(str(k))
        s.to_excel(name)
        i = i + 10001
        k = k + 1


