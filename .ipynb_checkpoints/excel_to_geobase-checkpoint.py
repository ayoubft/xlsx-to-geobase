# Script python permettant de convertir le fichier excel contenant 
# l'ensemble des observations sous forme de carnet Trimble GSM
# exploitable sous covadis
#
# usage:
# dans votre invite de commande tapez ce qui suit (changer le nom de votre fichier!):
#   python excel_to_geobase.py
#
# Auteur: Ayoub FATIHI (ayoubft)


import pandas as pd    # pip install pandas openpyxl (1ere fois)

# Le fichier excel
xl = 'OBSERVATION_19-09-2021.xlsx'

# Lire le fichier excel et le stocker dans la 'DataFrame' table df
df = pd.read_excel(xl, sheet_name='test-1', header=1)

# Nettoyer la table en conservant que les colonnes souhaitées
cible = df.loc[:, ['Station','Cible','lecture','distance']]
cible.drop(0, axis=0, inplace=True)
cible.fillna(method='ffill', inplace=True)

# Création d'une chaîne de charactères contenant le format voulu
s = ''
c = cible['Station'][1]
for i, r in cible.iterrows():
    if r['Station'] != c:
        s += '2=\n'
        c = r['Station']
    s += '2=' + r['Station'] + '\n' + '5=' + r['Cible'] + '\n' + '7=' + str(r['lecture']) + '\n' + '8=100' + '\n' + '9=' + str(r['distance']) + '\n' 

# Enregister le résultat intermediaire
prm = xl.strip('xlsx') + 'prm'
f = open(prm, "w")
n = f.write(s)
f.close()

# Nettoyer qlq duplicats de station
lines_seen = set() # holds lines already seen
with open(prm, "r+") as f:
    d = f.readlines()
    f.seek(0)
    for i in d:
        if i.startswith('2='):
            if i == '2=\n': f.write(i)
            elif (i not in lines_seen):
                f.write(i)
                lines_seen.add(i)
        else:
            f.write(i)
                
    f.truncate()
