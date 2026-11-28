import openpyxl


#------------------------------------
# Lecture du fichier excel
#------------------------------------


workbook = openpyxl.load_workbook('inscription_c0d1ngup (2).xlsx', data_only = True)
titres_onglets = workbook.sheetnames
onglet1 = workbook[titres_onglets[0]]
#Liste des lignes
lignes = []
for row in onglet1.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))

#affichage des listes/lignes
for r in lignes:
    print(r)

#ferme le fichier excel
workbook.close()