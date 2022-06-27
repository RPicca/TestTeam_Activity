import openpyxl
import matplotlib.pyplot as plt

#TODO1 : gérer les feuilles cachées (prio faible)

#===============================================================================
# Config
#===============================================================================
# Nom du fichier : celui-ci ne doit contenir que des feuilles de répartition d'activités
#(penser à vérifier les feuilles cachées)
xlsx_file ='Suivi_01.xlsx'

#===============================================================================
wb_obj = openpyxl.load_workbook(xlsx_file, data_only=True)


#Retourne une liste avec la valeur des cellules d'une plage d'une feuille
def read_range_cells(sheet, r):
    cells = sheet[r[0]:r[1]]
    l=[]
    for c in cells:
        l.append(c[0].value)
    return l

def find_range(sheet, word):
    #On cherche la première cellule des activités ou du temps
    for i in range(1,101):
        for j in range(1,101):
            if sheet.cell(i,j).value==word:
                #Cette activité est en fait 2 ligne en dessous du mot clé recherché, on ajoute donc 2...
                first_cell = sheet.cell(i+2,j).coordinate
                I=i+2
                J=j
                break
    for i in range(I,101):
        if(sheet.cell(i,J).value)==None:
            last_cell=sheet.cell(i-1,J).coordinate
            break
    return [first_cell,last_cell]
def update_dico(sheet, dico):
    c=-1
    range_topic=find_range(sheet, 'HEATMAP')
    range_time=find_range(sheet, 'Total')
    keys= read_range_cells(sheet, range_topic)
    values=read_range_cells(sheet, range_time)
    # On choppe la taille des valeurs du dico -> Nbr de semaines déjà remplies
    M=0

    if bool(dico):
        M=len(list(dico.values())[0])
    for k in keys:
        c+=1
        if k in dico.keys():
            dico[k].append(values[c])
        else:
            #On met tout plein de 0 puisque c'est une nouvelle entrée et qu'il faut compenser les ratés
            l=M*[0]
            l.append(values[c])
            dico[k]=l
    # Ne pas oublier d'ajouter des 0 pour les activités n'ayant pas eu lieu cette semaine
    for k in set(dico.keys())-set(keys):
        dico[k].append(0)
  

#====================================================================================================
#                                                Graph
#====================================================================================================
def stackplot(dico, weeks): 
    fig, ax = plt.subplots()
    plt.stackplot(X, list(dico.values()), labels = dico.keys())
    plt.xticks(X, weeks)
    plt.ylabel("Nombre de jours", fontsize=16)
    plt.xlabel("Semaine", fontsize=16)
    plt.title("Activités de l'équipe Test", fontsize=24)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05), fancybox=True, shadow=True, ncol=5)
    plt.show()

#====================================================================================================
#                                                Pie
#====================================================================================================
def pie(dico):
    # L'idée est de sommer chaque temps puis de le diviser par le temps total
    dico_total={}
    for k in dico.keys():
        dico_total[k]=sum(dico[k])
    temps_total = sum(dico_total.values())
    # Deux boucles quasi identiques à la suite c'est moche, je sais...
    for k in dico_total.keys():
        dico_total[k]/=temps_total
    explode=[0]*len(dico.keys())
    explode[1]=0.05
    plt.pie(dico_total.values(),explode=explode,labels=dico_total.keys(),autopct='%1.1f%%', startangle = 30)
    plt.show()

#====================================================================================================
#                                                Write
#====================================================================================================
def write(dico, weeks):
    #On écrit le dico dans une feuille excel
    wb = openpyxl.Workbook()
    ws = ws1 = wb.active
    for i in range(1, len(weeks)+1):
        ws.cell(i+1,1,weeks[i-1])
    col=1
    for k in dico.keys():
        col+=1
        ws.cell(1,col,k)
        row=2
        for l in dico[k]:
            ws.cell(row,col,l)
            row+=1
    wb.save(filename = "Output.xlsx")

#====================================================================================================
#                                                Main
#====================================================================================================
dico={}
weeks=[]
for sheetname in wb_obj.sheetnames:
    weeks.append(sheetname.split("_")[0])
    update_dico(wb_obj[sheetname], dico)

X=range(len(weeks))
legend=list(dico.keys())
#On se remet dans le bon sens
weeks.reverse()
for i in dico.keys():
    dico[i].reverse()
    
stackplot(dico, weeks)
#pie(dico)
write(dico, weeks)
