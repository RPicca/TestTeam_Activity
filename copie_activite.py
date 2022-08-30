import openpyxl
import matplotlib.pyplot as plt
import PySimpleGUI as sg 
import os
import tempfile
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

#Retourne une liste avec la valeur des cellules d'une plage d'une feuille
def read_range_cells(sheet, r):
    cells = sheet[r[0]:r[1]]
    l=[]
    for c in cells:
        l.append(c[0].value)
    return l

def find_range(sheet, word):
    I=J=-1
    #On cherche la première cellule des activités ou du temps
    for i in range(1,101):
        for j in range(1,101):
            if sheet.cell(i,j).value==word:
                #Cette activité est en fait 2 ligne en dessous du mot clé recherché, on ajoute donc 2...
                first_cell = sheet.cell(i+2,j).coordinate
                I=i+2
                J=j
                break
    if I==-1:
        return None
    for i in range(I,101):
        if(sheet.cell(i,J).value)==None:
            last_cell=sheet.cell(i-1,J).coordinate
            break
    return [first_cell,last_cell]

def update_dico(sheet, dico):
    c=-1
    range_topic=find_range(sheet, 'HEATMAP')
    range_time=find_range(sheet, 'Total')
    if range_time == None or range_topic == None:
        print("Sheet "+ sheet.title+" does not contain the HEATMAP and Total Keywords")
        return -1
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
    return 0
  

#====================================================================================================
#                                                Graph
#====================================================================================================
def stackplot(dico, Weeks, fntsize):
    if fntsize=="":
        fntsize=18
    else:
        fntsize=int(fntsize)
    weeks=[]
    for i in Weeks:
        # On ne garde que le numéro de semaine, ça prend trop de place sinon...
        # Normalement il y a un underscore -> on prend tout ce qu'il y a avant.
        # Sinon on ne garde que les 3 premiers caractères
        tmp=i.split("_")[0]
        if len(tmp)>3:
            tmp=tmp[:3]
        weeks.append(tmp)
    fig, ax = plt.subplots()
    X = range(len(weeks))
    if (len(weeks))>30:
        for i in range (0,len(weeks)-1,2):
            weeks[i]=""
    plt.stackplot(X, list(dico.values()), labels = dico.keys())
    plt.xticks(X, weeks, fontsize=fntsize)
    plt.yticks(fontsize=fntsize)
    plt.ylabel("Nombre de jours", fontsize=fntsize*1.25)
    plt.xlabel("Semaines", fontsize=fntsize*1.25)
    plt.title("Activités de l'équipe Test", fontsize=fntsize*1.25)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.1), fancybox=True, shadow=True, ncol=5, fontsize=fntsize)
    manager = plt.get_current_fig_manager()
    manager.window.state('zoomed')
    plt.subplots_adjust(bottom=0.2)
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
    ws = wb.active
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
#                                       Recup d'info dans la GUI
#====================================================================================================
def interface_input(filepath=""):
    layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(filepath, key="file"), sg.FileBrowse(), sg.Button('Download')],
            [sg.Checkbox('Show Stackplot', default=True, key="stackplot")],
            [sg.Checkbox('Write output in xlsx file', default=True, key="write")],
            [sg.Checkbox('Show Pie (awful design)', default=False, key="pie")],
            [sg.Button('Run')]]
    # Create the window
    window = sg.Window('TestTeam Activity', layout)
    dico = {}
    while True:
        # Display and interact with the Window
        event, values = window.read()
        if event ==sg.WIN_CLOSED:
            break
        if event == "Run":
            dico["file"] = values["file"]
            dico["stackplot"] = values["stackplot"]
            dico["write"] = values["write"]
            dico["pie"] = values["pie"]
            break
        if event == "Download":
            sha = interface_sharepoint()
            filepath = download(sha["sharepoint"], sha["filepath"], sha["email"], sha["password"])
            window.Element("file").update(filepath)
    window.close()
    return dico
#====================================================================================================
#                                     GUI de selection de la plage
#====================================================================================================
def interface_data_range(path, weeks):
    layout = [[sg.T(""), sg.Text(path)],
              [sg.Text("Choose first week"), sg.Listbox(weeks, size=(30,3), key="first")],
              [sg.Text("Choose last week"), sg.Listbox(weeks, size=(30,3), key="last")],
              [sg.Text("Font size :"), sg.Input("18",key="font", size=(3,3))],
              [sg.Button('Run', key="Run")]]
    # Create the window
    window = sg.Window('Data Range', layout)
    # Display and interact with the Window
    event, values = window.read()
    if event=="Run":
        return[window.Element('first').Widget.curselection()[0], window.Element('last').Widget.curselection()[0], values['font']]
#====================================================================================================
#                                       Détection des feuilles Excel
#====================================================================================================
def filter_sheets(workbook):
    L=[]
    for sheetname in wb_obj.sheetnames:
        # On considère qu'une feuille est exploitable si elle contient un S au début
        if (sheetname[0] == "S" or sheetname[0]=="s") and any(char.isdigit() for char in sheetname):
            L.append(sheetname)
    return L
#====================================================================================================
#                                   Téléchargement depuis Sharepoint
#====================================================================================================
def download(site_url, file_url, email, password):
    if site_url=="":
        site_url="https://forsksas.sharepoint.com/sites/Testteam2/"
    if file_url=="":
        file_url="/sites/Testteam2/Documents partages/2021 - Réunions hebdo/Suivi des activités hebdomadaires.xlsx"
    download_path = os.path.join(tempfile.mkdtemp(), os.path.basename(file_url))
    ctx = ClientContext(site_url).with_credentials(UserCredential(email, password))
    with open(download_path, "wb") as local_file:
        file = ctx.web.get_file_by_server_relative_path(file_url).download(local_file).execute_query()
    print("[Ok] file has been downloaded into: {0}".format(download_path))
    return download_path

#====================================================================================================
#                                       GUI de download sharepoint
#====================================================================================================
def interface_sharepoint():
    layout = [[sg.Text("Email: "), sg.Input(key="email")],
              [sg.Text("Password: "), sg.Input(key="password", password_char="*")],
              [sg.Text("Path to Sharepoint"), sg.Input("https://forsksas.sharepoint.com/sites/Testteam2/",key="sharepoint")],
              [sg.Text("Path to file"),sg.Input("/sites/Testteam2/Documents partages/2021 - Réunions hebdo/Suivi des activités hebdomadaires.xlsx",key="filepath")],
            [sg.Button('Download')] ]
    # Create the window
    window = sg.Window('Sharepoint download', layout)

    # Display and interact with the Window
    event, values = window.read()
    window.close()
    dico={}
    dico["email"]=values["email"]
    dico["password"]=values["password"]
    dico["filepath"]=values["filepath"]
    dico["sharepoint"]=values["sharepoint"]
    return dico
#====================================================================================================
#                                                Main
#====================================================================================================
# dico=interface_sharepoint()
# download(dico["sharepoint"], dico["filepath"], dico["email"], dico["password"])
UI = interface_input()
xlsx_file = UI["file"]
stack = UI["stackplot"]
write_output=UI["write"]
pie_chart=UI["pie"]
#===============================================================================
wb_obj = openpyxl.load_workbook(xlsx_file, data_only=True)
dico={}
weeks=[]
sheets = filter_sheets(wb_obj)
[first, last, fntsize] =interface_data_range(xlsx_file, sheets)
for sheetname in sheets[last:first+1]:
    if update_dico(wb_obj[sheetname], dico)==0:
        weeks.append(sheetname)
legend=list(dico.keys())
#On se remet dans le bon sens
weeks.reverse()
for i in dico.keys():
    dico[i].reverse()

if stack:
    stackplot(dico, weeks, fntsize)
if pie_chart:
    pie(dico)
if write_output:
    write(dico, weeks)
