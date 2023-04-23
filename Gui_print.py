
from tkinter import *
import time,locale,customtkinter,time
import tkinter as tk
from tkinter import ttk,messagebox as mb
from PIL import Image, ImageTk
import pandas as pd

#top level
customtkinter.set_appearance_mode("light")
app=customtkinter.CTk()
app.title('SUIVI POLYME DES AFFAIRES'.center(320))
app.geometry('1300x670+30+10')
app.resizable(False,False)
app.iconbitmap('C:\\Users\\Khaaoula\\Documents\\Stelia\\PFA_Stelia\\img\\stelia-logo.ico')

#frame de la date et le logo
choices_fr=customtkinter.CTkFrame(master=app,corner_radius=0,height=80,width=1250,
        fg_color='#1363DF')
choices_fr.pack(fill = BOTH)

#affichage du logo 
logo = Image.open("C:\\Users\\Khaaoula\\Documents\\Stelia\\PFA_Stelia\\img\\stelia_logo.png")
resized= logo.resize((180,62), Image.Resampling.LANCZOS)
new_logo= ImageTk.PhotoImage(resized)
label_i =customtkinter.CTkLabel(choices_fr,image=new_logo)
label_i.pack(side=LEFT,padx=80)

#Affichage de la date 
locale.setlocale(locale.LC_TIME,'')
date=time.strftime('%A %d/%m/%Y  %H:%M')
date_label =customtkinter.CTkLabel(choices_fr, text=date.capitalize(),
           text_color='white')
date_label.configure(font=("Helvetica", 18,'italic','bold'))
date_label.pack(pady=15,side=RIGHT,padx=150)


#create d'un separateur 
separator = ttk.Separator(app, orient='horizontal')
separator.pack(fill="x", pady=20)


#frame combobox d'affaire
combo_fr=customtkinter.CTkFrame(master=app,height=100,width=1350,
        fg_color='#ebebec')
combo_fr.pack(fill = BOTH,pady=5)


#ceation de combobox des affaire
aff=customtkinter.StringVar(value="Affaire")
affaire_combo=customtkinter.CTkComboBox(master=combo_fr, width=220,variable=aff,values=('AMN','CCK1et2','CCK3','CCK4','P PLIE CCK4-2','HUD'),
text_font=("Helvetica", 12,"italic"),dropdown_text_font=("Helvetica", 10),state='readonly',dropdown_hover_color="#7FB5FF",border_color="#185ADB",
button_color="#185ADB")
affaire_combo.pack(side=LEFT,padx=180, pady=10)


#creation de zone input des semaines

sem_en=customtkinter.CTkEntry(combo_fr, width=200,placeholder_text="Semaine",placeholder_text_color="black" ,text_font=("Helvetica", 12,"italic"),
border_color='#185ADB',)
sem_en.pack(side=LEFT,padx=100, pady=10)


#Frame of output
output_fr=customtkinter.CTkFrame(master=app,height=30,width=1350,corner_radius=6,fg_color='#DFDFDE')
output_fr.pack(pady=40)

scroll_y=Scrollbar(output_fr,orient=VERTICAL)
scroll_y.pack(side=RIGHT,fill=Y,padx=5)

s=ttk.Style()
s.theme_use('clam')
s.configure('Treeview', rowheight=30,background='#A7D1D2',border_width=2)


#creation de tableau
output_tab=ttk.Treeview(output_fr,height=80,selectmode='browse',columns=['af','ref','obj','real'],yscrollcommand=scroll_y.set)
output_tab.pack()
#creation de scroll
scroll_y.configure(command=output_tab.yview)


#show headingsK
output_tab['show'] = 'headings'
output_tab.heading('af',text='Affaire')
output_tab.heading('ref',text='Réference')
output_tab.heading('obj',text='Objéctif')
output_tab.heading('real',text='Réalisable/Objéctif')
#output_tab.heading('mq',text='Manquant')
#output_tab.heading('os',text='Over Stock')

output_tab.column('af',width=200,anchor=CENTER)
output_tab.column('ref',width=350,anchor=CENTER)
output_tab.column('obj',width=200,anchor=CENTER)
output_tab.column('real',width=250,anchor=CENTER)
#output_tab.column('mq',width=200)
#output_tab.column('os',width=200)


#submit func
def submit():
    
    #1) lire le contenu du combobox       
    if len(sem_en.get())>0 and str(sem_en.get()[-2:])<= str(time.strftime('%V')):
        #fonction de submit
        global choice_af,sem
        choice_af=affaire_combo.get()
        sem=sem_en.get()
        for i in output_tab.get_children():
            output_tab.delete(i)
        workbook = pd.read_excel('C:\\Users\\Khaaoula\\Documents\\Stelia\\PFA_Stelia\\input\\SUIVI_POLYME.xlsx',sheet_name=choice_af,
        usecols=[1, 2, 3,14])
        workbook=workbook.dropna()
        list_w=workbook.values.tolist()
        
        
        # appel du treeview et affichage
        for sous_l in list_w:
            output_tab.insert("","end",values=sous_l)
        
             
        
                
    else: 
        mb.showwarning('Erreur','Veuillez entrer une semaine valide ou ecrire la semaine comme "sem 13"') #

#fct recuperer les valeur par le boutton

#creation de boutton 
btn_search=customtkinter.CTkButton(combo_fr,width=150,fg_color='#185ADB', text='Chercher',text_color='white',text_font=("Helvetica", 12,"italic"),
command=submit)
btn_search.pack(side=LEFT,padx=20, pady=10)
#app.bind('event', submit)



app.mainloop()