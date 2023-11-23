
import pyodbc
import pandas as pd
from tkinter import messagebox, Button, Label, Tk
from tkcalendar import *

root = Tk()
root.title('EMEA')
root.geometry("600x400")
#root.iconbitmap(r'C:\Users\ChristophRazek\PycharmProjects')

# Query Connection
connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)

cal = Calendar(root, selectmode="day", date_pattern="y-mm-dd")
cal.pack(pady=20, fill="both", expand=True)



def grab_date():
    my_lable.config(text='Ausgewähltes Datum: ' + cal.get_date())

    datum = "'" + cal.get_date() + "'"

    # Welche Läger sollen abgefragt werden
    sql_lagernr = '''SELECT Distinct[LAGERNR]
      FROM [emea_enventa_live].[dbo].[LAGERORT] 
      where lagernr not in (101,102,103,110,195,198,199)
      order by LAGERNR'''

    df_lagernr = pd.read_sql(sql_lagernr, conx)
    lagernr = []

    for index, row in df_lagernr.iterrows():
        lagernr.append((row['LAGERNR']))


    for l in lagernr:
        sql_history = f'''with cte_artikel as (select artikelnr,bezeichnung, code2 as auslister, vk1/VKPRO as 'VK/Stk', kek/EKPRO as 'KEK/Stk' from artikel),
     cte_lager as (SELECT Distinct[LAGERNR], BEZEICHNUNG FROM [emea_enventa_live].[dbo].[LAGERORT] 
      where lagernr not in (101,102,103,110,195,198,199))

        select l.ARTIKELNR, Max(cte_artikel.BEZEICHNUNG) as 'BEZEICHNUNG', SUM(menge) as 'BESTAND', Max(l.LAGERNR) as 'LAGERNR', Max(cte_lager.BEZEICHNUNG) as 'LAGER', 
		Max(auslister) as 'AUSLISTER', Max([VK/Stk]) as 'VK/Stk', Max([KEK/Stk]) as 'KEK/Stk'  from LAGERJOURNAL as l
		left join cte_artikel on l.ARTIKELNR = cte_artikel.ARTIKELNR
		left join cte_lager on l.LAGERNR = cte_lager.LAGERNR
        where branchkey = 0110 
        and DATUM < {datum}
        and l.LAGERNR = {l}
        group by l.ARTIKELNR, BranchKey
        having sum(menge) > 0'''

        # DataFrame erstellung aus SQL
        df = pd.read_sql(sql_history, conx)
        # Export

        #df.to_excel(rf'S:\EMEA\Hist_Best\{datum}_{l}.xlsx', index=False)
        df.to_excel(rf'S:\EMEA\Hist_Best\{datum}_{l}.xlsx', index=False)

    messagebox.showinfo('Update erfolgreich!', 'Du kannst das Programm jetzt schließen')



my_button = Button(root, text='Datum Eingeben', command=grab_date)
my_button.pack(pady=5)

'''my_update = Button(root, text='Update Ausführen', command=update)
my_update.pack(pady=20)'''

my_lable = Label(root, text="")
my_lable.pack(pady=20)


root.mainloop()








