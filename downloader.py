import json
import tkinter
import tkinter.messagebox
from customtkinter import *
from operator import itemgetter

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from datetime import *

def download_certificates(activecnpjs: list):
    result = ""
    downloaded = 0
    options = Options()
    # options.add_argument('--headless')
    options.add_argument('--start-maximized')
    options.add_experimental_option("detach", True)
    appState = {
        "recentDestinations": [
            {
                "id": "Save as PDF",
                "origin": "local"
            }
        ],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    profile = {'printing.print_preview_sticky_settings.appState': json.dumps(appState)}
    options.add_experimental_option('prefs', profile)
    options.add_argument('--kiosk-printing')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                              options=options)

    driver.get("https://cnd.pbh.gov.br/CNDOnline/")

    wait = WebDriverWait(driver, 6)

    original_window = driver.current_window_handle
    try:
        radiobutton = driver.find_element(By.XPATH,
                                          "//div[@id='meuForm:TIPO1']")

        radiobutton.click()

        textbox = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH,
                                                                                    "/html/body/div[2]/form/table/tbody/tr/td/fieldset/table[3]/tbody/tr[2]/td[2]/span/div/span/input")))

        for current in range(len(activecnpjs)):
            currentCNPJ = activecnpjs[current]['CNPJ']
            currentName = activecnpjs[current]['Nome']
            textbox.send_keys(f"\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b{currentCNPJ}")

            searchbutton = driver.find_element(By.XPATH,
                                               "//a[@id='meuForm:pesquisar']")
            searchbutton.click()

            wait.until(EC.number_of_windows_to_be(2))

            for window_handle in driver.window_handles:
                if window_handle != original_window:
                    driver.switch_to.window(window_handle)
                    break

            driver.execute_script('document.title="{}";'.format(currentName + datetime.today().strftime(' %m%Y')))
            driver.execute_script('window.print();')
            downloaded += 1
            driver.close()
            driver.switch_to.window(original_window)
            sleep(1)
        result = "Download concluído."
    except:
        result = "Ocorreu um erro no acesso ao site. Tente novamente mais tarde."
    finally:
        resultmessage = result + f"\n{downloaded} arquivos baixados."
        driver.quit()
        tkinter.messagebox.showinfo(title="Relatório", message=resultmessage)


class DatabaseEmpresas:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = self.load_data()

    def load_data(self):
        try:
            with open(self.file_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return {"empresas": []}

    def save_data(self):
        sorteddata = sorted(self.data['empresas'], key=itemgetter("Nome"))
        self.data['empresas'] = sorteddata
        with open(self.file_path, 'w') as f:
            json.dump(self.data, f, indent=2)

    def add_company(self, cnpj, nome, tem_ie: bool, cep, numero, tipocomplemento, complemento):
        dictempresa = {"CNPJ": str(cnpj),
                       "Nome": str(nome),
                       "temIE": tem_ie,
                       "CEP": str(cep),
                       "Numero": str(numero),
                       "tipoComplemento": str(tipocomplemento),
                       "Complemento": str(complemento)}
        self.data["empresas"].append(dictempresa)
        self.save_data()

    def delete_company(self, cnpjs: list):
        for cnpj in cnpjs:
            for empresa in self.data["empresas"][:]:
                if empresa["CNPJ"] == str(cnpj):
                    self.data["empresas"].remove(empresa)
        self.save_data()

    def edit_company(self, cnpj, nome, tem_ie: bool, cep, numero, tipocomplemento, complemento):
        dictempresa = {"CNPJ": str(cnpj),
                       "Nome": str(nome),
                       "temIE": tem_ie,
                       "CEP": str(cep),
                       "Numero": str(numero),
                       "tipoComplemento": str(tipocomplemento),
                       "Complemento": str(complemento)}
        for empresa in self.data["empresas"][:]:
            if empresa["CNPJ"] == str(cnpj):
                empindex = self.data["empresas"].index(empresa)
                self.data["empresas"][empindex] = dictempresa
        self.save_data()

    def get_company_data(self, cnpj):
        for empresa in self.data["empresas"][:]:
            if empresa["CNPJ"] == str(cnpj):
                return empresa


empresas = DatabaseEmpresas("clist.json")

set_appearance_mode("dark")
set_default_color_theme("dark-blue")

root = CTk()
root.title("Downloader de CND")
root.geometry('600x600')
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)

# tabview
tabview = CTkTabview(root)
tabview.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
tabview.add("CNPJ")
tabview.add("CPF")

# tabview row configure
tabview.tab("CNPJ").columnconfigure(0, weight=1)
tabview.tab("CNPJ").columnconfigure(1, weight=0)
tabview.tab("CNPJ").columnconfigure(2, weight=0)
tabview.tab("CNPJ").rowconfigure(0, weight=1)
tabview.tab("CNPJ").rowconfigure(1, weight=0)

# CNPJ
# checkbox list
CNPJ_checkboxframe = CTkFrame(tabview.tab("CNPJ"), width=200)
CNPJ_checkboxframe.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
CNPJ_checkboxframe.columnconfigure(1, weight=1)
CNPJ_checkboxframe.columnconfigure(1, weight=1)
CNPJ_checkboxframe.rowconfigure(0, weight=0)
CNPJ_checkboxframe.rowconfigure(1, weight=1)

def select_all():
    for box in empbox:
        box.select()

def deselect_all():
    for box in empbox:
        box.deselect()

selectall = CTkButton(CNPJ_checkboxframe, text="Selecionar Todos", command=select_all)
selectall.grid(row=0, column=0, pady=5, sticky="nsw")
deselectall = CTkButton(CNPJ_checkboxframe, text="Desmarcar Todos", command=deselect_all)
deselectall.grid(row=0, column=1, pady=5, padx=5, sticky="nsw")
CNPJ_empresasframe = CTkScrollableFrame(CNPJ_checkboxframe)
CNPJ_empresasframe.grid(row=1, column=0, sticky="nsew", columnspan=2)
empbox = []
empvar = []


def create_checkboxes():
    for emp in range(len(empresas.data["empresas"])):
        empvar.append(StringVar())
        empbox.append(CTkCheckBox(CNPJ_empresasframe,
                                  text=f"{empresas.data['empresas'][emp]['CNPJ']}, {empresas.data['empresas'][emp]['Nome']}",
                                  variable=empvar[emp], onvalue=empresas.data['empresas'][emp]['CNPJ'], offvalue='0'))
        empbox[emp].deselect()
        empbox[emp].grid(row=emp, column=0, pady=1, sticky="nw")


def destroy_checkboxes():
    for emp in range(len(empbox)):
        empbox[emp].destroy()
    empbox.clear()
    empvar.clear()


def update_checkboxes():
    destroy_checkboxes()
    create_checkboxes()


def get_selected_companies():
    activecompanies = []
    for i in empvar:
        if i.get() != '0':
            activecompanies.append(empresas.get_company_data(i.get()))
    print(activecompanies)
    return activecompanies


def get_selected_companies_cnpj():
    activecompanies = []
    for i in empvar:
        if i.get() != '0':
            activecompanies.append(empresas.get_company_data(i.get())["CNPJ"])
    return activecompanies

def download_selected():
    activecompanies = get_selected_companies()
    if not activecompanies:
        tkinter.messagebox.showerror(title="Erro", message="Nenhuma empresa selecionada.")
        return
    download_certificates(activecompanies)
    return


create_checkboxes()

# right frame
CNPJ_manageframe = CTkFrame(tabview.tab("CNPJ"), width=500)
CNPJ_manageframe.grid(row=0, column=2, rowspan=2, padx=5, pady=5, sticky="nsew")
CNPJ_manageframe.rowconfigure((0,1,2,3,4,5,6,8,9), weight=0)
CNPJ_manageframe.rowconfigure(7, weight=1)
# CNPJ Labels and Entries
# entryCNPJlabel = CTkLabel(CNPJ_manageframe, text="CNPJ:")
# entryCNPJlabel.grid(row=0, column=0, padx=5, pady=1, sticky="nws")
# entryCNPJ = CTkEntry(CNPJ_manageframe, width=180)
# entryCNPJ.grid(row=1, column=0, padx=5, pady=1, sticky="nw")
# entryNOME1label = CTkLabel(CNPJ_manageframe, text="Nome:")
# entryNOME1label.grid(row=2, column=0, padx=5, pady=1, sticky="nws")
# entryNOME1 = CTkEntry(CNPJ_manageframe, width=180)
# entryNOME1.grid(row=3, column=0, padx=5, pady=1, sticky="nw")


def new_entry_window():
    entry_window = CTkToplevel(root)
    entry_window.title('Nova Empresa')
    entry_window.rowconfigure((0,1), weight=0)
    entry_window.rowconfigure((2,3), weight=0)
    entry_window.rowconfigure((4,5), weight=0)

    def checkint(inp):
        if inp.isdigit():
            return True
        elif inp == "":
            return True
        else:
            return False

    reg = entry_window.register(checkint)
    NE_cnpjlabel = CTkLabel(entry_window, text="CNPJ:*")
    NE_cnpjlabel.grid(row=0, column=0, padx=5, pady=5, sticky='nsw')
    NE_cnpjentry = CTkEntry(entry_window, validate="key", validatecommand=(reg, '%P'))
    NE_cnpjentry.grid(row=1, column=0, padx=5, pady=5, sticky='nsw')

    NE_nomelabel = CTkLabel(entry_window, text="Nome:*")
    NE_nomelabel.grid(row=0, column=1, padx=5, pady=1, sticky='nsw')
    NE_nomeentry = CTkEntry(entry_window)
    NE_nomeentry.grid(row=1, column=1, padx=5, pady=5, sticky='nsw')

    IEvar = IntVar()
    NE_temIE = CTkCheckBox(entry_window, text="Tem Inscrição Estadual?", variable=IEvar)
    NE_temIE.grid(row=1, column=2, padx=5, pady=5, sticky='nw')

    NE_CEPlabel = CTkLabel(entry_window, text="CEP:")
    NE_CEPlabel.grid(row=2, column=0, padx=5, pady=1, sticky='sw')
    NE_CEPentry = CTkEntry(entry_window, validate="key", validatecommand=(reg, '%P'))
    NE_CEPentry.grid(row=3, column=0, padx=5, pady=5, sticky='nw')

    NE_numerolabel = CTkLabel(entry_window, text="Número:")
    NE_numerolabel.grid(row=2, column=1, padx=5, pady=1, sticky='sw')
    NE_numeroentry = CTkEntry(entry_window, validate="key", validatecommand=(reg, '%P'))
    NE_numeroentry.grid(row=3, column=1, padx=5, pady=5, sticky='nw')

    NE_tipolabel = CTkLabel(entry_window, text="Tipo do Complemento:")
    NE_tipolabel.grid(row=4, column=0, padx=5, pady=1, sticky='sw')
    tipos = ['', 'Apto']
    tipovar = StringVar()
    NE_tipocombobox = CTkComboBox(entry_window, variable=tipovar, values=tipos)
    NE_tipocombobox.grid(row=5, column=0, padx=5, pady=5, sticky='nw')

    NE_complementolabel = CTkLabel(entry_window, text="Complemento:")
    NE_complementolabel.grid(row=4, column=1, padx=5, pady=1, sticky='sw')
    NE_complementoentry = CTkEntry(entry_window)
    NE_complementoentry.grid(row=5, column=1, padx=5, pady=5, sticky='nw')

    NE_feedbacklabel = CTkLabel(entry_window)
    NE_feedbacklabel.grid(row=4, column=2)
    errormessage = StringVar()
    errormessage.set('')
    NE_feedbacklabel.configure(textvariable=errormessage)

    def confirm_add_company():
        cnpj = NE_cnpjentry.get()
        if len(cnpj) != 14:
            errormessage.set('CNPJ inválido')
            return
        nome = NE_nomeentry.get()
        tem_ie = bool(IEvar.get())
        cep = NE_CEPentry.get()
        numero = NE_numeroentry.get()
        if not tem_ie:
            if cep == "" or numero == "":
                errormessage.set('Empresa sem IE deve informar endereço')
                return
            if len(cep) != 8:
                errormessage.set('CEP inválido')
                return
        tipo = NE_tipocombobox.get()
        complemento = NE_complementoentry.get()
        empresas.add_company(cnpj, nome, tem_ie, cep, numero, tipo, complemento)
        update_checkboxes()
        entry_window.destroy()
        return

    NE_addbutton = CTkButton(entry_window, text='Cadastrar Empresa', command=confirm_add_company)
    NE_addbutton.grid(row=5, column=2)

    # CNPJ Feedback label

    # Lock root until entry_window closes
    entry_window.transient(root)
    entry_window.grab_set()
    root.wait_window(entry_window)


# CNPJ Buttons
spacer1 = CTkLabel(CNPJ_manageframe, text="")
spacer1.grid(row=0, column=0, pady=1)
CNPJ_adicionarempresa = CTkButton(CNPJ_manageframe, text="Nova Empresa", command=new_entry_window)
CNPJ_adicionarempresa.grid(row=1, column=0, padx=5, pady=5)


def confirm_deletion():
    companiesfordeletion = []
    for i in get_selected_companies():
        companiesfordeletion.append(i['CNPJ'])
    if not companiesfordeletion:
        tkinter.messagebox.showerror(title="Erro",message="Nenhuma empresa selecionada.")
        return
    confirmation = tkinter.messagebox.askquestion('Are you sure?', 'Tem certeza que deseja excluir as empresas selecionadas?')
    if confirmation != 'yes':
        return
    empresas.delete_company(companiesfordeletion)
    tkinter.messagebox.showinfo(message=f"Empresa(s) {companiesfordeletion} removidas com sucesso.")
    update_checkboxes()

CNPJ_excluirempresa = CTkButton(CNPJ_manageframe, text="Excluir Empresa", command=confirm_deletion)
CNPJ_excluirempresa.grid(row=2, column=0, padx=5, pady=5)

# Download Button
CNPJ_baixarcnd = CTkButton(tabview.tab("CNPJ"), text="Baixar Certidões", command=download_selected)
CNPJ_baixarcnd.grid(row=1, column=0, padx=5, pady=10, sticky="sew", columnspan=2)

# Debug
def gettime():
    print(datetime.today().strftime(' %m%Y'))

debugbutton = CTkButton(tabview.tab("CNPJ"),text='Debug', command=gettime)
debugbutton.grid(row=2, column=0, padx=5, pady=10, sticky="sew", columnspan=2)

# CPF
# checkbox list

# right frame configure

# tabview

# CPF Labels and Entries

# CPF Buttons

# CPF Feedback label

# Manage DB buttons

# Download Button





# def myclick():
#     myLabel = CTkLabel(root, text=e.get())
#     myLabel.grid(row=3, column=0, pady=10)
#
#
# e = CTkEntry(root, width=120)
# e.grid(row=1, column=0)
# e.insert(0, "Enter CNPJ:")
#
# label1 = CTkLabel(master=root, text="Teste")
# label1.grid(row=0, column=0, pady=10)
#
# testButton = CTkButton(master=root, text="Confirmar CNPJ", width=100, height=100, command=myclick)
# testButton.grid(row=2, column=0, padx=50, pady=50)


root.mainloop()

