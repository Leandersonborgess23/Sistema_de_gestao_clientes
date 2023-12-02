import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import pathlib
from openpyxl import Workbook
import openpyxl

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.appearance()
        self.todo_sistema()

    def layout_config(self):
        self.title("Sistema de gestao de clientes")
        self.geometry("700x500")

    def appearance(self):
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color="#ffffff", text_color=['#000', '#000']).place(x=50, y=430)
        self.opt_apm = ctk.CTkOptionMenu(self, values=["Light", "Dark", "System"], command=self.change_apm).place(x=50, y=460)

    def todo_sistema(self):
        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, bg_color="teal", fg_color="teal")
        frame.place(x=0, y=10)
        title = ctk.CTkLabel(frame, text="Sistema de gestão de clientes", font=("Century Gothic bold", 24), text_color="#fff").place(x=190, y=10)

        span = ctk.CTkLabel(self, text="Por favor, preencha todos os campos do formulário!", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=50, y=70)

        ficheiro_path = pathlib.Path("Clientes.xlsx")
    

        if ficheiro_path.exists():
            pass

        else:
            workbook = Workbook()
            folha=ficheiro.active
            folha['A1']="Nome completo"
            folha['B1']="Contato"
            folha['C1']="Idade"
            folha['D1']="Genero"
            folha['E1']="Endereco"
            folha['F1']="Observacoes"

            ficheiro.save("Clientes.xlsx")


        def submit():

            #pegando os dados dos entrys
            nome = nome_valor.get()
            contato = contato_valor.get()
            idade = idade_valor.get()
            genero = genero_combobox.get()
            endereco = endereco_valor.get()
            obs = obs_entry.get(0.0, END)

            if (nome =="" or contato =="" or idade =="" or endereco ==""):
                messagebox.showerror("Sistema", "ERRO!\nPor favor, preencha todos os campos!") 

            else:
                ficheiro = openpyxl.load_workbook("Clientes.xlsx")
                folha = ficheiro.active
                folha.cell(column=1, row=folha.max_row+1, value=nome)
                folha.cell(column=2, row=folha.max_row, value=contato) 
                folha.cell(column=3, row=folha.max_row, value=idade) 
                folha.cell(column=4, row=folha.max_row, value=genero) 
                folha.cell(column=5, row=folha.max_row, value=endereco) 
                folha.cell(column=6, row=folha.max_row, value=obs)

            ficheiro.save(r"Clientes.xlsx")
            messagebox.showinfo("Sistema", "Dados salvos com sucesso!")



        def clear():
            nome_valor.set("")
            contato_valor.set("")
            idade_valor.set("")
            endereco_valor.set("")
            obs_entry.delete(0.0, END)
        

        #variaveis de texto
        nome_valor = StringVar()
        contato_valor = StringVar()
        idade_valor = StringVar()
        endereco_valor = StringVar()

        #entrys
        nome_entry = ctk.CTkEntry(self, width=350, textvariable=nome_valor, font=("Century Gothic", 16), fg_color="#ffffff")
        contato_entry = ctk.CTkEntry(self, width=200, textvariable=contato_valor, font=("Century Gothic", 16), fg_color="#ffffff")
        idade_entry = ctk.CTkEntry(self, width=150, textvariable=idade_valor, font=("Century Gothic", 16), fg_color="#ffffff")
        endereco_entry = ctk.CTkEntry(self, textvariable=endereco_valor, width=200, font=("Century Gothic", 16), fg_color="#ffffff")

        #combobox
        genero_combobox = ctk.CTkComboBox(self, values=["Masculino", "Feminino"], font=("Century Gothic bold", 14), width=150)
        genero_combobox.set("Masculino")

        #entrada de observacoes
        obs_entry = ctk.CTkTextbox(self, width=470, height=150, font=("ariel", 18), border_color="#aaa", border_width=2, fg_color="#ffffff")

        #labels
        label_nome = ctk.CTkLabel(self, text="Nome completo:", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        label_contato = ctk.CTkLabel(self, text="Contato", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        label_idade = ctk.CTkLabel(self, text="Idade", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        label_genero = ctk.CTkLabel(self, text="Gênero", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        label_endereco = ctk.CTkLabel(self, text="Endereço", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        label_obs = ctk.CTkLabel(self, text="Observações", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])

        botao_submit = ctk.CTkButton(self, text="Salvar dados".upper(), command=submit, fg_color="#151").place(x=300, y=420)
        botao_clear = ctk.CTkButton(self, text="Limpar dados".upper(), command=clear, fg_color="#555").place(x=500, y=420)


        #posicionando os elementos na janela
        label_nome.place(x=50, y=120)
        nome_entry.place(x=50, y=150)

        label_contato.place(x=450, y=120)
        contato_entry.place(x=450, y=150)

        label_idade.place(x=300, y=190)
        idade_entry.place(x=300, y=220)

        label_genero.place(x=500, y=190)
        genero_combobox.place(x=500, y=220)

        label_endereco.place(x=50, y=190)
        endereco_entry.place(x=50, y=220)

        label_obs.place(x=50, y=260)
        obs_entry.place(x=180, y=260)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)

if __name__ == "__main__":
    app = App()
    app.mainloop()