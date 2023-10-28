import tkinter as tk
from tkinter import filedialog
import coelhos

def open_sheet():
    root.filename = filedialog.askopenfilename(initialdir=".", title="Selecione a Planilha", filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
    if root.filename:
        input_update()

def set_txt_coelhos(txt):
    txt_coelhos.config(state="normal")
    txt_coelhos.delete("1.0", tk.END)
    txt_coelhos.insert("1.0", txt)
    txt_coelhos.config(state="disabled")

def set_aviso(msg, success=None):
    if success is None:
        ent_aviso.config(fg="black")
    elif success:
        ent_aviso.config(fg="green")
    else:
        ent_aviso.config(fg="red")
    ent_aviso.config(state="normal")
    ent_aviso.delete(0, tk.END)
    ent_aviso.insert(0, msg)
    ent_aviso.config(state="disabled")

def change_sheet(sheet):
    global gp
    set_txt_coelhos("")
    set_aviso("Carregando planilha...", success=None)

    print(sheet)

    filename = root.filename

    try:
        gp = coelhos.GrauParentesco(filename, sheet)
        set_txt_coelhos(gp.get_entradadados)
        set_aviso("Planilha carregada com sucesso!", success=True)

    except Exception as e:
        set_aviso(str(e), success=False)
    
    return tk._setit(var, sheet)


def input_update():
    edt_input.config(state="normal")
    edt_input.delete(0, tk.END)
    edt_input.insert(0, root.filename)
    edt_input.config(state="disabled")

    sheets = coelhos.get_sheets(root.filename)
    
    menu = dd_sheet["menu"]
    menu.delete(0, "end")
    for sheet in sheets:
        menu.add_command(
            label=sheet, 
            command=change_sheet(sheet)
        )
    dd_sheet["menu"].invoke(0)



root = tk.Tk()
root.title("Gerador de Coeficiente de Parentesco")
root.geometry("500x500")


edt_input = tk.Entry(root, state="disabled")
edt_input.pack(fill="x")

btn_input = tk.Button(root, text="Abrir", command=open_sheet)
btn_input.pack()

var = tk.StringVar()
dd_sheet = tk.OptionMenu(root, var, "option1", "option2", "option3")
dd_sheet.pack()

ent_aviso = tk.Entry(root, state="disabled")
ent_aviso.pack(fill="both")

#center 
txt_coelhos = tk.Text(root, state="disabled", wrap="none")
txt_coelhos.pack(fill="both", expand=True)

ent_coelhos = tk.Entry(root)
ent_coelhos.insert(0, "Grau de Parentesco")
ent_coelhos.pack()

btn_calcular = tk.Button(root, text="Calcular")
btn_calcular.pack()


# lbl_input.grid(column=0, row=0)
# edt_input.grid(column=1, row=0)
# btn_input.grid(column=2, row=0)
# lbl_sheet.grid(column=0, row=1)
# dd_sheet.grid(column=1, row=1)

root.mainloop()