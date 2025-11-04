import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import sqlite3
from datetime import datetime
from openpyxl import Workbook
import pandas as pd



# Todos os bancos de dados e as tabelas

def conectar():
    conn = sqlite3.connect("Clientes.db")
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Clientes (
            id INTEGER PRIMARY KEY,
            nome TEXT NOT NULL,
            telefone TEXT NOT NULL,
            cidade TEXT NOT NULL,
            cpf TEXT NOT NULL
        );
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS OS (
            numero INTEGER PRIMARY KEY,
            data TEXT NOT NULL,
            cliente_id INTEGER NOT NULL,
            cidade TEXT NOT NULL,
            telefone TEXT NOT NULL,
            equipamentos TEXT NOT NULL,
            preco_total REAL NOT NULL,
            status TEXT NOT NULL DEFAULT 'Aguardando',
            FOREIGN KEY(cliente_id) REFERENCES Clientes(id)
        );
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Pecas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            os_numero INTEGER NOT NULL,
            nome_peca TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            preco_unitario REAL NOT NULL,
            preco_total REAL NOT NULL,
            FOREIGN KEY(os_numero) REFERENCES OS(numero)
    );
''')


    conn.commit()
    return conn, cursor
def exportar_clientes_excel():
    conn = sqlite3.connect("Clientes.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Clientes")
    dados = cursor.fetchall()
    conn.close()

    if not dados:
        messagebox.showinfo("Exportar Clientes", "Não há clientes cadastrados para exportar.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Clientes"

    # Colunas:
    colunas = ["ID", "Nome", "Telefone", "Cidade", "CPF"]
    ws.append(colunas)

    # Todos os dados
    for linha in dados:
        ws.append(linha)

    nome_arquivo = "Clientes_Exportados.xlsx"
    wb.save(nome_arquivo)
    messagebox.showinfo("Exportar Clientes", f"Clientes exportados com sucesso!\nArquivo: {nome_arquivo}")


def exportar_os_excel():
    conn = sqlite3.connect("Clientes.db")
    cursor = conn.cursor()
    cursor.execute('''
        SELECT OS.numero, OS.data, Clientes.nome, OS.cidade, OS.telefone, OS.equipamentos, OS.preco_total, OS.status
        FROM OS
        JOIN Clientes ON OS.cliente_id = Clientes.id
    ''')
    dados = cursor.fetchall()
    conn.close()

    if not dados:
        messagebox.showinfo("Exportar OS", "Não há ordens de serviço cadastradas para exportar.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Ordens de Serviço"

    colunas = ["Número", "Data", "Cliente", "Cidade", "Telefone", "Equipamento", "Total", "Status"]
    ws.append(colunas)

    for linha in dados:
        ws.append(linha)

    nome_arquivo = "OrdensServico_Exportadas.xlsx"
    wb.save(nome_arquivo)
    messagebox.showinfo("Exportar OS", f"Ordens de serviço exportadas com sucesso!\nArquivo: {nome_arquivo}")



# Parte que separa as janelas cada botão abre uma 

def janela_cadastrar_cliente():
    win = tk.Toplevel()
    win.title("Cadastrar Cliente")

    tk.Label(win, text="Nome:").grid(row=0, column=0, padx=5, pady=5)
    nome = tk.Entry(win)
    nome.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(win, text="Telefone:").grid(row=1, column=0, padx=5, pady=5)
    telefone = tk.Entry(win)
    telefone.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(win, text="Cidade:").grid(row=2, column=0, padx=5, pady=5)
    cidade = tk.Entry(win)
    cidade.grid(row=2, column=1, padx=5, pady=5)

    tk.Label(win, text="CPF:").grid(row=3, column=0, padx=5, pady=5)
    cpf = tk.Entry(win)
    cpf.grid(row=3, column=1, padx=5, pady=5)

    def salvar():
        nome_val = nome.get().strip()
        telefone_val = telefone.get().strip()
        cidade_val = cidade.get().strip()
        cpf_val = cpf.get().strip()

        if not (nome_val and telefone_val and cidade_val and cpf_val):
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return

        # Validar o cpf e telefone tendo 11 digitos
        if not cpf_val.isdigit() or len(cpf_val) != 11:
            messagebox.showerror("Erro", "O CPF deve conter exatamente 11 números.")
            return

        if not telefone_val.isdigit() or len(telefone_val) != 11:
            messagebox.showerror("Erro", "O telefone deve conter exatamente 11 números (DDD + número).")
            return

        conn, cursor = conectar()

        cursor.execute("SELECT id FROM Clientes WHERE cpf = ?", (cpf_val,))
        cpf_existente = cursor.fetchone()

        cursor.execute("SELECT id FROM Clientes WHERE telefone = ?", (telefone_val,))
        telefone_existente = cursor.fetchone()

        if cpf_existente:
            messagebox.showerror("Erro", "Já existe um cliente cadastrado com esse CPF.")
            conn.close()
            return

        if telefone_existente:
            messagebox.showerror("Erro", "Já existe um cliente cadastrado com esse telefone.")
            conn.close()
            return

        cursor.execute("INSERT INTO Clientes (nome, telefone, cidade, cpf) VALUES (?, ?, ?, ?)",
                    (nome_val, telefone_val, cidade_val, cpf_val))
        conn.commit()
        conn.close()

        messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")
        win.destroy()


    tk.Button(win, text="Salvar", command=salvar).grid(row=4, column=0, columnspan=2, pady=10)


def janela_listar_clientes():
    win = tk.Toplevel()
    win.title("Lista de Clientes")
    tree = ttk.Treeview(win, columns=("ID", "Nome", "Telefone", "Cidade", "CPF", "OS"), show="headings")
    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    tree.pack(fill="both", expand=True)

    conn, cursor = conectar()
    cursor.execute("SELECT * FROM Clientes")
    clientes = cursor.fetchall()
    btn_exportar_clientes = tk.Button(win, text="Exportar Clientes", command=exportar_clientes_excel, bg="#4CAF50", fg="white")
    btn_exportar_clientes.pack(pady=10)

    for cliente in clientes:
        cursor.execute("SELECT numero FROM OS WHERE cliente_id = ?", (cliente[0],))
        os_list = [str(x[0]) for x in cursor.fetchall()]
        tree.insert("", tk.END, values=cliente + (", ".join(os_list),))

    conn.close()

def janela_remover_cliente():
    win = tk.Toplevel()
    win.title("Remover Cliente")

    tk.Label(win, text="ID do Cliente:").pack()
    id_entry = tk.Entry(win)
    id_entry.pack()

    def remover():
        conn, cursor = conectar()
        cursor.execute("DELETE FROM Clientes WHERE id = ?", (id_entry.get(),))
        if cursor.rowcount == 0:
            messagebox.showerror("Erro", "ID não encontrado.")
        else:
            conn.commit()
            messagebox.showinfo("Sucesso", "Cliente removido.")
        conn.close()
        win.destroy()

    tk.Button(win, text="Remover", command=remover).pack(pady=5)

def janela_criar_os():
    win = tk.Toplevel()
    win.title("Criar Ordem de Serviço")

    campos = ["Número", "ID do Cliente", "Equipamentos"]
    entries = []

    for i, campo in enumerate(campos):
        tk.Label(win, text=campo).grid(row=i, column=0, padx=5, pady=5)
        ent = tk.Entry(win)
        ent.grid(row=i, column=1, padx=5, pady=5)
        entries.append(ent)

    pecas = []


    def salvar():
        try:
            numero = int(entries[0].get())
            cliente_id = int(entries[1].get())
            equipamentos = entries[2].get()

            conn, cursor = conectar()

            # Linka a cidade e telefone pelo id
            cursor.execute("SELECT cidade, telefone FROM Clientes WHERE id = ?", (cliente_id,))
            dados_cliente = cursor.fetchone()

            if not dados_cliente:
                messagebox.showerror("Erro", "Cliente não encontrado!")
                conn.close()
                return

            cidade, telefone = dados_cliente
            preco_total_os = sum(item[3] for item in pecas)
            data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cursor.execute("""
                INSERT INTO OS (numero, data, cliente_id, cidade, telefone, equipamentos, preco_total)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (numero, data, cliente_id, cidade, telefone, equipamentos, preco_total_os))

            # Coloca as peças
            for nome, qtd, unit, total in pecas:
                cursor.execute("""
                    INSERT INTO Pecas (os_numero, nome_peca, preco)
                    VALUES (?, ?, ?)
                """, (numero, nome, total))

            conn.commit()
            conn.close()
            messagebox.showinfo("Sucesso", "OS criada com sucesso!")
            win.destroy()

        except Exception as e:
            messagebox.showerror("Erro", f"Verifique os dados inseridos.\n{e}")

    tk.Button(win, text="Salvar OS", command=salvar).grid(row=len(campos), column=0, pady=10)
def janela_editar_os():
    win = tk.Toplevel()
    win.title("Editar Ordem de Serviço")

    tk.Label(win, text="Número da OS:").grid(row=0, column=0, padx=5, pady=5)
    numero_entry = tk.Entry(win)
    numero_entry.grid(row=0, column=1, padx=5, pady=5)

    frame_pecas = tk.Frame(win)
    frame_pecas.grid(row=2, column=0, columnspan=2, pady=10)

    pecas_tree = ttk.Treeview(frame_pecas, columns=("ID", "Nome", "Quantidade", "Unitário", "Total"), show="headings")
    for col in pecas_tree["columns"]:
        pecas_tree.heading(col, text=col)
        pecas_tree.column(col, width=100)
    pecas_tree.pack()

    def carregar_pecas():
        for item in pecas_tree.get_children():
            pecas_tree.delete(item)

        numero = numero_entry.get().strip()
        if not numero:
            messagebox.showerror("Erro", "Informe o número da OS.")
            return

        conn, cursor = conectar()
        cursor.execute("SELECT nome_peca, quantidade, preco_unitario, preco_total, id FROM Pecas WHERE os_numero = ?", (numero,))
        pecas = cursor.fetchall()
        conn.close()

        if not pecas:
            messagebox.showinfo("Info", "Nenhuma peça cadastrada para esta OS.")
        else:
            for nome, qtd, unit, total, pid in pecas:
                pecas_tree.insert("", tk.END, values=(pid, nome, qtd, unit, total))

    def adicionar_peca():
        numero = numero_entry.get().strip()
        if not numero:
            messagebox.showerror("Erro", "Informe o número da OS antes de adicionar peças.")
            return

        nome = simpledialog.askstring("Peça", "Nome da peça:", parent=win).capitalize()
        if not nome:
            return
        try:
            qtd = int(simpledialog.askstring("Quantidade", "Quantidade:", parent=win))
            unit = float(simpledialog.askstring("Preço unitário", "Preço unitário:", parent=win))
            total = qtd * unit
        except:
            messagebox.showerror("Erro", "Valores inválidos.")
            return

        conn, cursor = conectar()
        cursor.execute("INSERT INTO Pecas (os_numero, nome_peca, quantidade, preco_unitario, preco_total) VALUES (?, ?, ?, ?, ?)",
                       (numero, nome, qtd, unit, total))
        conn.commit()

        
        cursor.execute("SELECT SUM(preco_total) FROM Pecas WHERE os_numero = ?", (numero,))
        total_os = cursor.fetchone()[0] or 0
        cursor.execute("UPDATE OS SET preco_total = ? WHERE numero = ?", (total_os, numero))

        conn.commit()
        conn.close()

        carregar_pecas()
        messagebox.showinfo("Sucesso", "Peça adicionada com sucesso!")

    def editar_peca():
        selecionado = pecas_tree.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma peça para editar.")
            return

        item = pecas_tree.item(selecionado)
        pid, nome, qtd, unit, total = item["values"]

        novo_nome = simpledialog.askstring("Editar Peça", "Nome da peça:", initialvalue=nome, parent=win)
        try:
            nova_qtd = int(simpledialog.askstring("Editar Quantidade", "Quantidade:", initialvalue=qtd, parent=win))
            novo_unit = float(simpledialog.askstring("Editar Preço Unitário", "Preço unitário:", initialvalue=unit, parent=win))
        except:
            messagebox.showerror("Erro", "Valores inválidos.")
            return
        novo_total = nova_qtd * novo_unit

        conn, cursor = conectar()
        cursor.execute("""
            UPDATE Pecas
            SET nome_peca = ?, quantidade = ?, preco_unitario = ?, preco_total = ?
            WHERE id = ?
        """, (novo_nome, nova_qtd, novo_unit, novo_total, pid))

       
        numero = numero_entry.get().strip()
        cursor.execute("SELECT SUM(preco_total) FROM Pecas WHERE os_numero = ?", (numero,))
        total_os = cursor.fetchone()[0] or 0
        cursor.execute("UPDATE OS SET preco_total = ? WHERE numero = ?", (total_os, numero))

        conn.commit()
        conn.close()

        carregar_pecas()
        messagebox.showinfo("Sucesso", "Peça atualizada com sucesso!")

    def remover_peca():
        selecionado = pecas_tree.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma peça para remover.")
            return

        item = pecas_tree.item(selecionado)
        pid = item["values"][0]

        conn, cursor = conectar()
        cursor.execute("DELETE FROM Pecas WHERE id = ?", (pid,))

    
        numero = numero_entry.get().strip()
        cursor.execute("SELECT SUM(preco_total) FROM Pecas WHERE os_numero = ?", (numero,))
        total_os = cursor.fetchone()[0] or 0
        cursor.execute("UPDATE OS SET preco_total = ? WHERE numero = ?", (total_os, numero))

        conn.commit()
        conn.close()

        carregar_pecas()
        messagebox.showinfo("Sucesso", "Peça removida com sucesso!")

    
    tk.Button(win, text="Carregar Peças", command=carregar_pecas, bg="#2196F3", fg="white").grid(row=1, column=0, pady=5)
    tk.Button(win, text="Adicionar Peça", command=adicionar_peca, bg="#4CAF50", fg="white").grid(row=3, column=0, pady=5)
    tk.Button(win, text="Editar Peça", command=editar_peca, bg="#FFC107", fg="black").grid(row=3, column=1, pady=5)
    tk.Button(win, text="Remover Peça", command=remover_peca, bg="#F44336", fg="white").grid(row=4, column=0, columnspan=2, pady=10)



def janela_listar_os():
    win = tk.Toplevel()
    win.title("Lista de OS")
    win.geometry("850x600")

    
    top_frame = tk.Frame(win)
    top_frame.pack(fill="x", pady=10)

    tk.Label(top_frame, text="📋 Lista de Ordens de Serviço", font=("Arial", 14, "bold")).pack(side="left", padx=10)

   
    def exportar_excel():
        conn, cursor = conectar()
        cursor.execute('''
            SELECT OS.numero, OS.data, Clientes.nome, OS.cidade, OS.telefone, OS.equipamentos, OS.preco_total, OS.status
            FROM OS
            JOIN Clientes ON OS.cliente_id = Clientes.id
        ''')
        df = pd.DataFrame(cursor.fetchall(), columns=["Número", "Data", "Cliente", "Cidade", "Telefone", "Equipamento", "Total", "Status"])
        conn.close()

        if df.empty:
            messagebox.showwarning("Aviso", "Não há dados para exportar.")
            return

        nome_arquivo = f"Relatorio_OS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(nome_arquivo, index=False)
        messagebox.showinfo("Sucesso", f"Relatório exportado como {nome_arquivo}")
        btn_exportar_os = tk.Button(janela_listar_os, text="Exportar OS", command=exportar_os_excel)
        btn_exportar_os.pack(pady=5)


    tk.Button(top_frame, text="Exportar para Excel", command=exportar_excel, bg="#4CAF50", fg="white").pack(side="right", padx=10)
    

   
    canvas = tk.Canvas(win)
    scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

 
    conn, cursor = conectar()
    cursor.execute('''
        SELECT OS.numero, OS.data, Clientes.nome, OS.cidade, OS.telefone, OS.equipamentos, OS.preco_total, OS.status
        FROM OS
        JOIN Clientes ON OS.cliente_id = Clientes.id
    ''')
    ordens = cursor.fetchall()
    conn.close()

    if not ordens:
        tk.Label(scroll_frame, text="Nenhuma OS cadastrada ainda.", font=("Arial", 12, "italic"), fg="gray").pack(pady=20)
        return

   
    for os_item in ordens:
        card = tk.Frame(scroll_frame, bg="white", relief="groove", borderwidth=2)
        card.pack(fill="x", padx=15, pady=10)

        numero, data, cliente, cidade, telefone, equip, total, status = os_item
       
        conn, cursor = conectar()
        cursor.execute("SELECT nome_peca FROM Pecas WHERE os_numero = ?", (numero,))
        pecas_lista = [linha[0] for linha in cursor.fetchall()]
        conn.close()

       
        if pecas_lista:
            pecas_texto = ", ".join(pecas_lista)
        else:
            pecas_texto = "Nenhuma peça cadastrada"


        
        tk.Label(card, text=f"OS Nº {numero} — {cliente}", bg="white", font=("Arial", 12, "bold"), anchor="w").pack(fill="x", padx=10, pady=5)

    
        tk.Label(card, text=f" Data: {data}", bg="white", anchor="w").pack(fill="x", padx=15)
        tk.Label(card, text=f" Cidade: {cidade}", bg="white", anchor="w").pack(fill="x", padx=15)
        tk.Label(card, text=f" Telefone: {telefone}", bg="white", anchor="w").pack(fill="x", padx=15)
        tk.Label(card, text=f" Equipamento: {equip}", bg="white", anchor="w").pack(fill="x", padx=15)
        tk.Label(card, text=f" Peças: {pecas_texto}", bg="white", anchor="w").pack(fill="x", padx=15)
        tk.Label(card, text=f" Total: R${total:.2f}", bg="white", anchor="w").pack(fill="x", padx=15)
        tk.Label(card, text=f" Status: {status}", bg="white", anchor="w", fg="blue").pack(fill="x", padx=15, pady=(0,10))

        


def janela_atualizar_status():
    win = tk.Toplevel()
    win.title("Atualizar Status da OS")

    tk.Label(win, text="Número da OS:").pack()
    numero = tk.Entry(win)
    numero.pack()

    tk.Label(win, text="Novo Status:").pack()
    status_var = tk.StringVar()
    status_menu = ttk.Combobox(win, textvariable=status_var)
    status_menu['values'] = ["Aguardando", "Orçamento realizado", "Orçamento autorizado", "Equipamento pronto", "Entrega realizada"]
    status_menu.pack()

    def atualizar():
        conn, cursor = conectar()
        cursor.execute("UPDATE OS SET status = ? WHERE numero = ?", (status_var.get(), numero.get()))
        if cursor.rowcount == 0:
            messagebox.showerror("Erro", "Número de OS não encontrado.")
        else:
            conn.commit()
            messagebox.showinfo("Sucesso", "Status atualizado.")
        conn.close()
        win.destroy()

    tk.Button(win, text="Atualizar", command=atualizar).pack(pady=5)

def janela_procurar_cliente():
    win = tk.Toplevel()
    win.title("Procurar Cliente por Nome")

    tk.Label(win, text="Nome do Cliente:").pack(pady=5)
    nome_entry = tk.Entry(win)
    nome_entry.pack(pady=5)

    
    tree = ttk.Treeview(win, columns=("ID", "Nome", "Telefone", "Cidade", "CPF"), show="headings")
    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    tree.pack(fill="both", expand=True, pady=10)

    def procurar():
        
        for item in tree.get_children():
            tree.delete(item)

        nome = nome_entry.get().strip()

        if not nome:
            messagebox.showwarning("Aviso", "Digite um nome para procurar.")
            return

        conn, cursor = conectar()
        cursor.execute("SELECT * FROM Clientes WHERE nome LIKE ?", ('%' + nome + '%',))
        resultados = cursor.fetchall()
        conn.close()

        if not resultados:
            messagebox.showinfo("Resultado", "Nenhum cliente encontrado.")
        else:
            for row in resultados:
                tree.insert("", tk.END, values=row)

  
    tk.Button(win, text="Procurar", command=procurar).pack(pady=5)

    

def janela_remover_os():
    win = tk.Toplevel()
    win.title("Remover Ordem de Serviço")

    tk.Label(win, text="Número da OS:").pack(pady=5)
    numero_entry = tk.Entry(win)
    numero_entry.pack(pady=5)

    def remover():
        numero = numero_entry.get().strip()

        if not numero:
            messagebox.showerror("Erro", "Informe o número da OS.")
            return

        conn, cursor = conectar()
        cursor.execute("SELECT numero FROM OS WHERE numero = ?", (numero,))
        existe = cursor.fetchone()

        if not existe:
            messagebox.showerror("Erro", "OS não encontrada.")
            conn.close()
            return

        
        if not messagebox.askyesno("Confirmar", f"Tem certeza que deseja remover a OS Nº {numero}?"):
            conn.close()
            return

        
        cursor.execute("DELETE FROM Pecas WHERE os_numero = ?", (numero,))
        cursor.execute("DELETE FROM OS WHERE numero = ?", (numero,))

        conn.commit()
        conn.close()

        messagebox.showinfo("Sucesso", f"OS Nº {numero} removida com sucesso!")
        win.destroy()

    tk.Button(win, text="Remover", command=remover, bg="#F44336", fg="white").pack(pady=10)





# Front principal


janela = tk.Tk()
janela.title("Sistema de Clientes e OS")
janela.geometry("400x600")



tk.Label(janela, text="Menu Principal", font=("Arial", 20, "bold")).pack(pady=20)

botoes = [
    ("Cadastrar Cliente", janela_cadastrar_cliente),
    ("Listar Clientes", janela_listar_clientes),
    ("Procurar Cliente", janela_procurar_cliente),
    ("Remover Cliente", janela_remover_cliente),
    ("Criar OS", janela_criar_os),
    ("Listar OS", janela_listar_os),
    ("Remover OS", janela_remover_os),
    ("Atualizar Status da OS", janela_atualizar_status),
    ("Editar OS", janela_editar_os)

]

for texto, comando in botoes:
    tk.Button(janela, text=texto, width=30, command=comando, bg="gray53").pack(pady=5)

janela.mainloop()
