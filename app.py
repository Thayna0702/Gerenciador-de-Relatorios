# Biblioteca datas e horários do relógio do sistema
import datetime

# Biblioteca para caminhos de arquivos no computador
import os

# Biblioteca do Banco de Dados
import sqlite3

# Biblioteca responsável pelo design
import customtkinter as ctk

# Biblioteca que transforma em uma planilha Excel
import pandas as pd

# Ferramentas extras do Tkinter
from tkinter import messagebox, ttk

# CONFIGURANDO O BANCO DE DADOS

# Conectando ao arquivo de banco de dados
conn = sqlite3.connect("relatorios.db")
# O cursor que executes as instruções dentro do banco de dados:
cursor = conn.cursor()

# Cria a tabela de relatórios com as colunas necessárias
cursor.execute(
    """
    CREATE TABLE IF NOT EXISTS relatorios (
    id INTEGER PRIMARY KEY AUTOINCREMENT, -- Identificador único para cada relatório
    dia_faturamento TEXT,                 -- Dia do faturamento
    empresa TEXT,                         -- Nome da Empresa
    contrato TEXT,                        -- Número do Contrato
    descricao TEXT,                       -- Descrição
    chegada TEXT,                         -- Data em que chegou
    prazo TEXT,                           -- Prazo limite final
    feito TEXT DEFAULT '',                -- Guarda Data e Hora quando for produzido
    enviado TEXT DEFAULT ''               -- Guarda Data e Hora quando for enviado
    )
    """
)
# Salva as alterações da tabela
conn.commit()

# CÁLCULO DOS PRAZOS


def calcular_dias_para_faturamento(dia_faturamento_str):
    try:
        dia_fat = int(dia_faturamento_str)
        hoje = datetime.date.today()

        # Prox faturamento. prox mês
        if dia_fat <= hoje.day:
            proximo_mes = hoje.month + 1 if hoje.month < 12 else 1
            proximo_ano = hoje.year if hoje.month < 12 else hoje.year + 1
            data_fat = datetime.date(proximo_ano, proximo_mes, dia_fat)
        else:
            # Faturamento esse mês
            data_fat = datetime.date(hoje.year, hoje.month, dia_fat)

        return (data_fat - hoje).days
    except ValueError:
        return 999  # Se for xxx, joga para o fim da fila de prioridade


def converter_para_data(data_str):
    try:
        return datetime.datetime.strptime(data_str, "%d/%m/%Y").date()
    except:
        return datetime.date(2002, 1, 1)


# ESTILIZAÇÃO


class AppRelatorios(ctk.CTk):
    def __init__(self):
        super().__init__()

        # dimenção
        self.title("Gerenciador de Relatórios")
        self.geometry("1150x700")

        # cadastro
        self.frame_cadastro = ctk.CTkFrame(self, width=300, corner_radius=15)
        self.frame_cadastro.pack(side="left", fill="y", padx=15, pady=15, expand=False)

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_cadastro,
            text="Novo Relatório",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        self.lbl_titulo.pack(padx=20, pady=15)

        # Inputs de Texto (Campos para digitação)
        self.txt_empresa = ctk.CTkEntry(
            self.frame_cadastro, placeholder_text="Empresa", width=240
        )
        self.txt_empresa.pack(padx=20, pady=8)

        self.txt_dia_fat = ctk.CTkEntry(
            self.frame_cadastro,
            placeholder_text="Dia de Faturamento",
            width=240,
        )
        self.txt_dia_fat.pack(padx=20, pady=8)

        self.txt_contrato = ctk.CTkEntry(
            self.frame_cadastro, placeholder_text="Contrato", width=240
        )
        self.txt_contrato.pack(padx=20, pady=8)

        self.txt_desc = ctk.CTkEntry(
            self.frame_cadastro, placeholder_text="Código", width=240
        )
        self.txt_desc.pack(padx=20, pady=8)

        self.txt_chegada = ctk.CTkEntry(
            self.frame_cadastro,
            placeholder_text="Data Chegada",
            width=240,
        )
        self.txt_chegada.pack(padx=20, pady=8)

        self.txt_prazo = ctk.CTkEntry(
            self.frame_cadastro,
            placeholder_text="Data Prazo Final",
            width=240,
        )
        self.txt_prazo.pack(padx=20, pady=8)

        self.btn_salvar = ctk.CTkButton(
            self.frame_cadastro,
            text="➕ Cadastrar Relatório",
            fg_color="#1c3a6e",
            command=self.salvar_relatorio,
        )
        self.btn_salvar.pack(padx=20, pady=10)

        # botão de fato
        self.btn_deletar = ctk.CTkButton(
            self.frame_cadastro,
            text="🗑️ Deletar Selecionado",
            fg_color="#c92a2a",
            command=self.deletar_relatorio,
        )
        self.btn_deletar.pack(padx=20, pady=10)

        # abas
        self.abas = ctk.CTkTabview(self, corner_radius=15)
        self.abas.pack(side="right", fill="both", padx=15, pady=15, expand=True)

        self.tab_pendentes = self.abas.add("📋 Fila")
        self.tab_historico = self.abas.add("📜 Histórico")

        # cores
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#2a2d2e",
            foreground="white",
            rowheight=26,
            fieldbackground="#2a2d2e",
            bordercolor="#343638",
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.map("Treeview", background=[("selected", "#4374ac")])
        style.configure(
            "Treeview.Heading",
            background="#1f538d",
            foreground="white",
            relief="flat",
        )

        # estilização fila

        self.lbl_lista = ctk.CTkLabel(
            self.tab_pendentes,
            text="ORDEM DOS RELATÓRIOS",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.lbl_lista.pack(padx=20, pady=10)

        # Altere esta linha para incluir "Descrição"
        self.tabela_pendentes = ttk.Treeview(
            self.tab_pendentes,
            columns=("ID", "Empresa", "Fat.", "Prazo", "Contrato", "Descrição", "Status Feito"),
            show="headings",
        )
        self.tabela_pendentes.heading("ID", text="ID")
        self.tabela_pendentes.heading("Empresa", text="Empresa")
        self.tabela_pendentes.heading("Fat.", text="Dia Fat.")
        self.tabela_pendentes.heading("Prazo", text="Prazo Final")
        self.tabela_pendentes.heading("Contrato", text="Contrato")
        self.tabela_pendentes.heading("Descrição", text="Descrição") # <-- Adicione esta linha
        self.tabela_pendentes.heading("Status Feito", text="Relatório Pronto?")

        self.tabela_pendentes.column("ID", width=40, anchor="center")
        self.tabela_pendentes.column("Empresa", width=150)
        self.tabela_pendentes.column("Fat.", width=70, anchor="center")
        self.tabela_pendentes.column("Prazo", width=100, anchor="center")
        self.tabela_pendentes.column("Contrato", width=90)
        self.tabela_pendentes.column("Descrição", width=120)        # <-- Adicione esta linha
        self.tabela_pendentes.column("Status Feito", width=150, anchor="center")
        self.tabela_pendentes.pack(padx=20, fill="both", expand=True)

        # alerta de urgencia
        self.tabela_pendentes.tag_configure(
            "urgente_critico", background="#661414", foreground="white"
        )  # faturamento em até 3 dias
        self.tabela_pendentes.tag_configure(
            "urgente_atencao", background="#5c440b", foreground="white"
        )  # faturamento em até 7 dias

        # caixa de botões
        self.frame_botoes_acao = ctk.CTkFrame(
            self.tab_pendentes, fg_color="transparent"
        )
        self.frame_botoes_acao.pack(padx=20, pady=15, fill="x")

        # Botão feito
        self.btn_marcar_feito = ctk.CTkButton(
            self.frame_botoes_acao,
            text="🛠️ Marcar como FEITO (Pronto)",
            fg_color="#1c3a6e",
            hover_color="#1c3a6e",
            command=self.marcar_como_feito,
        )
        self.btn_marcar_feito.pack(side="left", padx=(0, 10), expand=True, fill="x")

        # Botão enviado
        self.btn_marcar_enviado = ctk.CTkButton(
            self.frame_botoes_acao,
            text="🚀 Marcar como ENVIADO (Arquivar)",
            fg_color="#1c3a6e",
            hover_color="#1c3a6e",
            command=self.marcar_como_enviado,
        )
        self.btn_marcar_enviado.pack(side="right", padx=(10, 0), expand=True, fill="x")

        # aba historico
        self.frame_busca = ctk.CTkFrame(self.tab_historico, fg_color="transparent")
        self.frame_busca.pack(padx=20, pady=10, fill="x")

        # barra de pesquisa
        self.txt_pesquisa = ctk.CTkEntry(
            self.frame_busca,
            placeholder_text="🔍 Filtre o histórico por Empresa, Contrato ou Descrição...",
        )
        self.txt_pesquisa.pack(side="left", padx=(0, 10), fill="x", expand=True)
        self.txt_pesquisa.bind("<KeyRelease>", self.filtrar_historico)

        # Exportar pro excel
        self.btn_exportar = ctk.CTkButton(
            self.frame_busca,
            text="📊 Exportar para Excel",
            fg_color="#1c3a6e",
            hover_color="#1c3a6e",
            width=160,
            command=self.exportar_para_excel,
        )
        self.btn_exportar.pack(side="right")

        # Tabela do histórico
        self.tabela_historico = ttk.Treeview(
            self.tab_historico,
            columns=(
                "ID",
                "Empresa",
                "Contrato",
                "Descrição",
                "Horário Feito",
                "Horário Enviado",
            ),
            show="headings",
        )
        self.tabela_historico.heading("ID", text="ID")
        self.tabela_historico.heading("Empresa", text="Empresa")
        self.tabela_historico.heading("Contrato", text="Contrato")
        self.tabela_historico.heading("Descrição", text="Descrição")
        self.tabela_historico.heading("Horário Feito", text="Feito Em")
        self.tabela_historico.heading("Horário Enviado", text="Enviado Em")

        self.tabela_historico.column("ID", width=40, anchor="center")
        self.tabela_historico.column("Empresa", width=140)
        self.tabela_historico.column("Contrato", width=90)
        self.tabela_historico.column("Descrição", width=150)
        self.tabela_historico.column("Horário Feito", width=140, anchor="center")
        self.tabela_historico.column("Horário Enviado", width=140, anchor="center")
        self.tabela_historico.pack(padx=20, pady=(0, 20), fill="both", expand=True)

        # deixa os arquivos anteriores salvos
        self.atualizar_tabela_pendentes()
        self.atualizar_tabela_historico()

    # COMANDOS DO SISTEMA
    # botão delete

    def deletar_relatorio(self):
        aba_atual = self.abas.get()

        if aba_atual == "📋 Fila":
            tabela = self.tabela_pendentes
        else:
            tabela = self.tabela_historico

        item_selecionado = tabela.selection()    

        if not item_selecionado:
            messagebox.showwarning(
                "Aviso", "Selecione um registro na tabela para deletá-lo!"
            )
            return

        valores = tabela.item(item_selecionado, "values")
        id_deletar = valores[0]
        empresa_nome = valores[1]

        confirmar = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Tem certeza que deseja deletar permanentemente o relatório da empresa '{empresa_nome}'?",
        )

        if confirmar:
            cursor.execute("DELETE FROM relatorios WHERE id = ?", (id_deletar,))
            conn.commit()

            messagebox.showinfo("Sucesso", "Registro deletado com sucesso!")

            self.atualizar_tabela_pendentes()
            self.atualizar_tabela_historico()

    def salvar_relatorio(self):
        if not self.txt_empresa.get() or not self.txt_dia_fat.get():
            messagebox.showwarning(
                "Aviso", "Preencha ao menos o Nome da Empresa e Dia de Faturamento!"
            )
            return

        cursor.execute(
            """
            INSERT INTO relatorios (dia_faturamento, empresa, contrato, descricao, chegada, prazo)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                self.txt_dia_fat.get(),
                self.txt_empresa.get(),
                self.txt_contrato.get(),
                self.txt_desc.get(),
                self.txt_chegada.get(),
                self.txt_prazo.get(),
            ),
        )
        conn.commit()

        messagebox.showinfo("Sucesso", "Relatório salvo na fila de trabalho!")
        self.atualizar_tabela_pendentes()

    def atualizar_tabela_pendentes(self):
        for i in self.tabela_pendentes.get_children():
            self.tabela_pendentes.delete(i)

        # quando termina sai da fila
        cursor.execute(
            "SELECT id, dia_faturamento, empresa, contrato, descricao, prazo, feito FROM relatorios WHERE feito = '' OR enviado = ''"
        )
        linhas = cursor.fetchall()

        tarefas = []
        for l in linhas:
            t_id, d_fat, emp, cont, desc, prz, feito_status = l
            dias_fat = calcular_dias_para_faturamento(d_fat)
            data_prz = converter_para_data(prz)
            tarefas.append(
                {
                    "id": t_id,
                    "dia_fat": d_fat,
                    "empresa": emp,
                    "contrato": cont,
                    "descricao": desc,
                    "prazo": prz,
                    "feito_status": feito_status,
                    "dias_para_fat": dias_fat,
                    "data_prazo_obj": data_prz,
                }
            )

        # ordem: faturamento/prazo
        tarefas_ordenadas = sorted(
            tarefas, key=lambda x: (x["dias_para_fat"], x["data_prazo_obj"])
        )

        for t in tarefas_ordenadas:
            tag_urgencia = ""
            if t["dias_para_fat"] <= 3:
                tag_urgencia = "urgente_critico"
            elif t["dias_para_fat"] <= 7:
                tag_urgencia = "urgente_atencao"
            status_visual = (
                "⏳ Pendente" if t["feito_status"] == "" else f"✅ {t['feito_status']}"
            )

            self.tabela_pendentes.insert(
                "",
                "end",
                values=(
                    t["id"],
                    t["empresa"],
                    f"Dia {t['dia_fat']}",
                    t["prazo"],
                    t["contrato"],
                    t["descricao"],
                    status_visual,
                ),
                tags=(tag_urgencia,),
            )

    def marcar_como_feito(self):
        item_selecionado = self.tabela_pendentes.selection()
        if not item_selecionado:
            messagebox.showwarning(
                "Aviso", "Selecione uma tarefa na tabela para marcar como Feito!"
            )
            return

        id_tarefa = self.tabela_pendentes.item(item_selecionado, "values")[0]
        agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        cursor.execute(
            "UPDATE relatorios SET feito = ? WHERE id = ?", (agora, id_tarefa)
        )
        conn.commit()

        messagebox.showinfo("Sucesso", "Status atualizado: Relatório Pronto!")
        self.atualizar_tabela_pendentes()
        self.atualizar_tabela_historico()

    def marcar_como_enviado(self):
        item_selecionado = self.tabela_pendentes.selection()
        if not item_selecionado:
            messagebox.showwarning(
                "Aviso",
                "Selecione uma tarefa na tabela para marcar como Enviado!",
            )
            return

        valores = self.tabela_pendentes.item(item_selecionado, "values")
        id_tarefa = valores[0]
        status_feito_atual = valores[6]

        agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        # se não marcou como feito, vai ficar tudo com mesmo horario
        if "⏳ Pendente" in status_feito_atual:
            cursor.execute(
                "UPDATE relatorios SET feito = ?, enviado = ? WHERE id = ?",
                (agora, agora, id_tarefa),
            )
        else:
            cursor.execute(
                "UPDATE relatorios SET enviado = ? WHERE id = ?",
                (agora, id_tarefa),
            )
        conn.commit()

        messagebox.showinfo("Sucesso", "Relatório enviado com sucesso e arquivado!")
        self.atualizar_tabela_pendentes()
        self.atualizar_tabela_historico()

    # atualiza o historico
    def atualizar_tabela_historico(self, termo_busca=""):
        for i in self.tabela_historico.get_children():
            self.tabela_historico.delete(i)

        if termo_busca:
            cursor.execute(
                """
                SELECT id, empresa, contrato, descricao, feito, enviado 
                FROM relatorios 
                WHERE (feito != '' AND enviado != '') 
                AND (empresa LIKE ? OR contrato LIKE ? OR descricao LIKE ?)
                ORDER BY id DESC
                """,
                (f"%{termo_busca}%", f"%{termo_busca}%", f"%{termo_busca}%"),
            )
        else:
            cursor.execute(
                """
                SELECT id, empresa, contrato, descricao, feito, enviado 
                FROM relatorios 
                WHERE feito != '' AND enviado != '' 
                ORDER BY id DESC
                """
            )

        linhas = cursor.fetchall()
        for l in linhas:
            self.tabela_historico.insert("", "end", values=l)

    def filtrar_historico(self, event):
        termo = self.txt_pesquisa.get().strip()
        self.atualizar_tabela_historico(termo)

    def exportar_para_excel(self):
        termo_busca = self.txt_pesquisa.get().strip()

        if termo_busca:
            query = """
                SELECT id, empresa, contrato, descricao, chegada, prazo, feito, enviado
                FROM relatorios 
                WHERE (feito != '' AND enviado != '') 
                AND (empresa LIKE ? OR contrato LIKE ? OR descricao LIKE ?)
                ORDER BY id DESC
            """
            df_dados = pd.read_sql_query(
                query,
                conn,
                params=(f"%{termo_busca}%", f"%{termo_busca}%", f"%{termo_busca}%"),
            )
        else:
            query = """
                SELECT id, empresa, contrato, descricao, chegada, prazo, feito, enviado
                FROM relatorios 
                WHERE feito != '' AND enviado != '' 
                ORDER BY id DESC
            """
            df_dados = pd.read_sql_query(query, conn)

        if df_dados.empty:
            messagebox.showwarning(
                "Aviso", "Não há dados no histórico atual para exportar!"
            )
            return

        data_hoje_str = datetime.date.today().strftime("%Y-%m-%d")
        nome_arquivo = f"Relatorio_Entregas_Completo_{data_hoje_str}.xlsx"

        try:
            # Nomeia a planilha excel
            df_dados.columns = [
                "ID",
                "Empresa",
                "Contrato",
                "Descrição/Código",
                "Data de Chegada",
                "Prazo Final",
                "Horário Conclusão (Feito)",
                "Horário de Envio",
            ]
            df_dados.to_excel(nome_arquivo, index=False)
            caminho_completo = os.path.abspath(nome_arquivo)
            messagebox.showinfo(
                "Sucesso!",
                f"Relatório exportado!\n\nSalvo em:\n{caminho_completo}",
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível exportar os dados:\n{e}")


# "F5"
if __name__ == "__main__":
    app = AppRelatorios()
    app.mainloop()
    conn.close()