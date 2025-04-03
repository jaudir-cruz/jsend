# app.py
import os
import glob
import threading
import time
import pandas as pd
import ttkbootstrap as ttk
import logging  # Import necessário para usar logging.ERROR
from tkinter import filedialog, messagebox, scrolledtext, END, StringVar
from config import SESSION_ID
from logger_util import LoggerUtil
from email_client import EmailSender
from smtp_config import SMTPConfigWindow
from customization import CustomizationWindow

class App:
    def __init__(self):
        self.root = ttk.Window(themename="flatly")
        self.root.title("Envio de e-mails dinâmicos - xlsx")
        self.root.geometry("850x900")
        self.root.configure(bg="#f0f0f0")
        self.root.attributes("-fullscreen", True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.logger_util = LoggerUtil()
        self.email_sender = EmailSender(self.logger_util)
        self._build_ui()

    def _build_ui(self):
        self.frame = ttk.Frame(self.root, padding=15)
        self.frame.pack(fill='both', expand=True)

        ttk.Label(self.frame, text="Arquivo .xlsx/.xls:", font=("Arial", 12)).pack(anchor='w', pady=(0,2))
        self.entry_file = ttk.Entry(self.frame, width=60, state=ttk.DISABLED)
        self.entry_file.pack(fill='x', padx=5, pady=(0,5))
        file_button_frame = ttk.Frame(self.frame)
        file_button_frame.pack(fill='x', padx=5, pady=(0,10))
        self.btn_select_file = ttk.Button(file_button_frame, text="Selecionar", bootstyle="success", command=self.select_file, state=ttk.DISABLED)
        self.btn_select_file.pack(side='left', padx=(0,5))
        self.btn_update_vars = ttk.Button(file_button_frame, text="Atualizar Variáveis", bootstyle="info", command=self.update_variables, state=ttk.DISABLED)
        self.btn_update_vars.pack(side='left')

        ttk.Label(self.frame, text="E-mail do remetente (Send as):", font=("Arial", 12)).pack(anchor='w', pady=(0,2))
        self.entry_sender = ttk.Entry(self.frame, width=60, state=ttk.DISABLED)
        self.entry_sender.pack(fill='x', padx=5, pady=(0,5))

        ttk.Label(self.frame, text="Assunto do e-mail:", font=("Arial", 12)).pack(anchor='w', pady=(0,2))
        self.entry_subject = ttk.Entry(self.frame, width=60, state=ttk.DISABLED)
        self.entry_subject.pack(fill='x', padx=5, pady=(0,5))

        self.label_email_message = ttk.Label(self.frame, text="Corpo do e-mail: ", font=("Arial", 12))
        self.label_email_message.pack(anchor='w', pady=(0,2))

        msg_frame = ttk.Frame(self.frame)
        msg_frame.pack(fill='both', padx=5, pady=(0,5))
        msg_frame.columnconfigure(0, weight=3)
        msg_frame.columnconfigure(1, weight=1)

        self.text_body = scrolledtext.ScrolledText(msg_frame, width=70, height=10, state=ttk.DISABLED)
        self.text_body.grid(row=0, column=0, sticky="nsew", padx=(0,5), pady=5)

        button_frame = ttk.Frame(msg_frame)
        button_frame.grid(row=0, column=1, sticky="ns", pady=5)
        self.btn_attachment = ttk.Button(button_frame, text="📎 Anexar", bootstyle="info", command=self.add_attachment)
        self.btn_attachment.pack(fill='x', padx=5, pady=(0,5))
        self.btn_otimizar = ttk.Button(button_frame, text="Otimizar com IA", bootstyle="success", command=self.otimizar_email_com_ia)
        self.btn_otimizar.pack(fill='x', padx=5, pady=(0,5))
        self.btn_melhorar = ttk.Button(button_frame, text="Melhorar corpo do e-mail", bootstyle="success", command=self.open_customization_window)
        self.btn_melhorar.pack(fill='x', padx=5, pady=(0,5))

        self.attachment_label = ttk.Label(self.frame, text="Nenhum anexo selecionado.", font=("Arial", 10))
        self.attachment_label.pack(anchor='w', padx=5, pady=(0,10))

        self.var_teste = ttk.BooleanVar()
        self.chk_teste = ttk.Checkbutton(self.frame, text="Este é um envio de teste? ", variable=self.var_teste, command=self.toggle_test_fields, state=ttk.DISABLED)
        self.chk_teste.pack(anchor='w', padx=5, pady=(0,5))

        ttk.Label(self.frame, text="Destinatários dos testes (e-mail):", font=("Arial", 12)).pack(anchor='w', padx=5, pady=(0,2))
        self.entry_teste = ttk.Entry(self.frame, width=60, state=ttk.DISABLED)
        self.entry_teste.pack(fill='x', padx=5, pady=(0,5))

        test_frame = ttk.Frame(self.frame)
        test_frame.pack(fill='x', padx=5, pady=(0,10))
        ttk.Label(test_frame, text="Quantidade de envios de teste:", font=("Arial", 12)).pack(side='left', anchor='w')
        self.entry_qtd_teste = ttk.Entry(test_frame, width=10, state=ttk.DISABLED)
        self.entry_qtd_teste.pack(side='left', padx=5)

        self.var_producao = ttk.BooleanVar()
        style = ttk.Style()
        style.configure("Danger.TCheckbutton", font=("Arial", 12, "bold"))
        self.chk_producao = ttk.Checkbutton(self.frame, text="Confirma Envio para os nomes da planilha (marque se não for um teste)", variable=self.var_producao, bootstyle="danger", state=ttk.DISABLED, style="Danger.TCheckbutton")
        self.chk_producao.pack(anchor='w', padx=5, pady=(0,5))

        button_frame_global = ttk.Frame(self.frame)
        button_frame_global.pack(pady=(0,10))
        self.btn_send_email = ttk.Button(button_frame_global, text="Iniciar Envio", bootstyle="primary", command=self.enviar_emails, state=ttk.DISABLED)
        self.btn_send_email.pack(side='left', padx=5)
        ttk.Button(button_frame_global, text="Configurar SMTP", bootstyle="info", command=self.open_smtp_config).pack(side='left', padx=5)
        ttk.Button(button_frame_global, text="Sair", bootstyle="warning", command=self.on_closing).pack(side='left', padx=5)

        self.smtp_log_label_main = ttk.Label(self.frame, text="SMTP Log:", font=("Arial", 12))
        self.smtp_log_label_main.pack(anchor='w', padx=5, pady=(0,2))
        self.smtp_log_text_main = scrolledtext.ScrolledText(self.frame, width=80, height=5, state=ttk.DISABLED)
        self.smtp_log_text_main.pack(fill='x', padx=5, pady=(0,10))

        self.progress_bar = ttk.Progressbar(self.frame, mode='determinate', bootstyle="info")
        self.progress_bar.pack(fill='x', padx=5, pady=(0,10))
        self.progress_label = ttk.Label(self.frame, text="0 de 0 - 0%")
        self.progress_label.pack(fill='x', padx=5, pady=(0,10))

        ttk.Label(self.frame, text="Log de Execução:", font=("Arial", 12)).pack(anchor='w', padx=5, pady=(0,2))
        self.log_text = scrolledtext.ScrolledText(self.frame, width=80, height=10, state=ttk.DISABLED)
        self.log_text.pack(fill='both', padx=5, pady=(0,5))

        self.root.after(100, self.update_log)
        self.root.after(500, self.update_send_button_state)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return
        self.entry_file.config(state=ttk.NORMAL)
        self.entry_file.delete(0, END)
        self.entry_file.insert(0, file_path)
        self.update_send_button_state()
        self.logger_util.log_message(f"Arquivo selecionado: {file_path}", logging.DEBUG)
        try:
            df_temp = pd.read_excel(file_path, engine="openpyxl")
            columns = df_temp.columns.tolist()
            num_records = len(df_temp)
            messagebox.showinfo("Dados Importados", f"Identificado {num_records} registros na planilha.")
            self.label_email_message.config(text="Mensagem do E-mail (use os placeholders correspondentes).")
            self.btn_update_vars.config(state=ttk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo: {e}")
            self.logger_util.log_message(f"Erro ao ler arquivo: {e}", logging.ERROR)

    def update_variables(self):
        self.select_file()

    def add_attachment(self):
        path = filedialog.askopenfilename()
        if path:
            self.email_sender.attachment_path = path
            self.attachment_label.config(text=f"Anexo: {path.split('/')[-1]}")
            self.logger_util.log_message(f"Anexo selecionado: {path}", logging.DEBUG)

    def otimizar_email_com_ia(self):
        # Função de otimização (a ser implementada)
        pass

    def open_customization_window(self):
        cw = CustomizationWindow(self.root, self.text_body, self.logger_util)
        cw.open_window()

    def open_smtp_config(self):
        from smtp_config import SMTPConfigWindow
        smtp_window = SMTPConfigWindow(self.root, self.logger_util)
        smtp_window.open_window()

    def toggle_test_fields(self):
        state = ttk.NORMAL if self.var_teste.get() else ttk.DISABLED
        self.entry_teste.config(state=state)
        self.entry_qtd_teste.config(state=state)
        self.update_send_button_state()
        self.logger_util.log_message("Toggle dos campos de teste alterado.", logging.DEBUG)

    def update_send_button_state(self):
        if self.entry_file.get().strip() and self.entry_sender.get().strip() and self.entry_subject.get().strip() and self.text_body.get("1.0", END).strip():
            self.btn_send_email.config(state=ttk.NORMAL)
        else:
            self.btn_send_email.config(state=ttk.DISABLED)
        self.root.after(500, self.update_send_button_state)

    def update_log(self):
        log_q = self.logger_util.get_log_queue()
        while not log_q.empty():
            msg = log_q.get_nowait()
            self.log_text.config(state=ttk.NORMAL)
            self.log_text.insert(END, msg + "\n")
            self.log_text.see(END)
            self.log_text.config(state=ttk.DISABLED)
            self.smtp_log_text_main.config(state=ttk.NORMAL)
            self.smtp_log_text_main.insert(END, msg + "\n")
            self.smtp_log_text_main.see(END)
            self.smtp_log_text_main.config(state=ttk.DISABLED)
        self.root.after(100, self.update_log)

    def enviar_emails(self):
        file_path = self.entry_file.get()
        email_sender = self.entry_sender.get()
        email_subject = self.entry_subject.get()
        email_body = self.text_body.get("1.0", END).strip()
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            self.logger_util.log_message(f"Arquivo carregado: {file_path}", logging.DEBUG)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo: {e}")
            self.logger_util.log_message(f"Erro ao carregar arquivo: {e}", logging.ERROR)
            return

        recipients = []
        for _, row in df.iterrows():
            email_receiver = str(row.get("Email", "")).strip()
            if email_receiver:
                recipients.append((email_receiver, row))
        self.logger_util.log_message(f"📨 Iniciando envio de {len(recipients)} e-mails...", logging.DEBUG)
        self.email_sender.total_emails = len(recipients)
        self.email_sender.emails_processed = 0
        self.email_sender.start_time = time.time()

        def sending_process():
            final_results = self.email_sender.process_recipients(recipients, email_body, email_sender, email_subject)
            success_count = len([res for res in final_results.values() if res[2].startswith("Enviado")])
            failure_count = len([res for res in final_results.values() if not res[2].startswith("Enviado")])
            elapsed_time = time.time() - self.email_sender.start_time
            if failure_count:
                messagebox.showinfo("Concluído", f"Envio finalizado!\nSucesso: {success_count} e Falhas: {failure_count}")
                self.logger_util.log_message("Envio finalizado com falhas.", logging.DEBUG)
            else:
                messagebox.showinfo("Concluído", "Envio finalizado! Todos os e-mails foram enviados com sucesso.")
                self.logger_util.log_message("Todos os e-mails enviados com sucesso.", logging.DEBUG)
                send_notif = messagebox.askyesno("Notificação", "Deseja enviar um relatório de envio por e-mail para ti@crn2.org.br?")
                if send_notif:
                    report_html = self.email_sender.build_report_html(len(recipients), elapsed_time)
                    if self.email_sender.send_notification(report_html):
                        messagebox.showinfo("Notificação", "Relatório enviado com sucesso!")
                    else:
                        messagebox.showerror("Notificação", "Falha ao enviar relatório.")
            historico_filename = f"historico_envio_{SESSION_ID}.xlsx"
            pd.DataFrame(list(final_results.values()), columns=["Email", "Dados", "Resultado"]).to_excel(historico_filename, index=False, engine="openpyxl")
            self.logger_util.log_message(f"Histórico de envio salvo em {historico_filename}", logging.DEBUG)

        threading.Thread(target=sending_process).start()

    def on_closing(self):
        resposta = messagebox.askyesnocancel("Sair", "Deseja salvar os logs e arquivos gerados nesta sessão?")
        if resposta is None:
            return
        elif resposta is False:
            # Encerra o sistema de logging para liberar os arquivos
            logging.shutdown()
            patterns = [f"log_execucao_{SESSION_ID}.txt", f"preview_email_{SESSION_ID}.html", f"historico_envio_{SESSION_ID}.xlsx"]
            for pat in patterns:
                for f in glob.glob(os.path.join(os.getcwd(), pat)):
                    try:
                        os.remove(f)
                    except Exception as e:
                        self.logger_util.log_message(f"Erro ao excluir {f}: {e}", logging.ERROR)
        self.root.destroy()

if __name__ == "__main__":
    app = App()
    app.root.mainloop()
