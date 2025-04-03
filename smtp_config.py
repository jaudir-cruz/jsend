# smtp_config.py
import ssl
import smtplib
from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import SUCCESS  # Importação necessária
from config import SMTP_SERVER, SMTP_PORT, EMAIL_AUTH, EMAIL_PASSWORD
from logger_util import LoggerUtil

class SMTPConfigWindow:
    def __init__(self, root, logger_util):
        self.root = root
        self.logger_util = logger_util

    def open_window(self):
        self.config_window = ttk.Toplevel(self.root)
        self.config_window.title("Configuração SMTP")
        self.config_window.geometry("410x500")
        self.config_window.grab_set()
        self._build_ui()

    def _build_ui(self):
        frame_smtp = ttk.Frame(self.config_window, padding=10)
        frame_smtp.pack(fill="both", expand=True)
        style = ttk.Style()
        style.configure("Header.TFrame", background="#4a4a4a")
        style.configure("Header.TLabel", background="#4a4a4a", foreground="white", font=("Arial", 14, "bold"))
        header_frame = ttk.Frame(frame_smtp, style="Header.TFrame")
        header_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0,10))
        frame_smtp.columnconfigure(0, weight=1)
        ttk.Label(header_frame, text="Configurar Credenciais SMTP", style="Header.TLabel").pack(fill="both", expand=True, padx=20, pady=15)
        ttk.Label(frame_smtp, text="Servidor SMTP:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=5)
        self.entry_smtp_server = ttk.Entry(frame_smtp, width=40)
        self.entry_smtp_server.grid(row=1, column=1, sticky="w", pady=5)
        self.entry_smtp_server.insert(0, SMTP_SERVER)
        ttk.Label(frame_smtp, text="Porta SMTP:", font=("Arial", 12)).grid(row=2, column=0, sticky="w", pady=5)
        self.entry_smtp_port = ttk.Entry(frame_smtp, width=40)
        self.entry_smtp_port.grid(row=2, column=1, sticky="w", pady=5)
        self.entry_smtp_port.insert(0, SMTP_PORT)
        ttk.Label(frame_smtp, text="E-mail/usuário:", font=("Arial", 12)).grid(row=3, column=0, sticky="w", pady=5)
        self.entry_smtp_email = ttk.Entry(frame_smtp, width=40)
        self.entry_smtp_email.grid(row=3, column=1, sticky="w", pady=5)
        self.entry_smtp_email.insert(0, EMAIL_AUTH)
        ttk.Label(frame_smtp, text="Senha:", font=("Arial", 12)).grid(row=4, column=0, sticky="w", pady=5)
        self.entry_smtp_pass = ttk.Entry(frame_smtp, width=40, show="*")
        self.entry_smtp_pass.grid(row=4, column=1, sticky="w", pady=5)
        self.entry_smtp_pass.insert(0, EMAIL_PASSWORD)
        ttk.Label(frame_smtp, text="SMTP Log:", font=("Arial", 10, "bold")).grid(row=5, column=0, sticky="w", pady=5)
        self.smtp_log_text = ttk.ScrolledText(frame_smtp, width=60, height=8, state=ttk.DISABLED)
        self.smtp_log_text.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Button(frame_smtp, text="Testar e Aplicar", bootstyle=SUCCESS, command=self.test_and_apply_smtp).grid(row=7, column=0, columnspan=2, pady=10)
        ttk.Button(frame_smtp, text="Sair", bootstyle="warning", command=self.config_window.destroy).grid(row=8, column=0, columnspan=2, pady=5)

    def log_smtp_message(self, msg):
        self.smtp_log_text.config(state=ttk.NORMAL)
        self.smtp_log_text.insert("end", msg + "\n")
        self.smtp_log_text.see("end")
        self.smtp_log_text.config(state=ttk.DISABLED)

    def test_and_apply_smtp(self):
        new_server = self.entry_smtp_server.get().strip()
        new_port = self.entry_smtp_port.get().strip()
        new_email = self.entry_smtp_email.get().strip()
        new_pass = self.entry_smtp_pass.get().strip()
        if not new_server or not new_port or not new_email or not new_pass:
            messagebox.showerror("Erro", "Preencha todos os campos de configuração SMTP.")
            return
        try:
            self.log_smtp_message("Tentando conexão SMTP...")
            port_int = int(new_port)
            context = ssl.create_default_context()
            with smtplib.SMTP(new_server, port_int) as server:
                server.starttls(context=context)
                server.login(new_email, new_pass)
                self.log_smtp_message("Conexão estabelecida com sucesso!")
            messagebox.showinfo("Sucesso", "Credenciais SMTP autenticadas com sucesso!")
            self.logger_util.log_message("Credenciais SMTP autenticadas.", logging.DEBUG)
            self.config_window.destroy()
        except Exception as e:
            self.log_smtp_message(f"Falha na conexão: {e}")
            messagebox.showerror("Erro", f"Falha ao autenticar: {e}")
