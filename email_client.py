# email_client.py
import smtplib
import ssl
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import concurrent.futures
import threading
import time
import pandas as pd
from config import SMTP_SERVER, SMTP_PORT, EMAIL_AUTH, EMAIL_PASSWORD, HF_TOKEN, SESSION_ID, NOTIFICATION_RECIPIENT
from logger_util import LoggerUtil

class EmailSender:
    def __init__(self, logger_util):
        self.logger_util = logger_util
        self.emails_processed = 0
        self.total_emails = 0
        self.counter_lock = threading.Lock()
        self.start_time = None
        self.attachment_path = None  # Caminho do anexo
        self.reenvios_count = 0

    def send_email(self, email_receiver, dados_linha, email_body, email_sender, email_subject):
        inicio = time.time()
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls(context=context)
                server.login(EMAIL_AUTH, EMAIL_PASSWORD)
                email_content = email_body
                for key, value in dados_linha.items():
                    placeholder = "{" + str(key).upper() + "}"
                    email_content = email_content.replace(placeholder, str(value).strip() if pd.notna(value) else "")
                msg = MIMEMultipart()
                msg["From"] = email_sender
                msg["To"] = email_receiver
                msg["Subject"] = email_subject
                msg.attach(MIMEText(email_content, "html"))
                if self.attachment_path and os.path.exists(self.attachment_path):
                    with open(self.attachment_path, "rb") as attachment_file:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment_file.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(self.attachment_path)}")
                    msg.attach(part)
                server.sendmail(email_sender, email_receiver, msg.as_string())
            duracao = time.time() - inicio
            self.logger_util.log_message(f"✅ Enviado para {email_receiver} ({duracao:.2f}s)", logging.DEBUG)
            return (email_receiver, dados_linha, f"Enviado ({duracao:.2f}s)")
        except Exception as e:
            self.logger_util.log_message(f"❌ Erro ao enviar para {email_receiver}: {e}", logging.ERROR)
            return (email_receiver, dados_linha, f"Erro: {e}")
        finally:
            with self.counter_lock:
                self.emails_processed += 1

    def update_progress_info(self, progress_bar, progress_label, callback=None):
        with self.counter_lock:
            processed = self.emails_processed
        perc = (processed / self.total_emails * 100) if self.total_emails > 0 else 0
        elapsed = (time.time() - self.start_time) if self.start_time else 0
        if processed > 0:
            avg_time = elapsed / processed
            estimated_total = avg_time * self.total_emails
            remaining = estimated_total - elapsed
        else:
            remaining = 0
        progress_bar.config(value=perc)
        progress_label.config(text=f"{processed} de {self.total_emails} - {perc:.1f}%\nTempo decorrido: {elapsed:.1f}s | Estimado restante: {remaining:.1f}s")
        if processed < self.total_emails:
            if callback:
                progress_bar.after(1000, lambda: callback(progress_bar, progress_label, callback))

    def process_recipients(self, recipients, email_body, email_sender, email_subject):
        final_results_dict = {}
        current_recipients = recipients
        while True:
            results = self._process_emails_for_recipients(current_recipients, email_body, email_sender, email_subject)
            for res in results:
                final_results_dict[res[0]] = res
            failures = [res for res in final_results_dict.values() if not res[2].startswith("Enviado")]
            if failures:
                self.reenvios_count += len(failures)
                # Aqui você pode implementar uma lógica para perguntar ao usuário (por meio de interface) se deseja reenviar
                # Para fins de exemplo, vamos assumir que não se deseja reenviar após a primeira tentativa.
                break
            else:
                break
        return final_results_dict

    def _process_emails_for_recipients(self, recipient_list, email_body, email_sender, email_subject):
        results = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(self.send_email, email, dados, email_body, email_sender, email_subject)
                       for (email, dados) in recipient_list]
            concurrent.futures.wait(futures)
            for future in futures:
                results.append(future.result())
        return results

    def send_notification(self, report_html):
        # Monta e envia o e-mail de notificação
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls(context=context)
                server.login(EMAIL_AUTH, EMAIL_PASSWORD)
                msg = MIMEMultipart()
                msg["From"] = EMAIL_AUTH
                msg["To"] = NOTIFICATION_RECIPIENT
                msg["Subject"] = "Relatório de Envio - JSend Notificação"
                msg.attach(MIMEText(report_html, "html"))
                log_filepath = os.path.join(os.getcwd(), f"log_execucao_{SESSION_ID}.txt")
                if os.path.exists(log_filepath):
                    with open(log_filepath, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(log_filepath)}")
                    msg.attach(part)
                server.sendmail(EMAIL_AUTH, NOTIFICATION_RECIPIENT, msg.as_string())
                self.logger_util.log_message(f"Relatório de notificação enviado para {NOTIFICATION_RECIPIENT}.", logging.DEBUG)
                return True
        except Exception as e:
            self.logger_util.log_message(f"❌ Erro ao enviar relatório de notificação: {e}", logging.ERROR)
            return False

    def build_report_html(self, total_emails, elapsed_time):
        # Personalizado com tema azul e amarelo, nome do sistema JSend
        report_html = f"""\
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>JSend - Notificação</title>
</head>
<body style="font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0;">
  <div style="max-width:600px; margin:20px auto; background-color: #fff; border: 2px solid #0000FF; padding:20px; border-radius:8px;">
    <h2 style="color: #0000FF;">JSend - Notificação</h2>
    <p>Relatório Executivo de Envio:</p>
    <ul>
      <li>Total de e-mails enviados: {total_emails}</li>
      <li>Tempo decorrido: {elapsed_time:.1f} segundos</li>
      <li>Quantidade de reenvios: {self.reenvios_count}</li>
    </ul>
  </div>
</body>
</html>
"""
        return report_html
