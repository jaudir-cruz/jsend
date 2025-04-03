# customization.py
import base64
from tkinter import filedialog, END, messagebox
import ttkbootstrap as ttk
from email_client import EmailSender
from logger_util import LoggerUtil

class CustomizationWindow:
    def __init__(self, root, text_body_widget, logger_util):
        self.root = root
        self.text_body = text_body_widget
        self.logger_util = logger_util
        self.header_image_data = None
        self.header_image_mime = None

    def open_window(self):
        self.custom_win = ttk.Toplevel(self.root)
        self.custom_win.title("Personalizar Corpo do E-mail")
        self.custom_win.geometry("850x700")
        self.custom_win.grab_set()
        self._build_ui()

    def _build_ui(self):
        # Configura o grid da janela
        self.custom_win.columnconfigure(0, weight=1)
        self.custom_win.rowconfigure(1, weight=1)

        # Painel de Configurações
        config_frame = ttk.LabelFrame(self.custom_win, text="Configurações", padding=10)
        config_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
        ttk.Label(config_frame, text="Cor do Tema:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.theme_color_var = ttk.StringVar(value="verde")
        cmb_theme = ttk.Combobox(config_frame, textvariable=self.theme_color_var, values=["verde", "azul", "vermelho", "preto", "laranja"], state="readonly", width=15)
        cmb_theme.grid(row=0, column=1, sticky="w", pady=5, padx=5)
        ttk.Label(config_frame, text="Tipo de Cor:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.color_type_var = ttk.StringVar(value="sólida")
        cmb_type = ttk.Combobox(config_frame, textvariable=self.color_type_var, values=["sólida", "fosca"], state="readonly", width=15)
        cmb_type.grid(row=1, column=1, sticky="w", pady=5, padx=5)
        btn_imagem = ttk.Button(config_frame, text="Importar Imagem de Cabeçalho", command=self.importar_imagem)
        btn_imagem.grid(row=2, column=0, columnspan=2, pady=5, padx=5)
        self.lbl_imagem = ttk.Label(config_frame, text="Nenhuma imagem selecionada.")
        self.lbl_imagem.grid(row=3, column=0, columnspan=2, pady=5, padx=5)

        # Editor e Prévia
        editor_frame = ttk.Frame(self.custom_win, padding=10)
        editor_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        ttk.Label(editor_frame, text="Conteúdo do E-mail (texto base):").pack(anchor="w", pady=(0,5))
        self.txt_body_custom = ttk.ScrolledText(editor_frame, width=90, height=8)
        self.txt_body_custom.pack(fill="both", pady=(0,10))
        self.txt_body_custom.insert(END, self.text_body.get("1.0", END).strip())
        ttk.Label(editor_frame, text="Prévia do HTML gerado:").pack(anchor="w", pady=(10,5))
        self.txt_preview = ttk.ScrolledText(editor_frame, width=90, height=15, state=ttk.DISABLED)
        self.txt_preview.pack(fill="both", pady=(0,10))

        # Botões de Ação
        btn_frame = ttk.Frame(self.custom_win)
        btn_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(5,10))
        btn_gerar = ttk.Button(btn_frame, text="Gerar Corpo do E-mail", bootstyle="primary", command=self.gerar_corpo)
        btn_gerar.pack(side="left", padx=5)
        btn_aplicar = ttk.Button(btn_frame, text="Aplicar Modificações", bootstyle="success", command=self.aplicar_modificacoes)
        btn_aplicar.pack(side="left", padx=5)

    def importar_imagem(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif")])
        if path:
            if os.path.getsize(path) > 500 * 1024:
                messagebox.showerror("Erro", "A imagem excede o tamanho máximo de 500KB.")
                return
            with open(path, "rb") as img_file:
                data = img_file.read()
                self.header_image_data = base64.b64encode(data).decode("utf-8")
            ext = path.split(".")[-1].lower()
            if ext == "png":
                self.header_image_mime = "image/png"
            elif ext in ["jpg", "jpeg"]:
                self.header_image_mime = "image/jpeg"
            elif ext == "gif":
                self.header_image_mime = "image/gif"
            else:
                self.header_image_mime = "application/octet-stream"
            self.lbl_imagem.config(text=f"Imagem selecionada: {path.split('/')[-1]}")

    def gerar_corpo(self):
        html = self.generate_custom_html(
            self.txt_body_custom.get("1.0", END).strip(),
            self.theme_color_var.get(),
            self.color_type_var.get(),
            self.header_image_data,
            self.header_image_mime
        )
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", END)
        self.txt_preview.insert(END, html)
        self.txt_preview.config(state="disabled")

    def aplicar_modificacoes(self):
        new_html = self.txt_preview.get("1.0", END).strip()
        self.text_body.config(state="normal")
        self.text_body.delete("1.0", END)
        self.text_body.insert(END, new_html)
        self.text_body.config(state="normal")
        self.logger_util.log_message("Corpo do e-mail atualizado com as personalizações.", logging.DEBUG)
        self.custom_win.destroy()

    def generate_custom_html(self, body_text, theme_color, color_type, header_image_data, header_image_mime):
        # Reaproveitando o HTML do script original
        theme_colors = {
            "verde": "#008000",
            "azul": "#0000FF",
            "vermelho": "#FF0000",
            "preto": "#000000",
            "laranja": "#FFA500"
        }
        color_code = theme_colors.get(theme_color, "#008000")
        if color_type.lower() == "fosca":
            container_style = f"border-top: 5px solid {color_code}; opacity: 0.8;"
        else:
            container_style = f"border-top: 5px solid {color_code};"
        header_img_html = ""
        if header_image_data and header_image_mime:
            header_img_html = f'<div style="text-align:center; margin-bottom:15px;"><img src="data:{header_image_mime};base64,{header_image_data}" style="max-width:100%; max-height:150px; border-radius:8px;"></div>'
        html_template = f"""\
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Aviso importante!</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            margin: 0;
            padding: 0;
        }}
        .container {{
            max-width: 600px;
            margin: 20px auto;
            background-color: #ffffff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            {container_style}
        }}
        .header {{
            text-align: center;
            color: {color_code};
            margin-bottom: 15px;
        }}
        .content {{
            color: #333;
            line-height: 1.6;
            margin-top: 20px;
        }}
        .button {{
            display: inline-block;
            padding: 10px 20px;
            margin-top: 20px;
            color: #ffffff !important;
            background-color: {color_code};
            text-decoration: none;
            border-radius: 5px;
            font-size: 16px;
            font-weight: bold;
        }}
        .button:hover {{
            opacity: 0.9;
        }}
        .footer {{
            margin-top: 20px;
            font-size: 12px;
            text-align: center;
            color: #666;
        }}
        a {{
            text-decoration: none;
            color: #ffffff;
        }}
    </style>
</head>
<body>
    <div class="container">
        {header_img_html}
        <h2 class="header">Aviso importante!</h2>
        <div class="content">
            {body_text}
        </div>
    </div>
</body>
</html>
"""
        return html_template
