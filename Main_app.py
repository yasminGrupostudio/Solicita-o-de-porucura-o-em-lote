import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import win32com.client as win32
import pythoncom
import pandas as pd
import re
import threading
import time
import os
import glob

# Mapeamento dinâmico de CNPJ por Razão Social (Coluna C)
MAPA_CNPJ = {
    "STUDIO OPERACIONAL": "23.448.109/0001-91",
    "STUDIO AUDIT": "23.448.109/0001-91",
    "STUDIO VAREJO": "44.189.727/0001-34",
    "OPERACIONAL 1": "62.700.834/0001-67",
    "STUDIO STORE": "48.552.493/0001-07",
    "ALIANÇA": "12.340.921/0001-82",
    "ALIANCA": "12.340.921/0001-82"
}


CRIAR_APENAS_RASCUNHO = True

# Configurações de Controle de Envio Direto
LIMITE_LOTE_EMAILS = 10        # Reduzido para 10 para evitar gatilhos de antispam
PAUSA_ENTRE_EMAILS = 30        # Pausa de 30 segundos
PAUSA_LOTE_SEGUNDOS = 900      # Pausa longa de 15 minutos se usar envio direto

def extrair_link_exportacao(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if match:
        id_planilha = match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{id_planilha}/export?format=xlsx"
    else:
        raise ValueError("URL do Google Sheets inválida ou incorreta.")

def limpar_emails(texto_email):
    if not texto_email or str(texto_email).lower() in ["nan", "none"]:
        return ""
    
    padrao_email = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    emails_encontrados = re.findall(padrao_email, str(texto_email).lower())
    emails_unicos = list(dict.fromkeys(emails_encontrados))
    
    return ";".join(emails_unicos)

def obter_cnpj(razao_social):
    chave = str(razao_social).strip().upper()
    return MAPA_CNPJ.get(chave, "23.448.109/0001-91")

def montar_corpo_mensagem(razao_social, cnpj, possui_imagem=False):
    html_imagem = '<p><img src="cid:imagem_corpo" style="max-width: 500px; height: auto;"></p><br>' if possui_imagem else ''
    
    return f"""
    <p>Prezados.</p>
    {html_imagem}
    <p>Informamos que a procuração eletrônica da empresa expirou.</p>
    <p>Solicitamos, por gentileza, a renovação do documento para que possamos dar continuidade.</p>
    <br>
    <p>Para auxiliá-los no procedimento, seguem abaixo os dados cadastrais da empresa e, em anexo, o manual de orientações:</p>
    <p><b>CNPJ:</b> {cnpj}<br>
    <b>Razão Social:</b> {razao_social}</p>
    <br>
    <p>Permanecemos à disposição para esclarecer qualquer dúvida que possa surgir durante o processo.</p>
    <p>Atenciosamente,</p>
    """

def localizar_arquivo_manual(pasta_trabalho):
    padrao_pdf = os.path.join(pasta_trabalho, "*.pdf")
    arquivos_pdf = glob.glob(padrao_pdf)
    
    if not arquivos_pdf:
        return None

    for pdf in arquivos_pdf:
        nome_arquivo = os.path.basename(pdf).lower()
        if "manual" in nome_arquivo or "procura" in nome_arquivo:
            return pdf
            
    return arquivos_pdf[0]

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador Procuração Eletrônica - Grupo Studio")
        self.root.geometry("680x520")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.theme_use("clam")

        lbl_instrucao = tk.Label(root, text="Cole abaixo a URL da planilha do Google Sheets:", font=("Segoe UI", 10, "bold"))
        lbl_instrucao.pack(anchor="w", padx=20, pady=(20, 5))

        self.txt_url = ttk.Entry(root, width=80, font=("Segoe UI", 10))
        self.txt_url.pack(padx=20, pady=5, fill="x")

        self.btn_iniciar = ttk.Button(root, text="🚀 Iniciar Processamento", command=self.iniciar_processo_thread)
        self.btn_iniciar.pack(padx=20, pady=15, fill="x")

        lbl_log = tk.Label(root, text="Status e Registros de Envio:", font=("Segoe UI", 10, "bold"))
        lbl_log.pack(anchor="w", padx=20, pady=(10, 5))

        self.log_area = scrolledtext.ScrolledText(root, width=80, height=15, font=("Consolas", 9))
        self.log_area.pack(padx=20, pady=(0, 20), fill="both", expand=True)

    def log(self, mensagem):
        self.root.after(0, self._atualizar_log, mensagem)

    def _atualizar_log(self, mensagem):
        self.log_area.insert(tk.END, mensagem + "\n")
        self.log_area.see(tk.END)

    def iniciar_processo_thread(self):
        url = self.txt_url.get().strip()
        if not url:
            messagebox.showwarning("Aviso", "Por favor, insira o link do Google Sheets.")
            return

        self.btn_iniciar.config(state="disabled")
        self.log_area.delete(1.0, tk.END)
        
        threading.Thread(target=self.processar_envio, args=(url,), daemon=True).start()

    def processar_envio(self, url):
        pythoncom.CoInitialize()
        try:
            self.log("Carregando dados da planilha...")
            link_direto = extrair_link_exportacao(url)

            excel_file = pd.ExcelFile(link_direto)
            
            if "banco de dados" not in excel_file.sheet_names:
                self.log("❌ ERRO: Aba 'banco de dados' não encontrada na planilha!")
                return

            df_bd = pd.read_excel(excel_file, sheet_name="banco de dados")
            
            col_empresa_bd = None
            col_email_cliente = None
            col_email_interno = None

            for col in df_bd.columns:
                nome_limpo = str(col).strip().lower()
                if 'empresa' in nome_limpo or 'razão' in nome_limpo or 'razao' in nome_limpo:
                    col_empresa_bd = col
                elif 'e-mails cliente' in nome_limpo or 'email cliente' in nome_limpo:
                    col_email_cliente = col
                elif 'e-mails internos' in nome_limpo or 'email interno' in nome_limpo:
                    col_email_interno = col

            if not col_empresa_bd:
                col_empresa_bd = df_bd.columns[1] if len(df_bd.columns) > 1 else df_bd.columns[0]
            if not col_email_cliente:
                col_email_cliente = [c for c in df_bd.columns if 'cliente' in str(c).lower() or 'mail' in str(c).lower()][0]
            if not col_email_interno:
                col_email_interno = [c for c in df_bd.columns if 'interno' in str(c).lower() or 'mail' in str(c).lower()][1]

            pasta_script = os.path.dirname(os.path.abspath(__file__))

            caminho_anexo = localizar_arquivo_manual(pasta_script)
            possui_anexo = True if caminho_anexo and os.path.exists(caminho_anexo) else False

            caminho_imagem = os.path.join(pasta_script, "imagem.png")
            possui_imagem = os.path.exists(caminho_imagem)

            nome_aba_envio = excel_file.sheet_names[0]
            df_envio = pd.read_excel(excel_file, sheet_name=nome_aba_envio)

            outlook = win32.Dispatch('outlook.application')
            df_bd_empresas = df_bd[col_empresa_bd].astype(str).str.strip().str.upper()

            contador_sucesso = 0

            for idx, linha in df_envio.iterrows():
                job_raw = str(linha.get('job', linha.iloc[0])).strip()
                empresa = str(linha.get('Empresa', linha.iloc[1])).strip()
                razao_social = str(linha.get('Razão Social', linha.iloc[2])).strip()

                if not empresa or empresa.lower() in ['nan', 'none', '']:
                    continue

                cnpj_final = obter_cnpj(razao_social)
                empresa_busca = empresa.strip().upper()
                correspondencia = df_bd[df_bd_empresas == empresa_busca]

                if not correspondencia.empty:
                    raw_to = correspondencia.iloc[0][col_email_cliente]
                    raw_cc = correspondencia.iloc[0][col_email_interno]

                    email_to = limpar_emails(raw_to)
                    email_cc = limpar_emails(raw_cc)

                    if not email_to:
                        self.log(f"⚠️ AVISO: Empresa '{empresa}' sem e-mail cadastrado.")
                        continue

                    email_enviado_ou_salvo = False
                    tentativas = 0

                    while not email_enviado_ou_salvo and tentativas < 3:
                        try:
                            email = outlook.CreateItem(0)
                            email.To = email_to
                            if email_cc:
                                email.CC = email_cc

                            email.Subject = f"Renovação de Procuração Eletrônica - {empresa} - {job_raw}"

                            if possui_imagem:
                                attachment = email.Attachments.Add(caminho_imagem)
                                attachment.PropertyAccessor.SetProperty(
                                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E", 
                                    "imagem_corpo"
                                )

                            if possui_anexo:
                                email.Attachments.Add(caminho_anexo)

                            corpo_html = montar_corpo_mensagem(
                                razao_social=razao_social, 
                                cnpj=cnpj_final, 
                                possui_imagem=possui_imagem
                            )
                            
                            email.HTMLBody = corpo_html

                            if CRIAR_APENAS_RASCUNHO:
                                email.Save()  # Salva na pasta Rascunhos do Outlook
                                contador_sucesso += 1
                                self.log(f"📁 RASCUNHO CRIADO ({contador_sucesso}) -> {empresa}")
                                email_enviado_ou_salvo = True
                            else:
                                email.Send()
                                contador_sucesso += 1
                                self.log(f"✅ SUCESSO ({contador_sucesso}) -> Enviado para {empresa}")
                                email_enviado_ou_salvo = True

                                # Controle de cadência do envio direto
                                if contador_sucesso % LIMITE_LOTE_EMAILS == 0:
                                    self.log(f"⏸️ Cota temporária atingida. Pausando 15 minutos para desbloquear o servidor...")
                                    time.sleep(PAUSA_LOTE_SEGUNDOS)
                                else:
                                    time.sleep(PAUSA_ENTRE_EMAILS)

                        except Exception as err:
                            tentativas += 1
                            self.log(f"⚠️ Falha no envio para {empresa} (Tentativa {tentativas}): {err}")
                            if "554" in str(err) or "limit" in str(err).lower():
                                self.log("⏳ Bloqueio de limite detectado no servidor! Pausando por 15 minutos antes de tentar novamente...")
                                time.sleep(900)
                            else:
                                time.sleep(5)

                else:
                    self.log(f"⚠️ AVISO: Empresa '{empresa}' não foi localizada no Banco de Dados.")

            self.log("\n✅ Processo concluído com sucesso!")
            msg_final = "Rascunhos gerados no Outlook com sucesso!" if CRIAR_APENAS_RASCUNHO else "Envio dos e-mails concluído!"
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", msg_final))

        except Exception as e:
            self.log(f"\n❌ ERRO CRÍTICO: {e}")
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu uma falha: {e}"))
            
        finally:
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.btn_iniciar.config(state="normal"))

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()