import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select 
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException 
import time
import re
import os 
import base64 
from dotenv import load_dotenv 

# --- 1. CONFIGURAÇÕES INICIAIS ---

load_dotenv()

URL_PORTAL_LF = "https://erp.bem.inf.br/ords/f?p=8000000:9999:2963151475408:::::"
URL_PORTAL_WEG_JVE = "https://www.me.com.br/do/Login.mvc/LoginNew"
URL_PORTAL_WEG_DI = "https://www.me.com.br/do/Login.mvc/LoginNew"

CREDENCIAIS = {
    "LF": {
        "user": os.environ.get("LF_USER"), 
        "pass": os.environ.get("LF_PASS")
    },
    "WEG_JVE": {
        "user": os.environ.get("WEG_JVE_USER"), 
        "pass": os.environ.get("WEG_JVE_PASS")
    },
    "WEG_DI": {
        "user": os.environ.get("WEG_DI_USER"), 
        "pass": os.environ.get("WEG_DI_PASS")
    }
}

if not all([
    CREDENCIAIS["LF"]["user"], CREDENCIAIS["LF"]["pass"],
    CREDENCIAIS["WEG_JVE"]["user"], CREDENCIAIS["WEG_JVE"]["pass"],
    CREDENCIAIS["WEG_DI"]["user"], CREDENCIAIS["WEG_DI"]["pass"]
]):
    # Se não tiver .env, tenta usar valores hardcoded ou avisa
    # messagebox.showerror("Erro de Credenciais", "Verifique o arquivo .env")
    pass 

MAPA_APLICACAO = {
    "Industrialização": "INDUSTRIALIZAÇÃO", 
    "Consumo": "USO E CONSUMO/IMOBILIZADO" 
}

# XPATHs LF
XPATHS_LF = {
    "login_user": '//*[@id="P9999_USERNAME"]',
    "login_pass": '//*[@id="P9999_PASSWORD"]',
    "login_btn": "//button[contains(., 'Acessar')]", 
    "nav_menu_btn": "/html/body/form/header/div[1]/div[1]/button",
    "nav_pedido_link": "/html/body/form/div[1]/div[1]/div/ul/li[9]/div[2]/a",
    "novo_pedido_btn": "/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[1]/div/div/div[1]/div[3]/button[1]",
    "novo_pedido_iframe": "//iframe[@title='Pedido']", 
    "novo_pedido_iframe_check": "P7_AD_ORG_ID", 
    "campo_filial": "P7_AD_ORG_ID",
    "campo_finalidade": "P7_BEM_VEN_FINALIDADE_ID",
    "campo_tipo_nf": "P7_BEG_TIPO_NF_ID",
    "campo_ordem_compra": "P7_NRO_ORDEM_COMPRA",
    "campo_tipo_frete": "P7_BEM_TIPO_FRETE_ID",
    "campo_obs_nf": "P7_MSG_OBS",
    "campo_obs_pedido": "P7_OBS",
    "btn_salvar_cabecalho": "//button[span[normalize-space(.)='Incluir'] and not(ancestor::div[contains(@class,'ui-dialog-content')])]",
    "radio_ped_complementar_sim": "P7_IND_PED_COMPLEMENTAR_0",
    "checkbox_industrializacao": "//input[@data-fis-ttd-id='18' and @type='checkbox']",
    "btn_adicionar_item": "//button[@title='Novo Item Pedido']", 
    "item_popup_iframe": "//iframe[@title='Itens do Pedido']",
    "item_popup_iframe_check": "P8_BEG_PRD_SKU_ID_lov_btn", 
    "item_popup_sku_btn": "P8_BEG_PRD_SKU_ID_lov_btn",
    "item_sku_search_result": "//div[contains(@class, 'a-PopupLOV-results')]//table/tbody/tr[1]/td[1]",
    "item_popup_ordem_compra": "P8_NRO_ORDEM_COMPRA", 
    "item_popup_item_oc": "P8_NRO_ITEM_OC", 
    "item_popup_qtd": "P8_QUANTIDADE",
    "item_popup_preco": "P8_VLR_UNIT_SUGERIDO",
    "item_popup_btn_salvar": "btn-incluir",
    
    # --- CORREÇÃO DO BOTÃO NOVO (ANEXAR) ---
    # Removido o ID fixo e trocado por busca de texto 'Novo'
    "btn_anexar": "//button[span[normalize-space(.)='Novo']]",
    # ---------------------------------------

    "anexo_iframe": "//iframe[@title='Anexos Pedido']",
    "anexo_iframe_check": "P103_DESCRICAO", 
    "anexo_descricao_input": "P103_DESCRICAO", 
    "anexo_upload_input": "//input[@type='file']", 
    "anexo_btn_salvar": "//button[span[normalize-space(.)='Incluir']]"
}

# XPATHS WEG
XPATHS_WEG = {
    "login_user": '//*[@id="LoginName"]',
    "login_pass": '//*[@id="RAWSenha"]',
    "login_btn": '//*[@id="SubmitAuth"]',
    "campo_busca_pedido": "//input[@placeholder='Pesquisar']", 
    "popup_btn_fechar_X": "/html/body/div[4]/div[1]/div/div/header/div/div[1]/div/div/button/i",
    "popup_iframe": "/html/body/div[4]/div[1]/div/div/div/iframe", 
    "popup_iframe_check": "/html/body/main/form[1]/div/section[1]/div/div/div[1]/div[1]/p[1]", 
    "popup_cnpj": "/html/body/main/form[1]/div/section[1]/div/div/div[1]/div[1]/p[1]",
    "popup_cond_pagto": "/html/body/main/form[1]/div/section[2]/div/div/div[5]/div/div",
    "popup_cif_fob": "/html/body/main/form[1]/div/section[2]/div/div/div[4]/div/div/span",
    "popup_valor_total": "/html/body/main/form[1]/div/div[3]/div[2]/div[2]/h1",
    "linhas_itens": "//*[@id='divItensDetalhado']/table/tbody", 
    "item_numero_weg": ".//input[contains(@id, 'Item') and @type='hidden']", 
    "item_preco": ".//div[contains(@class, 'form-item') and contains(., 'Preço Unitário')]//span",
    "item_qtd": ".//div[contains(@class, 'form-item') and contains(., 'Quantidade')]//span", 
    "item_aplicacao": ".//div[contains(@class, 'form-item') and contains(., 'Aplicação')]//span",
    "item_desc_span": ".//h2/span[2]",
    "item_deposito": ".//div[label[normalize-space(.)='Depósito Entrega']]//span",
    "item_ponto_descarga": ".//div[label[normalize-space(.)='Ponto de Descarga']]//span",
    "item_recebedor": ".//div[label[normalize-space(.)='Recebedor']]//span",
    "item_st": ".//div[label[contains(., 'Substituição Tributária')]]//span",
    "item_valor_st": ".//div[label[contains(., 'Valor da Substituição Tributária')]]//span",
    "btn_opcoes": "//button[text()='Opções']", 
    "btn_imprimir_menu": "//a[text()='Imprimir']", 
    "btn_imprimir_pagina": "//input[@value='Imprimir' and @class='noprint']",
    "popup_rfq": "//td[normalize-space(.)='RFQ:']/following-sibling::td"
}

# --- FUNÇÕES ROBUSTAS (ATUALIZADAS COM CONTROLE DE WAIT) ---

def robust_wait_for(wait, by, locator, condition=EC.presence_of_element_located, infinite=True):
    """
    Se infinite=True: fica num loop eterno até achar (bom para campos obrigatórios).
    Se infinite=False: respeita o timeout do 'wait' e lança erro se não achar (bom para popups opcionais).
    """
    driver = wait._driver 
    short_wait = WebDriverWait(driver, 5) 
    
    start_time = time.time()
    
    while True:
        try:
            return short_wait.until(condition((by, locator)))
        except TimeoutException:
            # Se não for infinito, verifica se estourou o tempo total
            if not infinite:
                elapsed = time.time() - start_time
                if elapsed > wait._timeout:
                    # Lança o erro para quem chamou tratar
                    raise TimeoutException(f"Timeout: Elemento {locator} não encontrado.")
            
            print(f" >>> Aguardando elemento aparecer: {locator} ... (Verifique a tela)")
        except Exception:
            time.sleep(1)

def robust_click(driver, by, locator, wait_time=25, infinite=True):
    """
    Tenta clicar. Se infinite=False, desiste após 'wait_time' segundos.
    """
    wait = WebDriverWait(driver, wait_time)
    try:
        el = robust_wait_for(wait, by, locator, EC.presence_of_element_located, infinite=infinite)
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
        time.sleep(0.3) 
        el = robust_wait_for(wait, by, locator, EC.element_to_be_clickable, infinite=infinite)
        el.click()
    except (StaleElementReferenceException, ElementClickInterceptedException):
        print(f"  > Aviso: Clique normal falhou. Tentando clique com JavaScript...")
        try:
            el = driver.find_element(by, locator) 
            driver.execute_script("arguments[0].click();", el)
        except Exception:
            if not infinite: raise 
            print(f"ERRO DE TRIGGER: Clique JS falhou. Tentando novamente...")
            raise # Vai cair no loop do caller se tiver retry, ou crashar
    except TimeoutException:
        if not infinite: 
            raise # Repassa o erro para ser tratado (ex: ignorar fechar popup)
        pass 

def robust_find_and_send_keys(driver, by, locator, keys, wait_time=25):
    wait = WebDriverWait(driver, wait_time)
    el = robust_wait_for(wait, by, locator, EC.visibility_of_element_located)
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
    time.sleep(0.3)
    el.clear()
    el.send_keys(keys)

def robust_switch_to_iframe(driver, iframe_tuple, check_tuple, wait_time=25):
    wait = WebDriverWait(driver, wait_time)
    driver.switch_to.default_content()
    robust_wait_for(wait, iframe_tuple[0], iframe_tuple[1], EC.frame_to_be_available_and_switch_to_it)
    robust_wait_for(wait, check_tuple[0], check_tuple[1], EC.presence_of_element_located)

def robust_check_checkbox(driver, by, locator, wait_time=25):
    wait = WebDriverWait(driver, wait_time)
    el = robust_wait_for(wait, by, locator, EC.presence_of_element_located)
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
    time.sleep(0.3)
    if not el.is_selected():
        driver.execute_script("arguments[0].click();", el)
        return True 
    return False 

# --- CLASSE PRINCIPAL ---

class AutomacaoPedidosApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Automação")
        self.root.geometry("300x150")
        self.root.attributes('-topmost', True)
        
        self.driver_lf = None
        self.driver_weg = None
        self.local_weg_escolhido = None
        
        self.centralizar_janela(300, 150)
        self.criar_tela_inicial()

    def centralizar_janela(self, w=300, h=150):
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def criar_tela_inicial(self):
        label = tk.Label(self.root, text="Escolha o portal WEG de origem:")
        label.pack(pady=10)
        btn_jve = tk.Button(self.root, text="Pedidos WEG Joinville", 
                            command=lambda: self.iniciar_automacao("JVE"))
        btn_jve.pack(pady=5, fill=tk.X, padx=20)
        btn_di = tk.Button(self.root, text="Pedidos WEG Dois Irmãos", 
                           command=lambda: self.iniciar_automacao("DI"))
        btn_di.pack(pady=5, fill=tk.X, padx=20)

    # --- HELPERS ---
    def force_top_window(self):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        self.centralizar_janela(350, 150)
        self.root.lift()
        self.root.focus_force() 

    def input_topmost(self, title, prompt):
        self.force_top_window()
        valor = simpledialog.askstring(title, prompt, parent=self.root)
        self.root.withdraw()
        return valor

    def info_topmost(self, title, message):
        self.force_top_window()
        messagebox.showinfo(title, message, parent=self.root)
        self.root.withdraw()

    def error_topmost(self, title, message):
        self.force_top_window()
        messagebox.showerror(title, message, parent=self.root)
        self.root.withdraw()

    def warning_topmost(self, title, message):
        self.force_top_window()
        messagebox.showwarning(title, message, parent=self.root)
        self.root.withdraw()

    def show_rfq_popup(self, rfq_number):
        self.force_top_window()
        top = tk.Toplevel(self.root)
        top.title("RFQ Encontrada")
        top.attributes('-topmost', True)
        
        ws = top.winfo_screenwidth()
        hs = top.winfo_screenheight()
        w, h = 250, 120
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        top.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        tk.Label(top, text=f"RFQ Encontrada:\n{rfq_number}", font=("Arial", 11, "bold")).pack(pady=10)
        
        def copy_and_close():
            self.root.clipboard_clear()
            self.root.clipboard_append(rfq_number)
            self.root.update()
            top.destroy()
            
        tk.Button(top, text="COPIAR", bg="green", fg="white", font=("Arial", 10, "bold"), command=copy_and_close).pack(pady=5, fill=tk.X, padx=20)
        
        self.root.wait_window(top)
    
    def organizar_janelas_lado_a_lado(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = screen_width // 2
        height = screen_height - 60 
        
        if self.driver_weg:
            self.driver_weg.set_window_rect(x=0, y=0, width=width, height=height)
        if self.driver_lf:
            self.driver_lf.set_window_rect(x=width, y=0, width=width, height=height)

    def iniciar_automacao(self, local_weg_escolhido):
        self.root.withdraw() 
        self.local_weg_escolhido = local_weg_escolhido
        try:
            creds_weg = CREDENCIAIS["WEG_JVE"] if self.local_weg_escolhido == "JVE" else CREDENCIAIS["WEG_DI"]
            creds_lf = CREDENCIAIS["LF"]
            
            print("Iniciando Portal LF...")
            self.driver_lf = webdriver.Chrome()
            self.driver_lf.get(URL_PORTAL_LF)
            # Login LF
            robust_find_and_send_keys(self.driver_lf, By.XPATH, XPATHS_LF["login_user"], creds_lf["user"])
            robust_find_and_send_keys(self.driver_lf, By.XPATH, XPATHS_LF["login_pass"], creds_lf["pass"])
            robust_click(self.driver_lf, By.XPATH, XPATHS_LF["login_btn"])
            print("Login LF OK.")
            
            # Config de Download
            print(f"Iniciando Portal WEG ({self.local_weg_escolhido}) com config de download...")
            self.download_dir = os.path.join(os.getcwd(), "pedidos_pdf")
            os.makedirs(self.download_dir, exist_ok=True)

            chrome_options = webdriver.ChromeOptions()
            prefs = {
                "download.default_directory": self.download_dir, 
                "download.prompt_for_download": False, 
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "printing.enabled": True,
                "savefile.default_directory": self.download_dir
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            self.driver_weg = webdriver.Chrome(options=chrome_options)
            self.driver_weg.get(URL_PORTAL_WEG_JVE)
            robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["login_user"], creds_weg["user"])
            robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["login_pass"], creds_weg["pass"])
            robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["login_btn"])
            print("Login WEG OK.")
            
            time.sleep(3) 
            self.organizar_janelas_lado_a_lado()
            
            self.info_topmost("Pronto", "Logins efetuados. Vamos começar.")
            self.loop_pedir_numero_pedido()
        except Exception as e:
            self.error_topmost("Erro no Login", f"Não foi possível fazer login.\nErro: {e}")
            self.fechar_tudo()
            
    def loop_pedir_numero_pedido(self):
        while True:
            self.organizar_janelas_lado_a_lado()
            self.numero_pedido_atual = self.input_topmost("Novo Pedido", "Cole o número do pedido WEG:")
            if not self.numero_pedido_atual: break 
            
            try:
                print(f"Buscando dados do pedido {self.numero_pedido_atual} no portal WEG...")
                dados_pedido = self.buscar_dados_weg(self.numero_pedido_atual)
                
                if not dados_pedido:
                    self.warning_topmost("Não Encontrado", f"Pedido {self.numero_pedido_atual} não localizado.")
                    continue
                if dados_pedido.get("aviso_extracao"):
                    self.warning_topmost("Aviso na Extração", dados_pedido["aviso_extracao"])
                    
                if dados_pedido.get("rfq"):
                    self.show_rfq_popup(dados_pedido["rfq"])
                
                print("Dados encontrados. Preenchendo portal LF...")
                try:
                    if len(self.driver_lf.find_elements(By.XPATH, XPATHS_LF["nav_pedido_link"])) == 0:
                        raise Exception("Menu fechado")
                    if not self.driver_lf.find_element(By.XPATH, XPATHS_LF["nav_pedido_link"]).is_displayed():
                        raise Exception("Menu fechado")
                except Exception:
                     robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_menu_btn"], wait_time=5)
                     time.sleep(0.5)
                
                robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_pedido_link"])
                time.sleep(1) 
                
                robust_click(self.driver_lf, By.XPATH, XPATHS_LF["novo_pedido_btn"])
                print("Página 'Novo Pedido' aberta. Preenchendo...")

                self.preencher_pedido_lf(dados_pedido) 
                
                print("Cabeçalho salvo. Verificando checkbox de industrialização...")
                if self.local_weg_escolhido == "JVE" and dados_pedido.get("aplicacao_final") == "Industrialização":
                    self.clicar_checkbox_industrializacao() 
                
                print("Iniciando preenchimento de itens no LF...")
                self.preencher_itens_lf(dados_pedido) 

                print("Itens preenchidos. Voltando ao portal WEG para baixar o PDF.")
                caminho_pdf = self.baixar_pdf_weg()
                
                # --- CORREÇÃO DO TRAVAMENTO AO FECHAR POPUP WEG ---
                try:
                    print("Download concluído. Tentando fechar o pop-up do pedido WEG...")
                    self.driver_weg.switch_to.window(self.driver_weg.window_handles[0])
                    self.driver_weg.switch_to.default_content()
                    
                    # Usa infinite=False para não travar se o botão não existir ou mudar
                    robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["popup_btn_fechar_X"], wait_time=5, infinite=False)
                    print("  > Pop-up fechado.")
                except Exception as e:
                    print(f"  > (Aviso) Não foi possível fechar o pop-up da WEG automaticamente. Ignorando e seguindo... {e}")
                # --------------------------------------------------
                
                if not caminho_pdf:
                    self.error_topmost("Falha no Download", "Não foi possível baixar o PDF. O anexo não será feito.")
                    continue
                
                print(f"PDF baixado. Voltando ao portal LF para anexar o arquivo: {caminho_pdf}")
                self.anexar_pdf_lf(caminho_pdf) 
                
                self.info_topmost("Sucesso", f"Pedido {self.numero_pedido_atual} processado com sucesso!")
            except Exception as e:
                self.error_topmost("Erro no Processamento", f"Ocorreu um erro ao processar o pedido {self.numero_pedido_atual}.\nErro: {e}")
        self.fechar_tudo()

    def _limpar_valor_numerico(self, texto_sujo):
        try:
            return re.sub(r'[^\d,]', '', texto_sujo)
        except Exception: return ""

    def _converter_para_float(self, texto_sujo):
        try:
            valor_limpo = re.sub(r'[^\d,]', '', texto_sujo)
            return float(valor_limpo.replace(',', '.'))
        except Exception: return 0.0 

    def buscar_dados_weg(self, numero_pedido):
        print(f"Procurando campo 'Pesquisar'...")
        try:
            robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["campo_busca_pedido"], numero_pedido + Keys.RETURN)
            print(f"Busca por {numero_pedido} enviada...")
            time.sleep(2) 
            
            link_xpath_dinamico = f"//a[contains(text(), '{numero_pedido}')]"
            robust_click(self.driver_weg, By.XPATH, link_xpath_dinamico)
            print("Link do pedido clicado.")
        except Exception as e:
            print(f"Erro ao buscar pedido: {e}"); return None
        
        try:
            robust_switch_to_iframe(self.driver_weg, (By.XPATH, XPATHS_WEG["popup_iframe"]), (By.XPATH, XPATHS_WEG["popup_iframe_check"]))
        except Exception as e:
            self.error_topmost("Erro de Iframe", f"Não consegui 'focar' no pop-up (iframe)."); return None

        dados_extraidos = {"itens": [], "aviso_extracao": ""}
        wait_weg = WebDriverWait(self.driver_weg, 10)
        try:
            print("Extraindo dados...")
            cnpj_element = robust_wait_for(wait_weg, By.XPATH, XPATHS_WEG["popup_cnpj"])
            dados_extraidos["cnpj"] = cnpj_element.text
            
            cond_pagto_element = robust_wait_for(wait_weg, By.XPATH, XPATHS_WEG["popup_cond_pagto"])
            cond_valor_bruto = cond_pagto_element.text.split('\n')[-1].strip() 
            dados_extraidos["cond_pagto"] = "60" if "60" in cond_valor_bruto else "45"
            
            cif_fob_element = robust_wait_for(wait_weg, By.XPATH, XPATHS_WEG["popup_cif_fob"])
            frete_valor_bruto = cif_fob_element.text
            dados_extraidos["cif_fob"] = "FOB" if "FOB" in frete_valor_bruto.upper() else "CIF"
            
            valor_total_el = robust_wait_for(wait_weg, By.XPATH, XPATHS_WEG["popup_valor_total"])
            dados_extraidos["valor_total"] = self._converter_para_float(valor_total_el.text)
            
            try:
                rfq_el = self.driver_weg.find_element(By.XPATH, XPATHS_WEG["popup_rfq"])
                dados_extraidos["rfq"] = rfq_el.text.strip()
                print(f"RFQ encontrada: {dados_extraidos['rfq']}")
            except Exception:
                dados_extraidos["rfq"] = ""

            elementos_itens = wait_weg.until(EC.presence_of_all_elements_located((By.XPATH, XPATHS_WEG["linhas_itens"])))
            aplicacao_mestra = None
            dados_extraidos["obs_itens"] = []
            
            for i in range(0, len(elementos_itens), 2):
                item_header_el = elementos_itens[i]
                item_details_el = elementos_itens[i+1]
                
                item_data = {}
                item_data["preco"] = self._limpar_valor_numerico(item_details_el.find_element(By.XPATH, XPATHS_WEG["item_preco"]).text)
                item_data["qtd"] = self._limpar_valor_numerico(item_details_el.find_element(By.XPATH, XPATHS_WEG["item_qtd"]).text)
                item_data["aplicacao"] = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_aplicacao"]).text
                item_data["numero_weg"] = item_header_el.find_element(By.XPATH, XPATHS_WEG["item_numero_weg"]).get_attribute("value")
                try: item_data["descricao"] = item_header_el.find_element(By.XPATH, XPATHS_WEG["item_desc_span"]).text
                except Exception: item_data["descricao"] = f"Item {i//2 + 1}"

                if i == 0:
                    try: 
                        dep = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_deposito"]).text
                        if dep: dados_extraidos["obs_itens"].append(f"Depósito: {dep}")
                    except Exception: pass
                    try:
                        ponto = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_ponto_descarga"]).text
                        if ponto: dados_extraidos["obs_itens"].append(f"Ponto Descarga: {ponto}")
                    except Exception: pass
                    try:
                        receb = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_recebedor"]).text
                        if receb: dados_extraidos["obs_itens"].append(f"Recebedor: {receb}")
                    except Exception: pass
                    try:
                        sts = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_st"]).text
                        if sts: dados_extraidos["obs_itens"].append(f"ST: {sts}")
                    except Exception: pass
                    try:
                        vst = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_valor_st"]).text
                        if vst: dados_extraidos["obs_itens"].append(f"Valor ST: {vst}")
                    except Exception: pass

                if not aplicacao_mestra: aplicacao_mestra = item_data["aplicacao"]
                elif aplicacao_mestra != item_data["aplicacao"]:
                    dados_extraidos["aviso_extracao"] = "AVISO: Itens misturados (Consumo/Ind)!"
                
                dados_extraidos["itens"].append(item_data)
                
            dados_extraidos["aplicacao_final"] = aplicacao_mestra
            return dados_extraidos
        except Exception as e:
            print(f"Erro na extração: {e}")
            self.driver_weg.switch_to.default_content(); return None
        
    def preencher_pedido_lf(self, dados_pedido):
        wait = WebDriverWait(self.driver_lf, 25)
        
        def abrir_popup_e_selecionar(valor_busca, termo_click=None, exact_match=False):
            self.driver_lf.switch_to.default_content()
            texto_a_esperar = termo_click if termo_click else valor_busca
            try:
                base_xpath_popup_visivel = "//div[contains(@class,'ui-dialog-popuplov') and not(contains(@style,'display: none'))]"
                wait.until(EC.visibility_of_element_located((By.XPATH, base_xpath_popup_visivel)))

                xpath_campo_busca = f"{base_xpath_popup_visivel}//input[contains(@class, 'a-PopupLOV-search') and @type='text']"
                xpath_botao_busca = f"{base_xpath_popup_visivel}//button[contains(@class,'a-PopupLOV-doSearch')]"
                
                campo_busca = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_campo_busca)))
                botao_busca = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_busca)))
                
                campo_busca.clear()
                campo_busca.send_keys(valor_busca)
                botao_busca.click()
                
                # --- TRATAMENTO LOAD MORE ROWS ---
                try:
                    btn_load_more = WebDriverWait(self.driver_lf, 2).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "button.js-load.a-GV-loadMoreButton"))
                    )
                    if btn_load_more:
                        print("  > Clicando em 'Load More Rows'...")
                        btn_load_more.click()
                        time.sleep(1.5)
                except TimeoutException: pass
                # ---------------------------------

                xpath_resultado_especifico = ""
                if exact_match:
                    xpath_resultado_especifico = f"{base_xpath_popup_visivel}//ul/li[normalize-space(.)='{texto_a_esperar}']"
                else:
                    xpath_resultado_especifico = f"{base_xpath_popup_visivel}//ul/li[contains(., '{texto_a_esperar}')]"

                resultado = robust_wait_for(wait, By.XPATH, xpath_resultado_especifico, EC.element_to_be_clickable)
                resultado.click()
                print(f"✅ Selecionado: {texto_a_esperar}")
            except Exception as e:
                print(f"❌ Erro no popup: {e}"); raise e

        try:
            cnpj_limpo = re.sub(r'[^\d]', '', dados_pedido.get('cnpj', ''))
            cond_pagto = dados_pedido.get('cond_pagto', '45')
            frete = dados_pedido.get('cif_fob', 'CIF')
            valor_total = dados_pedido.get('valor_total', 0.0) 
            
            transportadora_texto = ""
            if cnpj_limpo == "10953379000108" and frete == "FOB": transportadora_texto = "83630053000113"
            elif self.local_weg_escolhido == "DI" and frete == "FOB": transportadora_texto = "83630053000113"
            elif self.local_weg_escolhido == "JVE": transportadora_texto = "46918010000101" if frete == "CIF" else "05399007000141"
            else: transportadora_texto = "03350624000550"

            robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

            filial_valor = "22" if self.local_weg_escolhido == "JVE" else "23"
            Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_filial"])).select_by_value(filial_valor)
            
            robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_VW_CLIENTE_ID_lov_btn')]")
            abrir_popup_e_selecionar(cnpj_limpo) 
            robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

            aplicacao_weg = dados_pedido.get("aplicacao_final", "Industrialização")
            aplicacao_lf = MAPA_APLICACAO.get(aplicacao_weg, "USO E CONSUMO/IMOBILIZADO")
            Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_finalidade"])).select_by_visible_text(aplicacao_lf)

            Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_tipo_nf"])).select_by_visible_text("VENDA")
            
            robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_ordem_compra"], self.numero_pedido_atual)

            robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_VEN_FORMA_PAGAMENTO_ID_lov_btn')]")
            abrir_popup_e_selecionar("CARTEIRA", exact_match=True) 
            robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

            robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_CONDICAO_PAGAMENTO_ID_lov_btn')]")
            abrir_popup_e_selecionar(cond_pagto, exact_match=True) 
            robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

            try: Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_tipo_frete"])).select_by_visible_text(frete)
            except Exception: pass

            robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_VW_TRANSPORTADORA_ID_lov_btn')]")
            abrir_popup_e_selecionar(transportadora_texto) 
            robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
            
            if frete == "CIF" and valor_total < 400.0:
                robust_check_checkbox(self.driver_lf, By.ID, XPATHS_LF["radio_ped_complementar_sim"])
            
            obs_texto = "\n".join(dados_pedido.get("obs_itens", []))
            robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_obs_nf"], obs_texto)
            
            if dados_pedido.get("rfq"):
                FRASE_PADRAO = "Preço conforme RFQ, aut. RAFAEL."
            else:
                FRASE_PADRAO = "Pedido acordo WEG"
                
            robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_obs_pedido"], FRASE_PADRAO)
            
            time.sleep(1) 
            robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_salvar_cabecalho"])
            self.driver_lf.switch_to.default_content()
        except Exception as e:
            self.error_topmost("Erro no Preenchimento LF", f"Erro: {e}")
            self.driver_lf.switch_to.default_content(); raise e
            
    def preencher_itens_lf(self, dados_pedido):
        wait_lf = WebDriverWait(self.driver_lf, 25)
        if not dados_pedido.get("itens"): return

        for item_weg in dados_pedido["itens"]:
            try:
                robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_adicionar_item"])
                
                robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["item_popup_iframe"]), (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                
                prompt_texto = f"Item: {item_weg['numero_weg']} - {item_weg['descricao']}\nQtd: {item_weg['qtd']}\nPreço: {item_weg['preco']}\n\nCódigo LF (SKU):"
                codigo_lf = self.input_topmost("SKU Item", prompt_texto)
                if not codigo_lf: raise Exception("Cancelado")
                
                codigo_lf = codigo_lf.strip()
                if '/' in codigo_lf:
                    codigo_lf = codigo_lf.split('/')[0]
                if len(codigo_lf) < 6 and codigo_lf.isdigit():
                    codigo_lf = codigo_lf.zfill(6)
                
                robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_sku_btn"])
                self.driver_lf.switch_to.default_content()
                
                base_xpath_popup_visivel = "//div[contains(@class,'ui-dialog-popuplov') and not(contains(@style,'display: none'))]"
                xpath_campo_busca_sku = f"{base_xpath_popup_visivel}//input[contains(@class, 'a-PopupLOV-search') and @type='text']"
                
                search_el = robust_wait_for(wait_lf, By.XPATH, xpath_campo_busca_sku, EC.visibility_of_element_located)
                search_el.clear(); search_el.send_keys(codigo_lf); search_el.send_keys(Keys.RETURN)
                
                xpath_resultado_sku = f"{base_xpath_popup_visivel}{XPATHS_LF['item_sku_search_result']}"
                robust_click(self.driver_lf, By.XPATH, xpath_resultado_sku)
                
                robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["item_popup_iframe"]), (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                
                robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_ordem_compra"], self.numero_pedido_atual)
                robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_item_oc"], item_weg['numero_weg'])
                robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_qtd"], item_weg['qtd'])
                robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_preco"], item_weg['preco'])
                
                robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_btn_salvar"])
                self.driver_lf.switch_to.default_content()
                
                wait_lf.until(EC.invisibility_of_element_located((By.XPATH, XPATHS_LF["item_popup_iframe"])))
            except Exception as e:
                self.error_topmost("Erro no Item", f"Erro no item {item_weg['numero_weg']}: {e}")
                self.driver_lf.switch_to.default_content(); break 

    def clicar_checkbox_industrializacao(self):
        try:
            wait_lf = WebDriverWait(self.driver_lf, 10) 
            checkbox = wait_lf.until(EC.presence_of_element_located((By.XPATH, XPATHS_LF["checkbox_industrializacao"])))
            if not checkbox.is_selected(): self.driver_lf.execute_script("arguments[0].click();", checkbox)
        except Exception: pass

    def baixar_pdf_weg(self):
        wait_weg = WebDriverWait(self.driver_weg, 20)
        nome_arquivo = f"{self.numero_pedido_atual}.pdf"
        caminho_completo_pdf = os.path.join(self.download_dir, nome_arquivo)
        if os.path.exists(caminho_completo_pdf): os.remove(caminho_completo_pdf)
            
        try:
            janela_principal = self.driver_weg.current_window_handle
            self.driver_weg.switch_to.window(janela_principal)
            
            robust_switch_to_iframe(self.driver_weg, (By.XPATH, XPATHS_WEG["popup_iframe"]), (By.XPATH, XPATHS_WEG["popup_iframe_check"]))
            robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["btn_opcoes"])
            robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["btn_imprimir_menu"])
            
            wait_weg.until(EC.number_of_windows_to_be(2))
            janela_impressao = [handle for handle in self.driver_weg.window_handles if handle != janela_principal][0]
            self.driver_weg.switch_to.window(janela_impressao)
            
            robust_wait_for(wait_weg, By.XPATH, XPATHS_WEG["btn_imprimir_pagina"])
            time.sleep(2) 
            
            print_settings = {"printBackground": True, "paperWidth": 8.27, "paperHeight": 11.69, "scale": 1.0}
            result = self.driver_weg.execute_cdp_cmd("Page.printToPDF", print_settings)
            with open(caminho_completo_pdf, 'wb') as f: f.write(base64.b64decode(result['data']))

            self.driver_weg.close()
            self.driver_weg.switch_to.window(janela_principal)
            return caminho_completo_pdf
        except Exception as e:
            try:
                if len(self.driver_weg.window_handles) > 1:
                    self.driver_weg.switch_to.window(self.driver_weg.window_handles[1]); self.driver_weg.close()
                    self.driver_weg.switch_to.window(self.driver_weg.window_handles[0])
            except: pass
            return None
    
    def anexar_pdf_lf(self, caminho_pdf):
        wait_lf = WebDriverWait(self.driver_lf, 25)
        try:
            self.driver_lf.switch_to.window(self.driver_lf.window_handles[0])
            self.driver_lf.switch_to.default_content()
            
            robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_anexar"])
            robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["anexo_iframe"]), (By.ID, XPATHS_LF["anexo_iframe_check"]))
            
            robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["anexo_descricao_input"], self.numero_pedido_atual)
            
            campo_upload = robust_wait_for(wait_lf, By.XPATH, XPATHS_LF["anexo_upload_input"], EC.presence_of_element_located)
            campo_upload.send_keys(os.path.abspath(caminho_pdf))
            
            robust_click(self.driver_lf, By.XPATH, XPATHS_LF["anexo_btn_salvar"])
            self.driver_lf.switch_to.default_content()
            wait_lf.until(EC.invisibility_of_element_located((By.XPATH, XPATHS_LF["anexo_iframe"])))
            
        except Exception as e:
            self.error_topmost("Erro ao Anexar", f"Erro: {e}")
            self.driver_lf.switch_to.default_content(); raise e
            
    def fechar_tudo(self):
        if self.driver_lf: self.driver_lf.quit()
        if self.driver_weg: self.driver_weg.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoPedidosApp(root)
    root.mainloop()