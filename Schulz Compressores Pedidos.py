import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select 
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
import time
import os 
import pandas as pd
from dotenv import load_dotenv 
import re

# --- 1. CONFIGURAÇÕES ---

load_dotenv()

URL_PORTAL_LF = "https://erp.bem.inf.br/ords/f?p=8000000:9999:2963151475408:::::"

# Dados Fixos Schulz
CNPJ_SCHULZ = "23635798000143"
CONDICAO_PAGTO_TEXTO = "90" 
FRETE_TIPO = "CIF"
CNPJ_TRANSPORTADORA = "53417977000175" # JEVERSON MELO

# --- TEXTOS DE OBSERVAÇÃO ---
# Texto que vai no campo "Observação NF" (P7_MSG_OBS) quando é Industrialização
OBS_NF_INDUSTRIALIZACAO = "Diferimento de ICMS TTD 490\nTermo de concessão nº 205000004022828"

# Texto FIXO que vai no campo "Observação Pedido" (P7_OBS)
OBS_FIXA_PEDIDO = "Preço autorizado Rafael Schulze"

CREDENCIAIS = {
    "LF": {
        "user": os.environ.get("LF_USER"), 
        "pass": os.environ.get("LF_PASS")
    }
}

if not all([CREDENCIAIS["LF"]["user"], CREDENCIAIS["LF"]["pass"]]):
    messagebox.showerror("Erro de Credenciais", "Configure as variáveis de ambiente LF_USER e LF_PASS no arquivo .env")
    exit()

# XPATHs LF 
XPATHS_LF = {
    "login_user": '//*[@id="P9999_USERNAME"]',
    "login_pass": '//*[@id="P9999_PASSWORD"]',
    "login_btn": "//button[contains(., 'Acessar')]", 
    "nav_menu_btn": "/html/body/form/header/div[1]/div[1]/button",
    "nav_pedido_link": "/html/body/form/div[1]/div[1]/div/ul/li[9]/div[2]/a",
    "novo_pedido_btn_opcoes": [
        "/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[1]/div/div/div[1]/div[3]/button[1]",
        "//button[span[contains(text(), 'Novo Pedido')]]",
        "//button[contains(., 'Novo')]"
    ],
    "novo_pedido_iframe": "//iframe[@title='Pedido']", 
    "novo_pedido_iframe_check": "P7_AD_ORG_ID", 
    "campo_filial": "P7_AD_ORG_ID",
    "campo_finalidade": "P7_BEM_VEN_FINALIDADE_ID",
    "campo_tipo_nf": "P7_BEG_TIPO_NF_ID",
    "campo_ordem_compra": "P7_NRO_ORDEM_COMPRA",
    "campo_tipo_frete": "P7_BEM_TIPO_FRETE_ID",
    
    # --- CAMPOS DE OBSERVAÇÃO ---
    "campo_obs_nf": "P7_MSG_OBS",      # Onde vai o TTD 490
    "campo_obs_pedido": "P7_OBS",      # Onde vai "Preço autorizado..."
    # ----------------------------

    "btn_salvar_cabecalho": "//button[span[normalize-space(.)='Incluir'] and not(ancestor::div[contains(@class,'ui-dialog-content')])]",
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
    
    # Checkbox e Radio
    "checkbox_industrializacao_smart": "//tr[.//td[contains(text(), 'TTD 490')]]//input[@type='checkbox']",
    "radio_ped_complementar_sim_label": "//label[@for='P7_IND_PED_COMPLEMENTAR_0']"
}

# --- FUNÇÕES ROBUSTAS ---

def robust_wait_for(wait, by, locator, condition=EC.presence_of_element_located):
    try:
        return wait.until(condition((by, locator)))
    except TimeoutException:
        raise Exception(f"Elemento não encontrado: {locator}")

def robust_click(driver, by, locator, wait_time=20):
    wait = WebDriverWait(driver, wait_time)
    try:
        el = robust_wait_for(wait, by, locator, EC.presence_of_element_located)
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
        time.sleep(0.5) 
        el = robust_wait_for(wait, by, locator, EC.element_to_be_clickable)
        el.click()
    except Exception:
        try:
            el = driver.find_element(by, locator) 
            driver.execute_script("arguments[0].click();", el)
        except Exception as e:
            raise Exception(f"Não consegui clicar em {locator}. Erro: {e}")

def robust_find_and_send_keys(driver, by, locator, keys, wait_time=20):
    wait = WebDriverWait(driver, wait_time)
    try:
        el = robust_wait_for(wait, by, locator, EC.visibility_of_element_located)
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
        time.sleep(0.3)
        el.clear()
        el.send_keys(str(keys))
    except Exception as e:
        raise Exception(f"Erro ao digitar em {locator}: {e}")

def robust_switch_to_iframe(driver, iframe_tuple, check_tuple, wait_time=20):
    wait = WebDriverWait(driver, wait_time)
    try:
        driver.switch_to.default_content()
        robust_wait_for(wait, iframe_tuple[0], iframe_tuple[1], EC.frame_to_be_available_and_switch_to_it)
        robust_wait_for(wait, check_tuple[0], check_tuple[1], EC.presence_of_element_located)
    except Exception as e:
        raise Exception(f"Erro ao entrar no iframe: {e}")

def robust_check_element(driver, by, locator, wait_time=15):
    wait = WebDriverWait(driver, wait_time)
    try:
        el = robust_wait_for(wait, by, locator, EC.presence_of_element_located)
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
        time.sleep(0.3)
        if "label" in locator:
            el.click()
        elif not el.is_selected():
            try: el.click()
            except: driver.execute_script("arguments[0].click();", el)
        print(f"   > Elemento clicado/marcado: {locator}")
    except Exception as e:
        print(f"   ! Erro ao marcar elemento {locator}: {e}")

# --- CLASSE PRINCIPAL ---

class AutomacaoSchulzApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Automação Schulz (Obs Corrigida)")
        self.root.geometry("400x250")
        self.root.attributes('-topmost', True)
        self.driver_lf = None
        self.pedidos_processados = [] 
        self.criar_tela_inicial()

    def criar_tela_inicial(self):
        tk.Label(self.root, text="Automação Schulz Completa", font=("Arial", 12, "bold")).pack(pady=10)

        frame_coluna = tk.Frame(self.root)
        frame_coluna.pack(pady=5)
        tk.Label(frame_coluna, text="Coluna de Preço (Planilha Códigos):").pack(side=tk.LEFT)
        self.coluna_preco_var = tk.StringVar(value="H")
        tk.Entry(frame_coluna, textvariable=self.coluna_preco_var, width=5, justify="center").pack(side=tk.LEFT, padx=5)

        tk.Button(self.root, text="Carregar Planilhas e Iniciar", command=self.iniciar_processo, height=2, bg="#ddd").pack(pady=10, fill=tk.X, padx=30)

    def force_top_window(self):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        self.root.lift()

    def input_topmost(self, title, prompt):
        self.force_top_window()
        return simpledialog.askstring(title, prompt, parent=self.root)

    def iniciar_processo(self):
        path_pedidos = filedialog.askopenfilename(title="Selecione PEDIDOS", filetypes=[("Excel", "*.xlsx *.xls")])
        if not path_pedidos: return
        path_codigos = filedialog.askopenfilename(title="Selecione CÓDIGOS (Com Tipo e Preço)", filetypes=[("Excel", "*.xlsx *.xls")])
        if not path_codigos: return

        try:
            col_letra = self.coluna_preco_var.get().strip().upper()
            if not col_letra.isalpha():
                col_letra = "H"
            
            idx_preco = 0
            for char in col_letra:
                idx_preco = idx_preco * 26 + (ord(char) - ord('A') + 1)
            idx_preco -= 1

            self.root.withdraw()
            # Leitura
            df_ped = pd.read_excel(path_pedidos, header=0, usecols=[2, 5, 10]) 
            df_ped.columns = ['cod_item', 'ordem_compra', 'qtd']

            df_cod = pd.read_excel(path_codigos, header=0, usecols=[0, 2, 3, idx_preco])
            df_cod.columns = ['cod_item', 'tipo_uso', 'cod_lf', 'preco']
            
            # Limpeza
            df_ped = df_ped.dropna(subset=['ordem_compra'])
            
            def padronizar_codigo(codigo):
                c = str(codigo).strip()
                if c.endswith('.0'): c = c[:-2]
                if '/' in c: c = c.split('/')[0]
                if len(c) < 6 and c.isdigit(): c = c.zfill(6)
                return c

            df_ped['cod_item'] = df_ped['cod_item'].apply(padronizar_codigo)
            df_ped['ordem_compra'] = df_ped['ordem_compra'].astype(str).str.strip().replace(r'\.0$', '', regex=True)
            
            df_cod['cod_item'] = df_cod['cod_item'].apply(padronizar_codigo)
            df_cod['cod_lf'] = df_cod['cod_lf'].apply(padronizar_codigo)
            df_cod['tipo_uso'] = df_cod['tipo_uso'].astype(str).str.strip()
            
            df_ped['qtd'] = pd.to_numeric(df_ped['qtd'], errors='coerce').fillna(0)
            df_cod['preco'] = pd.to_numeric(df_cod['preco'], errors='coerce').fillna(0)

            # Cruzamento
            self.pedidos_processados = pd.merge(df_ped, df_cod, on='cod_item', how='left')
            
            print("Planilhas carregadas. Iniciando automação...")
            self.iniciar_automacao()
            
        except Exception as e:
            self.force_top_window()
            messagebox.showerror("Erro", f"Erro planilhas: {e}")

    def iniciar_automacao(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            self.driver_lf = webdriver.Chrome(options=options)
            self.driver_lf.get(URL_PORTAL_LF)
            
            robust_find_and_send_keys(self.driver_lf, By.XPATH, XPATHS_LF["login_user"], CREDENCIAIS["LF"]["user"])
            robust_find_and_send_keys(self.driver_lf, By.XPATH, XPATHS_LF["login_pass"], CREDENCIAIS["LF"]["pass"])
            robust_click(self.driver_lf, By.XPATH, XPATHS_LF["login_btn"])
            
            WebDriverWait(self.driver_lf, 30).until(EC.presence_of_element_located((By.XPATH, XPATHS_LF["nav_menu_btn"])))
            self.loop_pedidos()
            
        except Exception as e:
            messagebox.showerror("Erro Login", f"{e}")

    def loop_pedidos(self):
        grupos = self.pedidos_processados.groupby('ordem_compra')
        
        for ordem_compra, itens in grupos:
            try:
                print(f"--- Pedido: {ordem_compra} ---")
                
                tipo_uso_pedido = itens.iloc[0]['tipo_uso']
                valor_total_pedido = (itens['qtd'] * itens['preco']).sum()
                
                print(f"   > Tipo: {tipo_uso_pedido}")
                print(f"   > Valor Total: R$ {valor_total_pedido:.2f}")

                self.driver_lf.switch_to.default_content()
                try: robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_pedido_link"], wait_time=3)
                except:
                    robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_menu_btn"])
                    time.sleep(1)
                    robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_pedido_link"])
                
                time.sleep(2) 

                btn_clicado = False
                for xpath in XPATHS_LF["novo_pedido_btn_opcoes"]:
                    try:
                        robust_click(self.driver_lf, By.XPATH, xpath, wait_time=2)
                        btn_clicado = True; break
                    except: continue
                if not btn_clicado: raise Exception("Botão Novo não encontrado.")

                # Preencher Cabeçalho
                self.preencher_cabecalho(ordem_compra, tipo_uso_pedido, valor_total_pedido)
                
                # Checkbox TTD (se for Industrialização)
                if "CONSUMO" not in str(tipo_uso_pedido).upper():
                    self.clicar_checkbox_industrializacao()

                # Itens
                self.preencher_itens(itens, ordem_compra)
                
                print(f"Pedido {ordem_compra} OK.")

            except Exception as e:
                try: self.driver_lf.find_element(By.CSS_SELECTOR, ".ui-dialog-titlebar-close").click()
                except: pass
                self.driver_lf.switch_to.default_content()
                
                self.force_top_window()
                if not messagebox.askyesno("Erro", f"Erro no pedido {ordem_compra}:\n{e}\n\nContinuar?"): break
        
        messagebox.showinfo("Fim", "Concluído")

    # --- FUNÇÃO DE POPUP ---
    def abrir_popup_e_selecionar(self, valor_busca, exact_match=False):
        wait = WebDriverWait(self.driver_lf, 20)
        self.driver_lf.switch_to.default_content()
        
        xpath_popup = "//div[contains(@class,'ui-dialog-popuplov') and not(contains(@style,'display: none'))]"
        wait.until(EC.visibility_of_element_located((By.XPATH, xpath_popup)))

        xpath_input = f"{xpath_popup}//input[contains(@class, 'a-PopupLOV-search')]"
        xpath_btn_search = f"{xpath_popup}//button[contains(@class,'a-PopupLOV-doSearch')]"
        xpath_load_more = f"{xpath_popup}//button[contains(text(), 'Load More') or contains(text(), 'Carregar Mais')]"
        
        inp = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_input)))
        inp.clear()
        valor_limpo = str(valor_busca).strip().replace('.0', '')
        print(f"   [POPUP] Pesquisando por: '{valor_limpo}' (Exato={exact_match})")
        
        inp.send_keys(valor_limpo)
        try: self.driver_lf.find_element(By.XPATH, xpath_btn_search).click()
        except: inp.send_keys(Keys.RETURN)
            
        time.sleep(2.5) 
        
        encontrado = False
        tentativas_load = 0
        max_load = 3

        while not encontrado and tentativas_load <= max_load:
            xpath_alvo = []
            if exact_match:
                xpath_alvo = [
                    f"{xpath_popup}//td[normalize-space(.)='{valor_limpo}']",
                    f"{xpath_popup}//li[normalize-space(.)='{valor_limpo}']",
                    f"{xpath_popup}//a[normalize-space(.)='{valor_limpo}']"
                ]
            else:
                xpath_alvo = [
                    f"{xpath_popup}//div[contains(@class, 'a-PopupLOV-results')]//table//tr[1]//td[1]", 
                    f"{xpath_popup}//div[contains(@class, 'a-PopupLOV-results')]//li[1]",             
                    f"{xpath_popup}//div[contains(@class, 'a-PopupLOV-results')]//a"                   
                ]

            for xpath in xpath_alvo:
                try:
                    elementos = self.driver_lf.find_elements(By.XPATH, xpath)
                    if len(elementos) > 0:
                        el = elementos[0]
                        self.driver_lf.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
                        time.sleep(0.5)
                        try: el.click()
                        except: self.driver_lf.execute_script("arguments[0].click();", el)
                        
                        print(f"✅ Sucesso: Selecionado via {xpath}")
                        encontrado = True
                        break
                except: continue
            
            if encontrado: break
            
            try:
                btn_load = self.driver_lf.find_element(By.XPATH, xpath_load_more)
                if btn_load.is_displayed():
                    print("   ⚠️ Clicando em 'Load More'...")
                    self.driver_lf.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", btn_load)
                    btn_load.click()
                    time.sleep(2)
                    tentativas_load += 1
                else: break
            except: break
        
        if not encontrado:
            print(f"❌ ERRO: Nada encontrado para '{valor_limpo}'")
            self.force_top_window()
            messagebox.showinfo("Ajuda", f"Selecione manualmente '{valor_limpo}' e dê OK.")

    def preencher_cabecalho(self, ordem_compra, tipo_uso, valor_total):
        wait = WebDriverWait(self.driver_lf, 25)
        
        robust_switch_to_iframe(self.driver_lf, 
                                (By.XPATH, XPATHS_LF["novo_pedido_iframe"]),
                                (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
        
        # 1. FILIAL 22
        try:
            el_filial = robust_wait_for(wait, By.ID, XPATHS_LF["campo_filial"])
            Select(el_filial).select_by_value("22") 
        except: pass

        # 2. CLIENTE
        robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_VW_CLIENTE_ID_lov_btn')]")
        self.abrir_popup_e_selecionar(CNPJ_SCHULZ, exact_match=False)
        robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

        # 3. DADOS BÁSICOS
        eh_industrializacao = "CONSUMO" not in str(tipo_uso).upper()
        
        # Define os textos de observação
        texto_obs_pedido = OBS_FIXA_PEDIDO
        texto_obs_nf = ""

        if eh_industrializacao:
            Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_finalidade"])).select_by_visible_text("INDUSTRIALIZAÇÃO")
            texto_obs_nf = OBS_NF_INDUSTRIALIZACAO
            print("   > Modo: Industrialização (TTD 490)")
        else:
            Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_finalidade"])).select_by_visible_text("USO E CONSUMO/IMOBILIZADO")
            print("   > Modo: Consumo (Padrão)")

        Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_tipo_nf"])).select_by_visible_text("VENDA")
        robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_ordem_compra"], ordem_compra)

        # 4. FORMA PAGAMENTO
        robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_VEN_FORMA_PAGAMENTO_ID_lov_btn')]")
        self.abrir_popup_e_selecionar("CARTEIRA", exact_match=True)
        robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

        # 5. CONDIÇÃO PAGTO
        robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_CONDICAO_PAGAMENTO_ID_lov_btn')]")
        self.abrir_popup_e_selecionar(CONDICAO_PAGTO_TEXTO, exact_match=True)
        robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

        # 6. TRANSP
        try: Select(robust_wait_for(wait, By.ID, XPATHS_LF["campo_tipo_frete"])).select_by_visible_text(FRETE_TIPO)
        except: pass

        robust_click(self.driver_lf, By.XPATH, "//button[contains(@id,'P7_BEG_VW_TRANSPORTADORA_ID_lov_btn')]")
        self.abrir_popup_e_selecionar(CNPJ_TRANSPORTADORA, exact_match=False)
        robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))

        # 7. PEDIDO COMPLEMENTAR (Sempre SIM se CIF < 400)
        if FRETE_TIPO == "CIF" and valor_total < 400.0:
            print(f"   > Valor R$ {valor_total:.2f} < 400 (CIF). Marcando Pedido Complementar.")
            robust_check_element(self.driver_lf, By.XPATH, XPATHS_LF["radio_ped_complementar_sim_label"])

        # 8. OBSERVAÇÕES (CORRIGIDO)
        # Observação NF (P7_MSG_OBS) -> Recebe TTD 490 se Ind.
        robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_obs_nf"], texto_obs_nf)
        
        # Observação Pedido (P7_OBS) -> Recebe a frase Fixa
        robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_obs_pedido"], texto_obs_pedido)

        robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_salvar_cabecalho"])
        self.driver_lf.switch_to.default_content()

    def clicar_checkbox_industrializacao(self):
        self.driver_lf.switch_to.default_content()
        print("   > Aguardando tela carregar para marcar Checkbox TTD...")
        time.sleep(3) 
        try:
            # Procura pela linha que tem "TTD 490" escrito
            robust_check_element(self.driver_lf, By.XPATH, XPATHS_LF["checkbox_industrializacao_smart"], wait_time=10)
            print("   > Checkbox Industrialização marcado com sucesso.")
        except Exception as e:
            print(f"   ⚠️ Não consegui marcar o checkbox TTD: {e}")
            self.warning_topmost("Aviso", "Verifique se o Checkbox TTD foi marcado.")

    def preencher_itens(self, itens_df, ordem_compra):
        wait_lf = WebDriverWait(self.driver_lf, 25)
        try: robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_adicionar_item"], wait_time=5)
        except: pass 

        for index, row in itens_df.iterrows():
            codigo_lf = str(row['cod_lf']) if pd.notna(row['cod_lf']) else ""
            qtd = row['qtd']
            preco = row['preco']

            while True: 
                if not codigo_lf:
                    codigo_lf = self.input_topmost("Código Manual", f"Pedido: {ordem_compra}\nQtd: {qtd}\n\nDigite Código LF:")
                    if not codigo_lf: 
                        if messagebox.askyesno("Pular?", "Sem código. Pular item?"): break
                        else: continue

                try:
                    robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_adicionar_item"])
                    robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["item_popup_iframe"]), (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                    
                    robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_sku_btn"])
                    self.abrir_popup_e_selecionar(codigo_lf, exact_match=False) 
                    
                    robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["item_popup_iframe"]), (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                    
                    robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_ordem_compra"], ordem_compra)
                    robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_qtd"], str(qtd).replace('.', ','))
                    robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_preco"], str(preco).replace('.', ','))
                    
                    robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_btn_salvar"])
                    self.driver_lf.switch_to.default_content()
                    wait_lf.until(EC.invisibility_of_element_located((By.XPATH, XPATHS_LF["item_popup_iframe"])))
                    break 

                except Exception as e:
                    try: self.driver_lf.find_element(By.CSS_SELECTOR, ".ui-dialog-titlebar-close").click()
                    except: pass
                    self.driver_lf.switch_to.default_content()
                    try: self.driver_lf.find_element(By.XPATH, "//button[@title='Close']").click()
                    except: pass
                    
                    codigo_lf = "" 
                    messagebox.showwarning("Erro Item", "Tente digitar o código manualmente.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoSchulzApp(root)
    root.mainloop()