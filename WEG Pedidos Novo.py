import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select 
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException, UnexpectedAlertPresentException 
import time
import re
import os 
import json
import glob
import pandas as pd
import base64 
import threading
from dotenv import load_dotenv 

try:
    import historico_rfq
except Exception as _e_hist:
    historico_rfq = None
    print(f"Aviso: módulo historico_rfq não carregado ({_e_hist}). Busca de códigos por RFQ desativada.")

# --- EXCEÇÃO CUSTOMIZADA PARA RECOMEÇAR ---
class RecomecarPedido(Exception):
    """Sinaliza ao robô para abandonar o pedido atual e voltar ao início."""
    pass

# --- 1. CONFIGURAÇÕES INICIAIS ---

load_dotenv()

URL_PORTAL_LF = "https://erp.bem.inf.br/ords/f?p=8000000:9999"
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
    pass 

MAPA_APLICACAO = {
    "Industrialização": "INDUSTRIALIZAÇÃO", 
    "Consumo": "USO E CONSUMO/IMOBILIZADO" 
}

# CNPJs da WEG com benefício Pró-Emprego (Isento de ICMS em Industrialização)
CNPJS_PRO_EMPREGO = ["07175725001050", "07175725000402"]

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
    "btn_anexar": "//button[span[normalize-space(.)='Novo']]",
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
    "popup_btn_fechar_X": "//button[.//i[contains(@class, 'icon-xmark')]] | //button[contains(@class, 'close')] | /html/body/div[4]/div[1]/div/div/header/div/div[1]/div/div/button/i",
    "popup_iframe": "//iframe[@id='document-details-iframe'] | //iframe[contains(@class, 'frame-in-modal')]", 
    "popup_iframe_check": "//*[@id='divItensDetalhado'] | //div[contains(@class, 'form-item')] | //*[contains(text(), 'CNPJ')]", 
    "popup_cnpj": "//*[contains(text(), '/000') and not(*[contains(text(), '/000')])] | //section[1]//p[1] | //p[contains(text(), '/000')] | //*[contains(text(), 'CNPJ')]/following-sibling::* | /html/body/main/form[1]/div/section[1]/div/div/div[1]/div[1]/p[1]",
    "popup_cond_pagto": "//*[(contains(text(), 'Dias') or contains(text(), 'dias')) and not(*[contains(text(), 'Dias') or contains(text(), 'dias')])] | //section[2]//div[5]//div | /html/body/main/form[1]/div/section[2]/div/div/div[5]/div/div",
    "popup_cif_fob": "//*[(contains(text(), 'CIF') or contains(text(), 'FOB')) and not(*[contains(text(), 'CIF') or contains(text(), 'FOB')])] | //section[2]//div[4]//span | /html/body/main/form[1]/div/section[2]/div/div/div[4]/div/div/span",
    "popup_valor_total": "//h1[contains(text(), 'R$')] | //main//h1 | //*[contains(text(), 'Valor Total')]/following-sibling::* | /html/body/main/form[1]/div/div[3]/div[2]/div[2]/h1",
    "linhas_itens": "//*[@id='divItensDetalhado']/table/tbody", 
    "item_numero_weg": ".//input[contains(@id, 'Item') and @type='hidden']", 
    "item_codigo": ".//div[contains(@class, 'form-item') and label[normalize-space(.)='Código']]//span",
    "item_preco": ".//div[contains(@class, 'form-item') and contains(., 'Preço Unitário')]//span",
    "item_qtd": ".//div[contains(@class, 'form-item') and contains(., 'Quantidade')]//span", 
    "item_aplicacao": ".//div[contains(@class, 'form-item') and contains(., 'Aplicação')]//span",
    "item_desc_span": ".//h2/span[2]",
    "item_deposito": ".//div[label[normalize-space(.)='Depósito Entrega']]//span",
    "item_ponto_descarga": ".//div[label[normalize-space(.)='Ponto de Descarga']]//span",
    "item_recebedor": ".//div[label[normalize-space(.)='Recebedor']]//span",
    "item_st": ".//div[label[contains(., 'Substituição Tributária')]]//span",
    "item_valor_st": ".//div[label[contains(., 'Valor da Substituição Tributária')]]//span",
    "btn_opcoes": "//button[contains(text(), 'Opções') or text()='Opções'] | //*[contains(@class, 'dropdown-toggle')]", 
    "btn_imprimir_menu": "//*[@id='MEComponentManager_MEButtonBar_1']//a[contains(., 'Imprimir')] | //a[@data-group='Consultas' and contains(., 'Imprimir')] | //a[contains(text(), 'Imprimir') or text()='Imprimir'] | /html/body/header/div/div/div/div/ul/li[3]/a", 
    "btn_imprimir_pagina": "//input[@value='Imprimir' and @class='noprint'] | //body",
    "popup_rfq": "//td[normalize-space(.)='RFQ:']/following-sibling::td"
}

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
        
        self.codigos_preenchidos_event = threading.Event()
        self.janela_codigos_pendentes = None
        self.lbl_status_robo_codigos = None
        self.screen_width = root_window.winfo_screenwidth()
        self.screen_height = root_window.winfo_screenheight()
        
        self.config = {
            "planilha_precos_path": "",
            "col_codigo_weg": "B",
            "col_codigo_lf": "D",
            "col_tipo": "E",
            "col_preco_pro_emprego": "I",
            "col_preco_ind_12": "K",
            "col_preco_consumo_17": "M"
        }
        self.tabela_precos = {}
        self.contrato_blumenau = {}
        self.mapa_blumenau_sap = {}
        self.dados_pedido_atual = {}
        self.itens_desconto_35 = []
        self.pedidos_nao_lidos = []
        self.transparencia_ativa = True
        self.alpha_padrao = 0.85
        self.carregar_config()
        
        self.centralizar_janela(300, 150)
        self.criar_tela_inicial()
        
    def normalizar_codigo(self, codigo):
        if not codigo:
            return ""
        codigo = str(codigo).strip()
        if codigo.endswith('.0'):
            codigo = codigo[:-2]
        cod_clean = codigo.lstrip('0')
        return cod_clean if cod_clean else codigo

    def _run_on_ui(self, func, *args, **kwargs):
        """Executa com segurança uma função na thread principal da interface gráfica."""
        if threading.current_thread() is threading.main_thread():
            return func(*args, **kwargs)
        event = threading.Event()
        res = [None]
        err = [None]
        def runner():
            try:
                res[0] = func(*args, **kwargs)
            except Exception as e:
                err[0] = e
            finally:
                event.set()
        self.root.after(0, runner)
        event.wait()
        if err[0]:
            raise err[0]
        return res[0]
        
    # --- FUNÇÕES ROBUSTAS COM INTERVENÇÃO MANUAL ---
    def dismiss_apex_alert_if_any(self, driver=None):
        """Fecha automaticamente alertas ou popups de erro de banco do ERP LF (ex: erro ORA-00904 BEG_VW_TRANSP_PREF_ID)."""
        d = driver if driver else getattr(self, 'driver_lf', None)
        if not d:
            return False
        dismissed = False

        # 1. Alerta nativo do navegador (window.alert)
        try:
            alert = d.switch_to.alert
            texto = alert.text
            alert.accept()
            print(f"   ℹ️ [ERP] Alerta nativo do navegador confirmado: '{texto}'")
            dismissed = True
        except Exception:
            pass

        # 2. Executa script JavaScript que varre o iframe atual e window.top (em caso de foco cruzado)
        try:
            js_script = """
            try {
                var dismissed = false;
                var docs = [document];
                try {
                    if (window.top && window.top.document && window.top.document !== document) {
                        docs.push(window.top.document);
                    }
                } catch(e) {}

                for (var i = 0; i < docs.length; i++) {
                    var doc = docs[i];
                    // Busca dialogs jQuery UI que NÃO sejam o popup LOV
                    var dialogs = doc.querySelectorAll(".ui-dialog:not(.ui-dialog-popuplov)");
                    for (var j = 0; j < dialogs.length; j++) {
                        var d = dialogs[j];
                        if (d.offsetParent !== null) {
                            var txt = (d.innerText || "").toLowerCase();
                            if (txt.indexOf("beg_vw_transp_pref_id") !== -1 || 
                                txt.indexOf("identificador") !== -1 || 
                                txt.indexOf("ora-") !== -1 || 
                                txt.indexOf("erro") !== -1 ||
                                txt.indexOf("transportadora") !== -1) {
                                var btns = d.querySelectorAll(".ui-dialog-buttonpane button, button.js-confirm, button");
                                for (var k = 0; k < btns.length; k++) {
                                    btns[k].click();
                                    dismissed = true;
                                    break;
                                }
                            }
                        }
                    }
                    // Busca mensagens do APEX (#apexPageMessage, .a-AlertMessage, etc)
                    var alerts = doc.querySelectorAll("#apexPageMessage, .a-AlertMessage, .t-Alert--error, .t-Alert");
                    for (var a = 0; a < alerts.length; a++) {
                        var el = alerts[a];
                        if (el.offsetParent !== null) {
                            var btns = el.querySelectorAll("button");
                            for (var b = 0; b < btns.length; b++) {
                                btns[b].click();
                                dismissed = true;
                            }
                        }
                    }
                }
                return dismissed;
            } catch(e) {
                return false;
            }
            """
            js_res = d.execute_script(js_script)
            if js_res:
                print("   ℹ️ [ERP] Popup de aviso do ERP fechado automaticamente via JS.")
                dismissed = True
        except Exception:
            pass

        # 3. Varredura via Selenium caso o JS não tenha pego
        try:
            dialogs = d.find_elements(By.XPATH, "//div[contains(@class,'ui-dialog') and not(contains(@class,'ui-dialog-popuplov')) and not(contains(@style,'display: none'))]")
            for dlg in dialogs:
                try:
                    txt = dlg.text.lower()
                    if any(k in txt for k in ["beg_vw_transp_pref_id", "identificador", "ora-", "transportadora"]):
                        botoes = dlg.find_elements(By.XPATH, ".//div[contains(@class,'ui-dialog-buttonpane')]//button | .//button")
                        for b in botoes:
                            if b.is_displayed():
                                d.execute_script("arguments[0].click();", b)
                                print(f"   ℹ️ [ERP] Popup de aviso do ERP confirmado via Selenium: '{dlg.text.strip()[:60]}...'")
                                dismissed = True
                                break
                except Exception:
                    pass
        except Exception:
            pass

        return dismissed

    def wait_for_apex_ready(self, driver=None, timeout=10):
        """Aguarda o Oracle APEX e jQuery terminarem todas as requisições AJAX e animações."""
        d = driver if driver else getattr(self, 'driver_lf', None)
        if not d:
            return
        self.dismiss_apex_alert_if_any(d)
        t_end = time.time() + timeout
        time.sleep(0.2)
        while time.time() < t_end:
            self.dismiss_apex_alert_if_any(d)
            try:
                is_ajax_active = d.execute_script("return (typeof jQuery !== 'undefined') ? (jQuery.active > 0) : false;")
                has_spinner = d.execute_script("""
                    return document.querySelectorAll('.u-Processing, .apex_wait_overlay, .a-LoadingIndicator').length > 0;
                """)
                if not is_ajax_active and not has_spinner:
                    break
            except Exception:
                pass
            time.sleep(0.15)
        time.sleep(0.15)
        self.dismiss_apex_alert_if_any(d)

    def robust_wait_for(self, driver, by, locator, condition=EC.presence_of_element_located, wait_time=15):
        self.dismiss_apex_alert_if_any(driver)
        wait = WebDriverWait(driver, wait_time)
        try:
            return wait.until(condition((by, locator)))
        except Exception as e:
            self.dismiss_apex_alert_if_any(driver)
            resp = self.confirm_topmost(
                "Intervenção Manual Necessária",
                f"O robô não conseguiu encontrar/aguardar o elemento:\n{locator}\n\n"
                "Por favor, realize a ação manualmente no navegador e clique em 'Sim' para continuar.\n"
                "Ou clique em 'Não' para abortar o pedido."
            )
            if resp:
                try: return driver.find_element(by, locator)
                except: return None
            else:
                raise Exception(f"Ação cancelada pelo usuário. {e}")

    def robust_click(self, driver, by, locator, wait_time=15):
        self.dismiss_apex_alert_if_any(driver)
        try:
            el = self.robust_wait_for(driver, by, locator, EC.presence_of_element_located, wait_time)
            if el:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
                time.sleep(0.3)
                el = self.robust_wait_for(driver, by, locator, EC.element_to_be_clickable, 5)
                if el: el.click()
            return
        except Exception:
            try:
                el = driver.find_element(by, locator)
                driver.execute_script("arguments[0].click();", el)
                return
            except Exception as e:
                resp = self.confirm_topmost(
                    "Intervenção Manual Necessária",
                    f"Não foi possível clicar em:\n{locator}\n\n"
                    "Por favor, clique manualmente no navegador e aperte 'Sim' para continuar.\n"
                    "Ou clique em 'Não' para cancelar."
                )
                if resp: return
                else: raise Exception(f"Cancelado pelo usuário ao falhar o clique. {e}")

    def robust_find_and_send_keys(self, driver, by, locator, keys, wait_time=15):
        try:
            el = self.robust_wait_for(driver, by, locator, EC.visibility_of_element_located, wait_time)
            if el:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", el)
                time.sleep(0.3)
                el.clear()
                el.send_keys(keys)
        except Exception as e:
            resp = self.confirm_topmost(
                "Intervenção Manual",
                f"Falha ao digitar '{keys}' em:\n{locator}\n\n"
                "Preencha manualmente e clique 'Sim' para continuar."
            )
            if not resp: raise Exception(f"Ação manual cancelada. {e}")

    def robust_switch_to_iframe(self, driver, iframe_tuple, check_tuple, wait_time=15):
        # Verifica se já está com foco no iframe correto
        try:
            if driver.find_elements(check_tuple[0], check_tuple[1]):
                return
        except Exception:
            pass

        driver.switch_to.default_content()
        time.sleep(1.0)
        try:
            wait = WebDriverWait(driver, wait_time)
            wait.until(EC.frame_to_be_available_and_switch_to_it((iframe_tuple[0], iframe_tuple[1])))
            return
        except Exception:
            # Fallback buscando todos os iframes da página
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for f in iframes:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(f)
                    if driver.find_elements(check_tuple[0], check_tuple[1]) or driver.find_elements(By.XPATH, "//button[contains(., 'Opções')]"):
                        return
            except Exception:
                pass

            self.force_top_window()
            resp = messagebox.askyesno(
                "Intervenção Manual",
                "Não foi possível focar no pop-up (iframe).\n"
                "Verifique se a tela carregou. Clique em 'Sim' para tentar prosseguir."
            )
            if not resp: raise Exception("Falha de iframe.")

    def robust_check_checkbox(self, driver, by, locator, wait_time=15):
        try:
            el = self.robust_wait_for(driver, by, locator, EC.presence_of_element_located, wait_time)
            if el and not el.is_selected():
                driver.execute_script("arguments[0].click();", el)
        except Exception:
            pass # Ignora erro de checkbox silenciosamente, ou loga

    def centralizar_janela(self, w=380, h=350):
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def posicionar_painel_embaixo(self, w=270, h=270):
        """Posiciona a janela de controle centralizada na horizontal, mas bem embaixo na tela,
        rente acima da barra de tarefas, liberando a visão das duas telas do navegador no fundo."""
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = int((ws - w) / 2)
        y = max(10, int(hs - h - 65))
        self.root.geometry(f'{w}x{h}+{x}+{y}')

    def criar_tela_inicial(self):
        for widget in self.root.winfo_children():
            widget.destroy()
            
        self.root.configure(bg="#F8FAFC")
        self.root.attributes('-alpha', 1.0)
        self.centralizar_janela(380, 360)

        # Header Frame
        header = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=15, pady=12)
        header.pack(fill=tk.X, padx=12, pady=(12, 10))

        tk.Label(header, text="⚡ Automação WEG Pedidos", font=("Segoe UI", 14, "bold"), fg="#0F172A", bg="#FFFFFF").pack(anchor="w")
        
        lbl_v = tk.Label(header, text="✅ Atualizado em 03/09/2026 às 11:20 | Regras CNPJ & ICMS", font=("Segoe UI", 8, "bold"), fg="#15803D", bg="#DCFCE7", padx=6, pady=2)
        lbl_v.pack(anchor="w", pady=(3, 0))

        # Body Container
        body = tk.Frame(self.root, bg="#F8FAFC", padx=15)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text="Selecione o portal de origem:", font=("Segoe UI", 10, "bold"), fg="#334155", bg="#F8FAFC").pack(anchor="w", pady=(5, 10))

        btn_jve = tk.Button(body, text="🏢 Pedidos WEG Joinville (JVE)", font=("Segoe UI", 10, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", relief="flat", pady=8, cursor="hand2", command=lambda: self.iniciar_automacao("JVE"))
        btn_jve.pack(fill=tk.X, pady=4)

        btn_di = tk.Button(body, text="🏢 Pedidos WEG Dois Irmãos (DI)", font=("Segoe UI", 10, "bold"), bg="#0284C7", fg="white", activebackground="#0369A1", activeforeground="white", relief="flat", pady=8, cursor="hand2", command=lambda: self.iniciar_automacao("DI"))
        btn_di.pack(fill=tk.X, pady=4)

        btn_config = tk.Button(body, text="⚙️ Configurações (Planilha de Preços)", font=("Segoe UI", 9, "bold"), bg="#E2E8F0", fg="#334155", activebackground="#CBD5E1", relief="flat", pady=6, cursor="hand2", command=self.abrir_configuracoes)
        btn_config.pack(fill=tk.X, pady=(12, 4))

    # --- GESTÃO DE FILA DE PEDIDOS NÃO LIDOS WEG ---
    def extrair_pedidos_nao_lidos_portal(self):
        """Varre a tabela de transações do Mercado Eletrônico de baixo para cima (do mais antigo para o mais novo) e extrai os números dos pedidos com status 'Não lida'."""
        if not getattr(self, 'driver_weg', None):
            return []
        
        pedidos = []
        try:
            self.driver_weg.switch_to.default_content()
            js_code = """
            var unreadList = [];
            var rows = document.querySelectorAll("table tbody tr");
            // Varre de baixo para cima (do mais antigo para o mais novo)
            for (var i = rows.length - 1; i >= 0; i--) {
                var row = rows[i];
                var text = (row.innerText || "").toLowerCase();
                if (text.indexOf("não lida") === -1 && text.indexOf("nao lida") === -1 &&
                    text.indexOf("não lido") === -1 && text.indexOf("nao lido") === -1) {
                    continue;
                }
                var foundNum = null;
                // Coluna 2 é estritamente o "Número do Pedido" WEG (ex: 4519264337)
                // A coluna 1 é "Serial ME" (ex: 64643483) e DEVE ser ignorada!
                var col2 = row.querySelector("td[aria-colindex='2']") || row.querySelectorAll("td")[1];
                if (col2) {
                    var m = (col2.innerText || "").match(/\b(45\d{8})\b/);
                    if (m) foundNum = m[1];
                }
                // Fallback na linha: procura estritamente por número de 10 dígitos começando com 45
                if (!foundNum) {
                    var m2 = (row.innerText || "").match(/\b(45\d{8})\b/);
                    if (m2) foundNum = m2[1];
                }
                if (foundNum && unreadList.indexOf(foundNum) === -1) {
                    unreadList.push(foundNum);
                }
            }
            return unreadList;
            """
            res = self.driver_weg.execute_script(js_code)
            if res and isinstance(res, list):
                pedidos = [str(p).strip() for p in res if str(p).strip()]
        except Exception as e:
            print(f"Aviso ao extrair pedidos não lidos via JS: {e}")

        # Fallback via Selenium (varre de baixo para cima)
        if not pedidos:
            try:
                self.driver_weg.switch_to.default_content()
                rows = self.driver_weg.find_elements(By.XPATH, "//table//tbody/tr")
                for r in reversed(rows):
                    try:
                        txt = r.text.lower()
                        if any(k in txt for k in ["não lida", "nao lida", "não lido", "nao lido"]):
                            cols = r.find_elements(By.XPATH, "./td")
                            foundNum = None
                            if len(cols) >= 2:
                                m = re.search(r'\b(45\d{8})\b', cols[1].text)
                                if m:
                                    foundNum = m.group(1)
                            if not foundNum:
                                m = re.search(r'\b(45\d{8})\b', r.text)
                                if m:
                                    foundNum = m.group(1)
                            if foundNum and foundNum not in pedidos:
                                pedidos.append(foundNum)
                    except Exception:
                        pass
            except Exception:
                pass

        return pedidos

    def atualizar_fila_pedidos(self, dar_f5=False):
        """Atualiza a lista interna de pedidos 'Não lido' e sincroniza a interface."""
        if not getattr(self, 'driver_weg', None):
            return
        try:
            try:
                janela_princ = self.driver_weg.window_handles[0]
                self.driver_weg.switch_to.window(janela_princ)
                self.driver_weg.switch_to.default_content()
            except Exception:
                pass

            if dar_f5:
                # Se houver pop-up aberto da WEG, tenta fechar antes do refresh
                try:
                    btns_fechar = self.driver_weg.find_elements(By.XPATH, XPATHS_WEG["popup_btn_fechar_X"])
                    for bf in btns_fechar:
                        if bf.is_displayed():
                            self.driver_weg.execute_script("arguments[0].click();", bf)
                            break
                except Exception:
                    pass

                print("Atualizando página da WEG (F5)...")
                self.driver_weg.refresh()
                try:
                    WebDriverWait(self.driver_weg, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//table//tbody/tr"))
                    )
                except Exception:
                    pass
                time.sleep(0.4)

            pedidos = self.extrair_pedidos_nao_lidos_portal()
            self.pedidos_nao_lidos = pedidos
            print(f"Lista de pedidos 'Não lido' encontrada: {pedidos}")
        except Exception as e:
            print(f"Aviso ao buscar pedidos não lidos: {e}")
        finally:
            self.root.after(0, self._sincronizar_ui_fila_pedidos)

    def atualizar_fila_pedidos_async(self):
        """Dispara a atualização com F5 em thread separada para não travar a interface."""
        if hasattr(self, 'btn_atualizar_fila'):
            self.btn_atualizar_fila.config(text="⏳ Buscando...", state=tk.DISABLED)
        def _task():
            self.atualizar_fila_pedidos(dar_f5=True)
            self.root.after(0, lambda: getattr(self, 'btn_atualizar_fila', None) and self.btn_atualizar_fila.config(text="🔄 Atualizar", state=tk.NORMAL))
        threading.Thread(target=_task, daemon=True).start()

    def _sincronizar_ui_fila_pedidos(self):
        """Atualiza a caixa de texto, rótulos e os botões da interface conforme a fila pendente."""
        pedidos = getattr(self, 'pedidos_nao_lidos', [])
        if pedidos:
            prox = pedidos[0]
            # Preenche na caixa o próximo sugerido da fila (o mais antigo)
            self.entry_pedido.delete(0, tk.END)
            self.entry_pedido.insert(0, prox)
            self.entry_pedido.select_range(0, tk.END)

            total = len(pedidos)
            if hasattr(self, 'lbl_fila_pedidos'):
                self.lbl_fila_pedidos.config(text=f"📬 Não lidos: 1 de {total} (mais antigo)", fg="#0369A1")
            if hasattr(self, 'btn_proximo'):
                self.btn_proximo.config(text=f"🚀 Fazer Próximo da Fila ({prox})", bg="#16A34A", state=tk.NORMAL)
            if hasattr(self, 'btn_digitado'):
                self.btn_digitado.config(text="▶️ Fazer Pedido da Caixa", state=tk.NORMAL)

            if hasattr(self, 'lbl_seq_pedidos'):
                if total > 1:
                    seq_str = ", ".join(pedidos[:3]) + ("..." if total > 3 else "")
                    self.lbl_seq_pedidos.config(text=f"Fila: {seq_str}", fg="#64748B")
                else:
                    self.lbl_seq_pedidos.config(text=f"Fila: {prox} (último não lido)", fg="#64748B")
        else:
            self.entry_pedido.delete(0, tk.END)
            if hasattr(self, 'lbl_fila_pedidos'):
                self.lbl_fila_pedidos.config(text="📋 Nº do Pedido WEG:", fg="#1E293B")
            if hasattr(self, 'btn_proximo'):
                self.btn_proximo.config(text="🚀 Sem pedidos na fila", bg="#94A3B8", state=tk.DISABLED)
            if hasattr(self, 'btn_digitado'):
                self.btn_digitado.config(text="▶️ Fazer Pedido Digitado", state=tk.NORMAL)
            if hasattr(self, 'lbl_seq_pedidos'):
                self.lbl_seq_pedidos.config(text="✅ Nenhum pedido 'Não lido'", fg="#166534")

    # --- PAINEL DE CONTROLE COMPACTO, INFERIOR E TRANSLÚCIDO ---
    def criar_painel_controle(self):
        self.root.deiconify()
        self.root.configure(bg="#F8FAFC")
        self.posicionar_painel_embaixo(270, 275)
        
        # Aplica a transparência suave (85%) para enxergar o conteúdo atrás
        if getattr(self, 'transparencia_ativa', True):
            self.root.attributes('-alpha', self.alpha_padrao)
        else:
            self.root.attributes('-alpha', 1.0)
        
        for widget in self.root.winfo_children():
            widget.destroy()

        # Efeito de hover suave: ao passar o mouse fica 98% nítido, ao sair volta para 85%
        def on_enter(event):
            if getattr(self, 'transparencia_ativa', False):
                self.root.attributes('-alpha', 0.98)

        def on_leave(event):
            if getattr(self, 'transparencia_ativa', False):
                x = self.root.winfo_pointerx() - self.root.winfo_rootx()
                y = self.root.winfo_pointery() - self.root.winfo_rooty()
                if x < 0 or y < 0 or x >= self.root.winfo_width() or y >= self.root.winfo_height():
                    self.root.attributes('-alpha', self.alpha_padrao)

        self.root.bind("<Enter>", on_enter)
        self.root.bind("<Leave>", on_leave)

        # Header Frame compacto
        header = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=8, pady=4)
        header.pack(fill=tk.X, padx=8, pady=(6, 2))

        top_row = tk.Frame(header, bg="#FFFFFF")
        top_row.pack(fill=tk.X)
        tk.Label(top_row, text="⚡ WEG Pedidos", font=("Segoe UI", 10, "bold"), fg="#0F172A", bg="#FFFFFF").pack(side="left")

        nome_local = "Joinville" if self.local_weg_escolhido == "JVE" else "Dois Irmãos"
        lbl_portal = tk.Label(top_row, text=f"🟢 {nome_local}", font=("Segoe UI", 8, "bold"), fg="#15803D", bg="#DCFCE7", padx=6, pady=1)
        lbl_portal.pack(side="right")

        # Body Frame estreito e enxuto
        body = tk.Frame(self.root, bg="#F8FAFC", padx=8)
        body.pack(fill=tk.BOTH, expand=True)

        status_frame = tk.Frame(body, bg="#F8FAFC")
        status_frame.pack(fill=tk.X, pady=(2, 1))

        self.lbl_fila_pedidos = tk.Label(status_frame, text="📋 Nº do Pedido WEG:", font=("Segoe UI", 8, "bold"), fg="#1E293B", bg="#F8FAFC")
        self.lbl_fila_pedidos.pack(side="left")

        self.btn_atualizar_fila = tk.Button(status_frame, text="🔄 Atualizar", font=("Segoe UI", 7, "bold"), bg="#E0F2FE", fg="#0369A1", activebackground="#BAE6FD", relief="flat", padx=4, pady=0, cursor="hand2", command=self.atualizar_fila_pedidos_async)
        self.btn_atualizar_fila.pack(side="right")

        self.entry_pedido = tk.Entry(body, font=("Segoe UI", 12, "bold"), justify="center", bg="#FFFFFF", fg="#0F172A", relief="solid", bd=1)
        self.entry_pedido.pack(fill=tk.X, pady=(1, 1), ipady=2)
        self.entry_pedido.focus()
        self.entry_pedido.bind("<Return>", lambda event: self.processar_pedido())

        self.lbl_seq_pedidos = tk.Label(body, text="", font=("Segoe UI", 7), fg="#64748B", bg="#F8FAFC")
        self.lbl_seq_pedidos.pack(anchor="w", pady=(0, 2))

        # Botão 1: Processar próximo da sugestão da fila (verde)
        self.btn_proximo = tk.Button(body, text="🚀 Fazer Próximo da Fila", bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), activebackground="#15803D", activeforeground="white", relief="flat", pady=3, cursor="hand2", command=self.processar_proximo_fila)
        self.btn_proximo.pack(fill=tk.X, pady=1)

        # Botão 2: Processar o que estiver digitado/colocado na caixa (azul)
        self.btn_digitado = tk.Button(body, text="▶️ Fazer Pedido da Caixa", bg="#2563EB", fg="white", font=("Segoe UI", 8, "bold"), activebackground="#1D4ED8", activeforeground="white", relief="flat", pady=2, cursor="hand2", command=self.processar_pedido_digitado)
        self.btn_digitado.pack(fill=tk.X, pady=1)

        # Ações secundárias lado a lado
        sec_row = tk.Frame(body, bg="#F8FAFC")
        sec_row.pack(fill=tk.X, pady=2)

        self.btn_recomecar = tk.Button(sec_row, text="🔄 Destravar LF", bg="#D97706", fg="white", font=("Segoe UI", 8, "bold"), activebackground="#B45309", activeforeground="white", relief="flat", pady=2, cursor="hand2", command=self.recomecar_lf)
        self.btn_recomecar.pack(side="left", fill=tk.X, expand=True, padx=(0, 2))
        
        outro_local = "DI" if self.local_weg_escolhido == "JVE" else "JVE"
        nome_outro = "Dois Irmãos" if outro_local == "DI" else "Joinville"
        self.btn_trocar = tk.Button(sec_row, text=f"🔁 Mudar Portal", bg="#475569", fg="white", font=("Segoe UI", 8, "bold"), activebackground="#334155", activeforeground="white", relief="flat", pady=2, cursor="hand2", command=lambda: self.trocar_portal_weg(outro_local))
        self.btn_trocar.pack(side="right", fill=tk.X, expand=True, padx=(2, 0))

        btn_row = tk.Frame(body, bg="#F8FAFC")
        btn_row.pack(fill=tk.X, pady=(3, 2))

        # Botão para alternar transparência se o usuário quiser 100% ou 85%
        def alternar_transparencia():
            self.transparencia_ativa = not self.transparencia_ativa
            if self.transparencia_ativa:
                self.root.attributes('-alpha', self.alpha_padrao)
                btn_transp.config(text="👻 85%", bg="#DCFCE7", fg="#166534")
            else:
                self.root.attributes('-alpha', 1.0)
                btn_transp.config(text="👁️ 100%", bg="#E2E8F0", fg="#334155")

        btn_transp = tk.Button(btn_row, text="👻 85%", font=("Segoe UI", 8, "bold"), bg="#DCFCE7", fg="#166534", relief="flat", pady=2, cursor="hand2", command=alternar_transparencia)
        btn_transp.pack(side="left", fill=tk.X, expand=True, padx=(0, 2))

        self.btn_config = tk.Button(btn_row, text="⚙️ Config", font=("Segoe UI", 8), bg="#E2E8F0", fg="#334155", relief="flat", pady=2, cursor="hand2", command=self.abrir_configuracoes)
        self.btn_config.pack(side="left", fill=tk.X, expand=True, padx=(0, 2))

        self.btn_encerrar = tk.Button(btn_row, text="❌ Sair", font=("Segoe UI", 8), bg="#FEE2E2", fg="#991B1B", relief="flat", pady=2, cursor="hand2", command=self.fechar_tudo)
        self.btn_encerrar.pack(side="right", fill=tk.X, expand=True, padx=(2, 0))

        # Sincroniza a interface com a fila de pedidos pendentes
        self._sincronizar_ui_fila_pedidos()

    def trocar_portal_weg(self, novo_local):
        """Fecha o portal WEG atual e loga no portal escolhido, mantendo o LF."""
        try:
            if hasattr(self, 'btn_proximo'):
                self.btn_proximo.config(state=tk.DISABLED)
            if hasattr(self, 'btn_digitado'):
                self.btn_digitado.config(state=tk.DISABLED)
            self.btn_recomecar.config(state=tk.DISABLED)
            self.btn_trocar.config(state=tk.DISABLED, text="Trocando de portal...")
            self.root.update()

            # Fecha o navegador da WEG atual
            if self.driver_weg:
                try:
                    self.driver_weg.quit()
                except Exception:
                    pass

            self.local_weg_escolhido = novo_local
            creds_weg = CREDENCIAIS["WEG_JVE"] if self.local_weg_escolhido == "JVE" else CREDENCIAIS["WEG_DI"]
            url_weg = URL_PORTAL_WEG_JVE if self.local_weg_escolhido == "JVE" else URL_PORTAL_WEG_DI

            # Inicializa o navegador da nova WEG
            chrome_options = webdriver.ChromeOptions()
            chrome_options.page_load_strategy = 'eager'
            chrome_options.add_argument("--disable-background-timer-throttling")
            chrome_options.add_argument("--disable-backgrounding-occluded-windows")
            chrome_options.add_argument("--disable-renderer-backgrounding")
            chrome_options.add_argument("--disable-features=CalculateWindowOcclusion")
            chrome_options.add_argument("--kiosk-printing")
            prefs = {
                "download.default_directory": self.download_dir, 
                "download.prompt_for_download": False, 
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
                "savefile.default_directory": self.download_dir
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            self.driver_weg = webdriver.Chrome(options=chrome_options)
            try:
                self.driver_weg.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(window, 'print', { value: function(){ console.log('Print preview suprimido'); }, writable: false, configurable: false });"
                })
            except Exception:
                pass
            
            try:
                self.driver_weg.get("chrome://settings/")
                self.driver_weg.execute_script("chrome.settingsPrivate.setDefaultZoom(0.9);")
            except Exception:
                pass
            
            self.driver_weg.get(url_weg)
            self.robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["login_user"], creds_weg["user"])
            self.robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["login_pass"], creds_weg["pass"])
            self.robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["login_btn"])
            
            time.sleep(2)
            self.organizar_janelas_lado_a_lado()
            self.criar_painel_controle() # Atualiza os textos do painel
            self.info_topmost("Sucesso", f"Portal trocado para WEG {self.local_weg_escolhido} com sucesso!")

        except Exception as e:
            self.error_topmost("Erro", f"Ocorreu um erro ao trocar o portal: {e}")
            self.fechar_tudo()

    def recomecar_lf(self):
        """Dá um refresh na tela do LF para limpar pop-ups travados"""
        if self.driver_lf:
            try:
                self.driver_lf.get(URL_PORTAL_LF)
                self.driver_lf.switch_to.default_content()
                self.info_topmost("Destravado", "A página da LF foi recarregada. Pode tentar o pedido novamente.")
            except Exception as e:
                self.error_topmost("Erro", f"Não foi possível destravar: {e}")

    # --- CONFIGURAÇÕES DE PREÇO E ARQUIVOS ---
    def carregar_config(self):
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
        cfg_candidates = [
            "config_weg_lf.json",
            os.path.join(script_dir, "config_weg_lf.json")
        ]
        cfg_file = next((f for f in cfg_candidates if os.path.exists(f)), None)
        if cfg_file:
            try:
                with open(cfg_file, "r", encoding="utf-8") as f:
                    self.config.update(json.load(f))
                path = self.config.get("planilha_precos_path", "")
                if path and not os.path.exists(path) and os.path.exists(os.path.join(script_dir, path)):
                    path = os.path.join(script_dir, path)
                    self.config["planilha_precos_path"] = path
                if path and os.path.exists(path):
                    self.carregar_precos_excel(path, silencioso=True)
            except Exception: pass
            
        # Carregar Contrato Kanban Blumenau
        blu_candidates = [
            "contrato_weg_blumenau.json",
            os.path.join(script_dir, "contrato_weg_blumenau.json"),
            os.path.join(os.path.dirname(script_dir), "contrato_weg_blumenau.json")
        ]
        blu_file = next((f for f in blu_candidates if os.path.exists(f)), None)
        if blu_file:
            try:
                with open(blu_file, "r", encoding="utf-8") as bf:
                    self.contrato_blumenau = json.load(bf)
                self.mapa_blumenau_sap = {}
                for it in self.contrato_blumenau.get("itens", []):
                    c_sap = self.normalizar_codigo(it.get("codigo_sap", ""))
                    if c_sap:
                        self.mapa_blumenau_sap[c_sap] = it
                print(f"   ℹ️ [CONTRATO BLUMENAU] Carregado com sucesso: {len(self.mapa_blumenau_sap)} itens mapeados.")
            except Exception as e:
                print(f"   ⚠️ Erro ao carregar contrato Blumenau: {e}")

        if not self.config.get("planilha_precos_path") or not os.path.exists(self.config.get("planilha_precos_path", "")):
            try:
                search_dirs = [".", script_dir]
                for d in search_dirs:
                    if not os.path.exists(d): continue
                    arquivos_xlsx = [os.path.join(d, f) for f in os.listdir(d) if f.endswith(".xlsx") and not f.startswith("~$")]
                    if arquivos_xlsx:
                        genesis_files = [f for f in arquivos_xlsx if "genesis" in os.path.basename(f).lower() or "arthur" in os.path.basename(f).lower()]
                        target_path = os.path.abspath(genesis_files[0] if genesis_files else arquivos_xlsx[0])
                        self.config["planilha_precos_path"] = target_path
                        self.carregar_precos_excel(target_path, silencioso=True)
                        break
            except Exception: pass

    def salvar_config(self):
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
        cfg_path = os.path.join(script_dir, "config_weg_lf.json")
        try:
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
        except Exception:
            with open("config_weg_lf.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)

    def carregar_precos_excel(self, path, silencioso=False):
        try:
            try:
                df = pd.read_excel(path, dtype=str, header=None)
            except Exception:
                try:
                    df = pd.read_excel(path, dtype=str, header=None, engine='openpyxl')
                except Exception:
                    try:
                        df = pd.read_excel(path, dtype=str, header=None, engine='xlrd')
                    except Exception:
                        df = pd.read_csv(path, dtype=str, header=None)
            self.tabela_precos = {}
            if df.empty: return
            
            # Limpeza rápida de linhas nulas/vazias
            df.dropna(how='all', inplace=True)

            col_w = self.config.get("col_codigo_weg", "B")
            col_l = self.config.get("col_codigo_lf", "D")
            col_tipo = self.config.get("col_tipo", "E")
            col_p_pro = self.config.get("col_preco_pro_emprego", "I")
            col_p_ind12 = self.config.get("col_preco_ind_12", "K")
            col_p_cons17 = self.config.get("col_preco_consumo_17", "M")

            first_row = df.iloc[0].values
            
            def get_col_index(c_str):
                c_str = str(c_str).strip()
                if len(c_str) <= 2 and c_str.isalpha() and c_str.upper() == c_str:
                    idx = 0
                    for char in c_str.upper():
                        idx = idx * 26 + (ord(char) - ord('A') + 1)
                    return idx - 1
                if c_str.isdigit(): return int(c_str)
                for i, val in enumerate(first_row):
                    if str(val).strip().lower() == c_str.lower(): return i
                return None

            idx_w = get_col_index(col_w)
            idx_l = get_col_index(col_l)
            idx_tipo = get_col_index(col_tipo)
            idx_p_pro = get_col_index(col_p_pro)
            idx_p_ind12 = get_col_index(col_p_ind12)
            idx_p_cons17 = get_col_index(col_p_cons17)

            def limpar_preco(val_str):
                if not val_str or str(val_str).lower() == 'nan': return None
                try:
                    val_clean = re.sub(r'[^\d,.]', '', str(val_str))
                    if not val_clean: return None
                    if ',' in val_clean and '.' in val_clean:
                        val_clean = val_clean.replace('.', '').replace(',', '.')
                    elif ',' in val_clean:
                        val_clean = val_clean.replace(',', '.')
                    num = float(val_clean)
                    if num <= 0: return None
                    return f"{num:.4f}".replace('.', ',')
                except Exception:
                    return None

            for row in df.itertuples(index=False):
                try:
                    cod_w = str(row[idx_w]).strip() if idx_w is not None and idx_w < len(row) else ""
                    cod_l = str(row[idx_l]).strip() if idx_l is not None and idx_l < len(row) else ""
                    tipo_str = str(row[idx_tipo]).strip().lower() if idx_tipo is not None and idx_tipo < len(row) else ""
                    
                    if not cod_w or cod_w.lower() == 'nan' or cod_w == str(first_row[idx_w] if idx_w is not None else ""):
                        continue
                    if cod_w.endswith('.0'): cod_w = cod_w[:-2]
                    if cod_l.endswith('.0'): cod_l = cod_l[:-2]
                    if cod_w == '0' or cod_w == 'nan': continue
                    
                    p_pro = limpar_preco(row[idx_p_pro] if idx_p_pro is not None and idx_p_pro < len(row) else None)
                    p_ind12 = limpar_preco(row[idx_p_ind12] if idx_p_ind12 is not None and idx_p_ind12 < len(row) else None)
                    p_cons17 = limpar_preco(row[idx_p_cons17] if idx_p_cons17 is not None and idx_p_cons17 < len(row) else None)
                    
                    cod_w_norm = self.normalizar_codigo(cod_w)
                    if not cod_w_norm: continue
                    
                    if cod_w_norm not in self.tabela_precos:
                        self.tabela_precos[cod_w_norm] = {
                            "industrializacao": {"codigo_lf": None, "preco_pro_emprego": None, "preco_ind_12": None},
                            "consumo": {"codigo_lf": None, "preco_consumo": None},
                            "geral": {"codigo_lf": None, "preco_pro_emprego": None, "preco_ind_12": None, "preco_consumo": None}
                        }
                        
                    entry = self.tabela_precos[cod_w_norm]
                    cod_lf_val = cod_l if cod_l and cod_l.lower() != 'nan' and cod_l != '0' else None
                    
                    is_tipo_consumo = "consumo" in tipo_str
                    is_tipo_ind = "industrializa" in tipo_str
                    
                    if is_tipo_consumo:
                        if cod_lf_val: entry["consumo"]["codigo_lf"] = cod_lf_val
                        if p_cons17: entry["consumo"]["preco_consumo"] = p_cons17
                        elif p_pro: entry["consumo"]["preco_consumo"] = p_pro
                    elif is_tipo_ind:
                        if cod_lf_val: entry["industrializacao"]["codigo_lf"] = cod_lf_val
                        if p_pro: entry["industrializacao"]["preco_pro_emprego"] = p_pro
                        if p_ind12: entry["industrializacao"]["preco_ind_12"] = p_ind12
                        if p_cons17: entry["consumo"]["preco_consumo"] = p_cons17
                        if cod_lf_val and not entry["consumo"]["codigo_lf"]: entry["consumo"]["codigo_lf"] = cod_lf_val
                    else:
                        if cod_lf_val: entry["industrializacao"]["codigo_lf"] = cod_lf_val
                        if p_pro: entry["industrializacao"]["preco_pro_emprego"] = p_pro
                        if p_ind12: entry["industrializacao"]["preco_ind_12"] = p_ind12
                        if p_cons17: entry["consumo"]["preco_consumo"] = p_cons17
                    
                    # Atualiza também no geral como fallback
                    if cod_lf_val and not entry["geral"]["codigo_lf"]: entry["geral"]["codigo_lf"] = cod_lf_val
                    if p_pro and not entry["geral"]["preco_pro_emprego"]: entry["geral"]["preco_pro_emprego"] = p_pro
                    if p_ind12 and not entry["geral"]["preco_ind_12"]: entry["geral"]["preco_ind_12"] = p_ind12
                    if p_cons17 and not entry["geral"]["preco_consumo"]: entry["geral"]["preco_consumo"] = p_cons17
                    
                except Exception:
                    continue
                    
            if not silencioso:
                self.info_topmost("Sucesso", f"Tabela de preços carregada ({len(self.tabela_precos)} itens mapeados).")
        except Exception as e:
            if not silencioso: self.error_topmost("Erro", f"Erro ao ler a planilha: {e}")

    def abrir_configuracoes(self):
        top = tk.Toplevel(self.root)
        top.title("Configurações")
        top.geometry("400x250")
        top.attributes('-topmost', True)
        
        tk.Label(top, text="Planilha de Preços (Excel):").pack(pady=5)
        lbl_path = tk.Label(top, text=self.config.get("planilha_precos_path", "Nenhuma"), fg="blue", wraplength=380)
        lbl_path.pack()
        
        def escolher_planilha():
            path = filedialog.askopenfilename(title="Selecione", filetypes=[("Excel", "*.xlsx *.xls")])
            if path:
                self.config["planilha_precos_path"] = path
                lbl_path.config(text=path)
                self.carregar_precos_excel(path)
                self.salvar_config()
        
        tk.Button(top, text="Selecionar Arquivo", command=escolher_planilha).pack(pady=5)
        
        tk.Label(top, text="Colunas na planilha (Ex: Código e Preço):").pack(pady=5)
        tk.Label(top, text="Para alterar as colunas, edite o arquivo config_weg_lf.json").pack()
        
        tk.Button(top, text="Fechar", command=top.destroy).pack(pady=10)

    # --- HELPERS DE INTERFACE ---
    def force_top_window(self):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        if getattr(self, 'transparencia_ativa', False):
            self.root.attributes('-alpha', self.alpha_padrao)
        self.root.lift()
        self.root.focus_force() 

    def input_topmost(self, title, prompt):
        def _inp():
            self.force_top_window()
            valor = simpledialog.askstring(title, prompt, parent=self.root)
            return valor
        return self._run_on_ui(_inp)

    def info_topmost(self, title, message):
        def _info():
            self.force_top_window()
            messagebox.showinfo(title, message, parent=self.root)
        self._run_on_ui(_info)

    def error_topmost(self, title, message):
        def _err():
            self.force_top_window()
            messagebox.showerror(title, message, parent=self.root)
        self._run_on_ui(_err)

    def warning_topmost(self, title, message):
        def _warn():
            self.force_top_window()
            messagebox.showwarning(title, message, parent=self.root)
        self._run_on_ui(_warn)

    def confirm_topmost(self, title, message):
        def _conf():
            self.force_top_window()
            return messagebox.askyesno(title, message, parent=self.root)
        return self._run_on_ui(_conf)


    def show_rfq_popup(self, rfq_number):
        def _show():
            self.force_top_window()
            top = tk.Toplevel(self.root)
            top.title("RFQ Encontrada")
            top.attributes('-topmost', True)
            
            ws = self.screen_width
            hs = self.screen_height
            w, h = 250, 120
            x = int((ws/2) - (w/2))
            y = int((hs/2) - (h/2))
            top.geometry('%dx%d+%d+%d' % (w, h, x, y))
            
            tk.Label(top, text=f"RFQ Encontrada:\n{rfq_number}", font=("Arial", 11, "bold")).pack(pady=10)
            
            def copy_and_close():
                import subprocess
                try:
                    subprocess.run("clip", text=True, input=str(rfq_number).strip(), check=True)
                except Exception as e:
                    print("Erro no clip:", e)
                    self.root.clipboard_clear()
                    self.root.clipboard_append(rfq_number)
                    self.root.update()
                top.destroy()
                
            tk.Button(top, text="COPIAR", bg="green", fg="white", font=("Arial", 10, "bold"), command=copy_and_close).pack(pady=5, fill=tk.X, padx=20)
            self.root.wait_window(top)
        self._run_on_ui(_show)

    def mostrar_aviso_desconto_35(self, mensagem):
        def _show_aviso():
            self.force_top_window()
            top = tk.Toplevel(self.root)
            top.title("Aviso: Desconto > 35%")
            top.attributes('-topmost', True)
            
            ws = self.screen_width
            hs = self.screen_height
            w, h = 500, 200
            x = int((ws/2) - (w/2))
            y = int((hs/2) - (h/2))
            top.geometry('%dx%d+%d+%d' % (w, h, x, y))
            
            tk.Label(top, text="Aviso: Desconto acima de 35% detectado!", font=("Arial", 11, "bold"), fg="red").pack(pady=15)
            
            entry_msg = tk.Entry(top, width=65, font=("Arial", 9), justify="center")
            entry_msg.pack(pady=10, padx=10)
            entry_msg.insert(0, mensagem)
            
            def copiar_mensagem():
                self.root.clipboard_clear()
                self.root.clipboard_append(mensagem)
                self.root.update()
                top.destroy()
                
            tk.Button(top, text="COPIAR E FECHAR", bg="green", fg="white", font=("Arial", 10, "bold"), command=copiar_mensagem).pack(pady=15, fill=tk.X, padx=30)
            self.root.wait_window(top)
        self._run_on_ui(_show_aviso)

    def detectar_erro_lote_minimo(self):
        """
        Verifica se há erro ou aviso de lote mínimo na tela do LF
        (tanto dentro do iframe do item quanto na página principal).
        Retorna (True, valor_lote, mensagem) se detectado, ou (False, None, None).
        """
        if not getattr(self, 'driver_lf', None):
            return False, None, None
            
        driver = self.driver_lf
        
        # 1. Alerta nativo do navegador
        try:
            alert = driver.switch_to.alert
            txt = alert.text
            if "lote" in txt.lower() and ("minimo" in txt.lower() or "mínimo" in txt.lower()):
                alert.accept()
                lote_val = self._extrair_numero_lote_minimo(txt)
                return True, lote_val, txt
        except Exception:
            pass

        # 2. Executa varredura JavaScript em todos os documentos (iframe e window.top)
        js_code = """
        try {
            var docs = [document];
            try {
                if (window.top && window.top.document && window.top.document !== document) {
                    docs.push(window.top.document);
                }
            } catch(e) {}
            
            try {
                var iframes = document.querySelectorAll("iframe");
                for (var f = 0; f < iframes.length; f++) {
                    try {
                        if (iframes[f].contentDocument) {
                            docs.push(iframes[f].contentDocument);
                        }
                    } catch(e) {}
                }
            } catch(e) {}

            for (var i = 0; i < docs.length; i++) {
                var doc = docs[i];
                var errorSelectors = [
                    ".a-Notification--error",
                    ".t-Alert--error",
                    ".a-Form-error",
                    "#apexPageMessage",
                    ".ui-dialog-content",
                    ".a-AlertMessage",
                    "#P8_QUANTIDADE_error",
                    "[id*='_error']",
                    ".u-Form-field-error",
                    ".t-Alert",
                    ".a-Notification"
                ];
                
                for (var s = 0; s < errorSelectors.length; s++) {
                    var els = doc.querySelectorAll(errorSelectors[s]);
                    for (var e = 0; e < els.length; e++) {
                        var el = els[e];
                        var txt = (el.innerText || el.textContent || "").trim();
                        if (/lote\\s*m[íi]nimo|quantidade\\s*m[íi]nima|qtd\\s*m[íi]nima/i.test(txt)) {
                            return { detected: true, text: txt };
                        }
                    }
                }
                
                var bodyText = doc.body ? (doc.body.innerText || doc.body.textContent || "") : "";
                if (/lote\\s*m[íi]nimo/i.test(bodyText)) {
                    var lines = bodyText.split("\\n");
                    for (var l = 0; l < lines.length; l++) {
                        if (/lote\\s*m[íi]nimo/i.test(lines[l])) {
                            return { detected: true, text: lines[l].trim() };
                        }
                    }
                }
            }
            return { detected: false, text: "" };
        } catch(e) {
            return { detected: false, text: "" };
        }
        """
        try:
            res = driver.execute_script(js_code)
            if res and res.get("detected"):
                txt = res.get("text", "")
                lote_val = self._extrair_numero_lote_minimo(txt)
                return True, lote_val, txt
        except Exception:
            pass

        return False, None, None

    def _extrair_numero_lote_minimo(self, texto):
        """Extrai o número do lote mínimo a partir do texto da mensagem."""
        if not texto:
            return None
        padroes = [
            r'lote\\s*m[íi]nimo[^\\d]*(\\d+(?:[.,]\\d+)?)',
            r'm[íi]nimo\\s*(?:de|:)?\\s*(\\d+(?:[.,]\\d+)?)',
            r'inferior\\s*(?:a|ao)?\\s*(\\d+(?:[.,]\\d+)?)',
            r'(\\d+(?:[.,]\\d+)?)\\s*(?:un|pc|pç|peça|peças|unidades?)'
        ]
        for padrao in padroes:
            m = re.search(padrao, texto, re.IGNORECASE)
            if m:
                num_str = m.group(1).replace(',', '.')
                try:
                    num_flt = float(num_str)
                    if num_flt.is_integer():
                        return str(int(num_flt))
                    return str(num_flt)
                except Exception:
                    return num_str
        return None

    def mostrar_dialogo_lote_minimo(self, item_weg, lote_minimo_sugerido=""):
        def _dialog():
            self.force_top_window()
            top = tk.Toplevel(self.root)
            top.title("⚠️ Lote Mínimo Detectado no Portal LF")
            top.attributes('-topmost', True)
            top.resizable(False, False)
            
            w, h = 530, 360
            ws = self.screen_width
            hs = self.screen_height
            x = int((ws/2) - (w/2))
            y = int((hs/2) - (h/2))
            top.geometry(f'{w}x{h}+{x}+{y}')
            top.configure(bg="#F8FAFC")
            
            resultado = {"acao": "cancelar", "lote_minimo": lote_minimo_sugerido}
            
            frame_header = tk.Frame(top, bg="#FEF3C7", padx=15, pady=10, relief="solid", bd=1)
            frame_header.pack(fill=tk.X, padx=15, pady=(15, 10))
            
            tk.Label(frame_header, text="⚠️ Restrição de Lote Mínimo Detectada!", font=("Segoe UI", 11, "bold"), fg="#92400E", bg="#FEF3C7").pack(anchor="w")
            tk.Label(frame_header, text="A quantidade pedida é inferior ao lote mínimo exigido pelo sistema.", font=("Segoe UI", 9), fg="#B45309", bg="#FEF3C7").pack(anchor="w")
            
            frame_info = tk.Frame(top, bg="#FFFFFF", padx=15, pady=10, relief="solid", bd=1)
            frame_info.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
            
            num_ped = self.numero_pedido_atual
            item_oc = item_weg.get("numero_weg", "")
            cod_prod = item_weg.get("codigo_produto", "")
            cod_lf = item_weg.get("codigo_lf", "")
            qtd_pedida = item_weg.get("qtd", "")
            nome_cli = self.dados_pedido_atual.get("cliente_nome") or self.dados_pedido_atual.get("comprador") or "WEG"
            
            info_txt = (
                f"• Pedido WEG:  {num_ped}\n"
                f"• Item OC:     {item_oc} (Cód: {cod_prod} | LF: {cod_lf})\n"
                f"• Cliente:     {nome_cli}\n"
                f"• Qtd Pedida:  {qtd_pedida}"
            )
            tk.Label(frame_info, text=info_txt, font=("Segoe UI", 10), justify=tk.LEFT, bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 8))
            
            frame_lote = tk.Frame(frame_info, bg="#F1F5F9", padx=10, pady=6, relief="groove", bd=1)
            frame_lote.pack(fill=tk.X, pady=2)
            
            tk.Label(frame_lote, text="Lote Mínimo Exigido:", font=("Segoe UI", 10, "bold"), bg="#F1F5F9", fg="#0F172A").pack(side=tk.LEFT, padx=(0, 8))
            
            entry_lote = tk.Entry(frame_lote, font=("Segoe UI", 11, "bold"), width=10, justify="center")
            entry_lote.pack(side=tk.LEFT)
            if lote_minimo_sugerido:
                entry_lote.insert(0, str(lote_minimo_sugerido))
            entry_lote.focus()
            
            frame_btns = tk.Frame(top, bg="#F8FAFC", pady=10)
            frame_btns.pack(fill=tk.X, padx=15)
            
            def on_avisar_lu():
                val_lote = entry_lote.get().strip() or str(lote_minimo_sugerido)
                resultado["acao"] = "avisar_lu"
                resultado["lote_minimo"] = val_lote
                top.destroy()
                
            def on_fazer_lote_minimo():
                val_lote = entry_lote.get().strip() or str(lote_minimo_sugerido)
                resultado["acao"] = "fazer_lote_minimo"
                resultado["lote_minimo"] = val_lote
                top.destroy()
                
            def on_cancelar():
                resultado["acao"] = "cancelar"
                top.destroy()
                
            btn_avisar = tk.Button(
                frame_btns, text="📢 Avisar Lu", 
                bg="#2563EB", fg="white", font=("Segoe UI", 10, "bold"), 
                activebackground="#1D4ED8", activeforeground="white",
                padx=15, pady=6, relief="flat", cursor="hand2",
                command=on_avisar_lu
            )
            btn_avisar.pack(side=tk.LEFT, padx=(0, 8))
            
            btn_fazer = tk.Button(
                frame_btns, text="📦 Fazer Lote Mínimo", 
                bg="#16A34A", fg="white", font=("Segoe UI", 10, "bold"), 
                activebackground="#15803D", activeforeground="white",
                padx=15, pady=6, relief="flat", cursor="hand2",
                command=on_fazer_lote_minimo
            )
            btn_fazer.pack(side=tk.LEFT, padx=(0, 8))
            
            btn_canc = tk.Button(
                frame_btns, text="Cancelar", 
                bg="#94A3B8", fg="white", font=("Segoe UI", 9), 
                activebackground="#64748B", activeforeground="white",
                padx=10, pady=6, relief="flat", cursor="hand2",
                command=on_cancelar
            )
            btn_canc.pack(side=tk.RIGHT)
            
            top.protocol("WM_DELETE_WINDOW", on_cancelar)
            self.root.wait_window(top)
            return resultado
            
        return self._run_on_ui(_dialog)

    def mostrar_popup_mensagem_lu(self, item_weg, lote_minimo):
        def _show():
            self.force_top_window()
            num_ped = self.numero_pedido_atual
            item_oc = item_weg.get("numero_weg", "")
            cod_prod = item_weg.get("codigo_produto", "")
            qtd_pedida = item_weg.get("qtd", "")
            nome_cli = self.dados_pedido_atual.get("cliente_nome") or self.dados_pedido_atual.get("comprador") or "WEG"
            
            if cod_prod:
                item_str = f"{item_oc} ({cod_prod})"
            else:
                item_str = f"{item_oc}"
                
            mensagem_lu = f'Pedido WEG {num_ped}, o item {item_str}, o cliente "{nome_cli}" pediu {qtd_pedida} e o lote mínimo é {lote_minimo}.'
            
            copiado = False
            try:
                import subprocess
                subprocess.run("clip", text=True, input=mensagem_lu, check=True)
                copiado = True
            except Exception:
                try:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(mensagem_lu)
                    self.root.update()
                    copiado = True
                except Exception:
                    pass

            top = tk.Toplevel(self.root)
            top.title("Mensagem para a Lu")
            top.attributes('-topmost', True)
            top.resizable(False, False)
            
            w, h = 550, 310
            ws = self.screen_width
            hs = self.screen_height
            x = int((ws/2) - (w/2))
            y = int((hs/2) - (h/2))
            top.geometry(f'{w}x{h}+{x}+{y}')
            top.configure(bg="#F8FAFC")
            
            resultado = {"acao": "cancelar"}
            
            frame_status = tk.Frame(top, bg="#DCFCE7", padx=12, pady=8, relief="solid", bd=1)
            frame_status.pack(fill=tk.X, padx=15, pady=(15, 8))
            
            status_txt = "✅ Mensagem copiada para a Área de Transferência! Basta colar (Ctrl+V) para a Lu." if copiado else "📋 Copie o texto abaixo para enviar para a Lu:"
            lbl_status = tk.Label(frame_status, text=status_txt, font=("Segoe UI", 9, "bold"), fg="#166534", bg="#DCFCE7")
            lbl_status.pack(anchor="w")
            
            frame_txt = tk.Frame(top, bg="#FFFFFF", padx=10, pady=10, relief="solid", bd=1)
            frame_txt.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
            
            text_area = tk.Text(frame_txt, font=("Segoe UI", 10), wrap="word", height=4, relief="flat", bg="#FFFFFF")
            text_area.pack(fill=tk.BOTH, expand=True)
            text_area.insert("1.0", mensagem_lu)
            text_area.focus()
            text_area.tag_add("sel", "1.0", "end")
            
            def copiar_novamente():
                try:
                    import subprocess
                    subprocess.run("clip", text=True, input=mensagem_lu, check=True)
                except Exception:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(mensagem_lu)
                    self.root.update()
                lbl_status.config(text="✅ Copiado novamente com sucesso! Cole com Ctrl+V.")
                btn_copy.config(text="✅ Copiado!")
            
            frame_botoes = tk.Frame(top, bg="#F8FAFC", pady=8)
            frame_botoes.pack(fill=tk.X, padx=15)
            
            btn_copy = tk.Button(
                frame_botoes, text="📋 Copiar Novamente", 
                bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"),
                activebackground="#1D4ED8", activeforeground="white",
                padx=12, pady=5, relief="flat", cursor="hand2",
                command=copiar_novamente
            )
            btn_copy.pack(side=tk.LEFT)
            
            def on_cancelar_pedido():
                resultado["acao"] = "cancelar"
                top.destroy()
                
            def on_voltar_fazer_lote():
                resultado["acao"] = "fazer_lote_minimo"
                top.destroy()

            btn_cancelar = tk.Button(
                frame_botoes, text="⏹️ Fechar e Cancelar Pedido", 
                bg="#EF4444", fg="white", font=("Segoe UI", 9, "bold"),
                activebackground="#DC2626", activeforeground="white",
                padx=12, pady=5, relief="flat", cursor="hand2",
                command=on_cancelar_pedido
            )
            btn_cancelar.pack(side=tk.RIGHT, padx=(6, 0))
            
            btn_voltar = tk.Button(
                frame_botoes, text="↩️ Fazer Lote Mínimo", 
                bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"),
                activebackground="#15803D", activeforeground="white",
                padx=12, pady=5, relief="flat", cursor="hand2",
                command=on_voltar_fazer_lote
            )
            btn_voltar.pack(side=tk.RIGHT)
            
            top.protocol("WM_DELETE_WINDOW", on_cancelar_pedido)
            self.root.wait_window(top)
            return resultado
            
        return self._run_on_ui(_show)
    
    def aplicar_historico_rfq(self, dados_pedido):
        """Procura a RFQ do pedido no histórico (historico_rfq.json) e define os códigos LF que encaixam
        com segurança (código WEG / preço / descrição). Itens em dúvida recebem sugestões para o usuário
        confirmar na janela de códigos pendentes."""
        if not historico_rfq:
            return
        rfq = dados_pedido.get("rfq", "")
        try:
            registro = historico_rfq.buscar_rfq(rfq)
        except Exception as e:
            print(f"  > Erro ao consultar histórico de RFQ: {e}")
            return
        if not registro:
            print(f"  > RFQ {rfq} não está no histórico (não foi respondida pelo robô de RFQ).")
            return

        itens_alvo = [it for it in dados_pedido.get("itens", []) if not it.get("codigo_lf")]
        if not itens_alvo:
            return

        print(f"  > RFQ {rfq} encontrada no histórico ({len(registro.get('itens', []))} itens, "
              f"status: {registro.get('status', '?')}). Comparando com os itens do pedido...")
        try:
            resultados = historico_rfq.casar_itens_pedido(itens_alvo, registro)
        except Exception as e:
            print(f"  > Erro ao comparar itens com a RFQ: {e}")
            return

        for item, res in zip(itens_alvo, resultados):
            num = item.get("numero_weg")
            item["rfq_historico"] = historico_rfq.normalizar_rfq(rfq)
            if res["status"] == "auto" and res["escolhido"]:
                esc = res["escolhido"]
                item["codigo_lf"] = esc["codigo_lf"]
                item["origem_codigo"] = f"RFQ {rfq}"
                print(f"  > [RFQ] Item {num}: código LF {esc['codigo_lf']} (item {esc.get('item_rfq')} da RFQ | "
                      f"preço RFQ {esc.get('preco')} x pedido {item.get('preco')})")
                if esc.get("nota_lote"):
                    print(f"    ⚠️ Observação de lote na RFQ para este item: {esc['nota_lote']}")
            elif res["candidatos"]:
                sugestoes = []
                for c in res["candidatos"][:5]:
                    ci = c["item"]
                    desc = ci.get("descricao_lf") or ci.get("texto_rfq") or ""
                    sugestoes.append({
                        "codigo_lf": ci.get("codigo_lf"),
                        "preco": ci.get("preco", ""),
                        "descricao": re.sub(r'\s+', ' ', str(desc))[:90],
                        "item_rfq": ci.get("item_rfq"),
                        "preco_ok": c["preco_ok"],
                        "codigo_ok": c["codigo_ok"],
                        "similaridade": c["similaridade"],
                        "nota_lote": ci.get("nota_lote", ""),
                    })
                item["sugestoes_rfq"] = sugestoes
                item["sugestao_rfq"] = res["escolhido"]["codigo_lf"] if res["escolhido"] else ""
                print(f"  > [RFQ] Item {num}: em dúvida entre {len(sugestoes)} opção(ões) — confirme na janela de códigos.")
            else:
                print(f"  > [RFQ] Item {num}: nenhum item respondido na RFQ para comparar.")

        if registro.get("avisos_lote"):
            print(f"  > ⚠️ A RFQ tinha observação de LOTE na planilha: {' | '.join(registro['avisos_lote'])}")

    def _atualizar_status_janela_codigos(self, texto, fg="#DC2626"):
        def _upd():
            if self.lbl_status_robo_codigos and self.lbl_status_robo_codigos.winfo_exists():
                self.lbl_status_robo_codigos.config(text=texto, fg=fg)
        self.root.after(0, _upd)

    def abrir_janela_codigos_pendentes(self, itens_sem_codigo):
        """Abre uma janela moderna e prática listando todos os itens sem código LF para o usuário colar."""
        if not itens_sem_codigo:
            self.codigos_preenchidos_event.set()
            return

        top = tk.Toplevel(self.root)
        self.janela_codigos_pendentes = top
        top.title(f"⚡ Códigos LF Faltantes - Pedido {self.numero_pedido_atual}")
        top.geometry("740x520")
        top.minsize(640, 380)
        top.configure(bg="#F8FAFC")
        top.attributes('-topmost', True)
        
        # Centralizar na tela
        ws = self.screen_width
        hs = self.screen_height
        w, h = 740, 520
        if any(it.get("sugestoes_rfq") for it in itens_sem_codigo):
            w, h = 820, min(720, max(520, hs - 120))
        x = max(0, int((ws - w) / 2))
        y = max(0, int((hs - h) / 2))
        top.geometry(f"{w}x{h}+{x}+{y}")

        # Cabeçalho
        header = tk.Frame(top, bg="#1E293B", padx=16, pady=12)
        header.pack(fill=tk.X)
        
        tk.Label(header, text=f"📋 Códigos LF Pendentes — Pedido {self.numero_pedido_atual}", 
                 font=("Segoe UI", 12, "bold"), fg="#FFFFFF", bg="#1E293B").pack(anchor="w")
        tk.Label(header, text=f"Foram identificados {len(itens_sem_codigo)} itens sem código LF correspondente.\n"
                              "Cole os códigos abaixo enquanto o robô prepara o pedido no LF. Ao terminar, clique em Salvar.", 
                 font=("Segoe UI", 9), fg="#94A3B8", bg="#1E293B", justify="left").pack(anchor="w", pady=(3, 0))

        # Container para a lista com Scrollbar
        container = tk.Frame(top, bg="#F8FAFC", padx=12, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, bg="#F8FAFC", highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#F8FAFC")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)
        
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        entries_map = []

        for idx, item in enumerate(itens_sem_codigo):
            card = tk.Frame(scrollable_frame, bg="#FFFFFF", highlightbackground="#CBD5E1", 
                            highlightthickness=1, padx=12, pady=10)
            card.pack(fill=tk.X, pady=4, padx=2)

            info_frame = tk.Frame(card, bg="#FFFFFF")
            info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            num_weg = item.get('numero_weg', str(10 * (idx + 1)))
            cod_prod = item.get('codigo_produto', '')
            preco_str = item.get('preco', '0,00')
            qtd_str = item.get('qtd', '0')
            desc_str = item.get('descricao', f'Item {idx + 1}')

            row_top = tk.Frame(info_frame, bg="#FFFFFF")
            row_top.pack(fill=tk.X)

            tk.Label(row_top, text=f"Item {idx + 1} (OC: {num_weg})", font=("Segoe UI", 10, "bold"), 
                     fg="#0F172A", bg="#FFFFFF").pack(side=tk.LEFT)
            
            if cod_prod:
                tk.Label(row_top, text=f"  •  Cód WEG: {cod_prod}", font=("Segoe UI", 9, "bold"), 
                         fg="#2563EB", bg="#FFFFFF").pack(side=tk.LEFT)
            
            tk.Label(row_top, text=f"  •  Preço: R$ {preco_str}", font=("Segoe UI", 9, "bold"), 
                     fg="#16A34A", bg="#FFFFFF").pack(side=tk.LEFT)
            tk.Label(row_top, text=f"  •  Qtd: {qtd_str}", font=("Segoe UI", 9), 
                     fg="#64748B", bg="#FFFFFF").pack(side=tk.LEFT)

            tk.Label(info_frame, text=desc_str, font=("Segoe UI", 9), 
                     fg="#475569", bg="#FFFFFF", wraplength=450, justify="left").pack(anchor="w", pady=(3, 0))

            # Lado direito: campo para colar o código LF
            entry_frame = tk.Frame(card, bg="#FFFFFF")
            entry_frame.pack(side=tk.RIGHT, padx=(10, 0))

            tk.Label(entry_frame, text="Código LF (SKU):", font=("Segoe UI", 8, "bold"), 
                     fg="#334155", bg="#FFFFFF").pack(anchor="w")
            
            ent = tk.Entry(entry_frame, font=("Segoe UI", 11, "bold"), width=14, 
                           justify="center", bg="#F8FAFC", fg="#0F172A", relief="solid", bd=1)
            ent.pack(ipady=3)
            entries_map.append((item, ent))

            # --- Sugestões vindas do histórico de RFQs (itens em dúvida) ---
            sugestoes = item.get("sugestoes_rfq") or []
            if sugestoes:
                card.config(highlightbackground="#F59E0B", highlightthickness=2)
                if item.get("sugestao_rfq"):
                    ent.insert(0, item["sugestao_rfq"])
                    ent.config(bg="#FEF3C7")

                sug_frame = tk.Frame(info_frame, bg="#FFFBEB", highlightbackground="#FCD34D", highlightthickness=1, padx=8, pady=5)
                sug_frame.pack(fill=tk.X, pady=(6, 0))
                tk.Label(sug_frame, text=f"⚠️ CONFIRME — RFQ {item.get('rfq_historico', '')}: não tenho certeza. Clique na opção correta (ou digite outro código):",
                         font=("Segoe UI", 8, "bold"), fg="#92400E", bg="#FFFBEB", wraplength=430, justify="left").pack(anchor="w")

                def make_pick(e=ent, cod=None):
                    def _pick():
                        e.delete(0, tk.END)
                        e.insert(0, cod or "")
                        e.config(bg="#DCFCE7")
                    return _pick

                for s in sugestoes:
                    marcas = []
                    if s.get("codigo_ok"): marcas.append("✔ cód. WEG")
                    if s.get("preco_ok"): marcas.append("✔ preço")
                    marcas.append(f"desc {int(round(s.get('similaridade', 0) * 100))}%")
                    txt_btn = f"{s.get('codigo_lf')}  •  R$ {s.get('preco', '')}  •  {s.get('descricao', '')}  [{' | '.join(marcas)}]"
                    if s.get("nota_lote"):
                        txt_btn += f"  ⚠️ {s['nota_lote']}"
                    tk.Button(sug_frame, text=txt_btn, command=make_pick(ent, s.get("codigo_lf")),
                              font=("Segoe UI", 8), bg="#FFFFFF", fg="#0F172A", activebackground="#FEF3C7",
                              relief="solid", bd=1, anchor="w", justify="left", wraplength=430, cursor="hand2"
                              ).pack(fill=tk.X, pady=1)

        # Navegação com Enter e setas
        def make_enter_handler(i):
            def handler(event):
                if i + 1 < len(entries_map):
                    entries_map[i + 1][1].focus_set()
                else:
                    finalizar_preenchimento(ignorar_vazios=False)
            return handler

        for i, (_, ent) in enumerate(entries_map):
            ent.bind("<Return>", make_enter_handler(i))
            if i > 0:
                ent.bind("<Up>", lambda e, prev_i=i-1: entries_map[prev_i][1].focus_set())
            if i + 1 < len(entries_map):
                ent.bind("<Down>", lambda e, next_i=i+1: entries_map[next_i][1].focus_set())

        if entries_map:
            entries_map[0][1].focus_set()

        # Rodapé com Status e Botões
        footer = tk.Frame(top, bg="#F1F5F9", padx=16, pady=10, highlightbackground="#E2E8F0", highlightthickness=1)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        status_lbl = tk.Label(footer, text="⚙️ Robô em andamento: preenchendo cabeçalho no portal LF...", 
                              font=("Segoe UI", 9, "bold"), fg="#0284C7", bg="#F1F5F9")
        status_lbl.pack(side=tk.LEFT)
        self.lbl_status_robo_codigos = status_lbl

        def finalizar_preenchimento(ignorar_vazios=False):
            try:
                canvas.unbind_all("<MouseWheel>")
            except Exception: pass
            
            for item, ent in entries_map:
                val = ent.get().strip()
                if val:
                    if '/' in val: val = val.split('/')[0].strip()
                    if len(val) < 6 and val.isdigit(): val = val.zfill(6)
                    item["codigo_lf"] = val
                    print(f"  > Código LF '{val}' definido para o item {item.get('numero_weg')}")
                elif not ignorar_vazios:
                    item["codigo_lf"] = None
                else:
                    item["codigo_lf"] = None

            self.janela_codigos_pendentes = None
            self.lbl_status_robo_codigos = None
            try:
                top.destroy()
            except Exception: pass
            self.codigos_preenchidos_event.set()

        btn_row = tk.Frame(footer, bg="#F1F5F9")
        btn_row.pack(side=tk.RIGHT)

        btn_ignorar = tk.Button(btn_row, text="⏭️ Continuar (Pular em branco)", font=("Segoe UI", 9), 
                                bg="#E2E8F0", fg="#334155", activebackground="#CBD5E1", 
                                relief="flat", padx=12, pady=6, cursor="hand2",
                                command=lambda: finalizar_preenchimento(ignorar_vazios=True))
        btn_ignorar.pack(side=tk.LEFT, padx=(0, 8))

        btn_salvar = tk.Button(btn_row, text="🚀 Salvar e Continuar", font=("Segoe UI", 10, "bold"), 
                               bg="#16A34A", fg="white", activebackground="#15803D", activeforeground="white", 
                               relief="flat", padx=16, pady=6, cursor="hand2",
                               command=lambda: finalizar_preenchimento(ignorar_vazios=False))
        btn_salvar.pack(side=tk.LEFT)

        top.protocol("WM_DELETE_WINDOW", lambda: finalizar_preenchimento(ignorar_vazios=True))

    def organizar_janelas_lado_a_lado(self):
        width = self.screen_width // 2
        height = self.screen_height - 60 
        
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
            chrome_options_lf = webdriver.ChromeOptions()
            chrome_options_lf.page_load_strategy = 'eager'
            chrome_options_lf.add_argument("--disable-background-timer-throttling")
            chrome_options_lf.add_argument("--disable-backgrounding-occluded-windows")
            chrome_options_lf.add_argument("--disable-renderer-backgrounding")
            chrome_options_lf.add_argument("--disable-features=CalculateWindowOcclusion")
            self.driver_lf = webdriver.Chrome(options=chrome_options_lf)
            
            try:
                self.driver_lf.get("chrome://settings/")
                self.driver_lf.execute_script("chrome.settingsPrivate.setDefaultZoom(0.67);")
            except Exception:
                pass
            
            self.driver_lf.get(URL_PORTAL_LF)
            self.robust_find_and_send_keys(self.driver_lf, By.XPATH, XPATHS_LF["login_user"], creds_lf["user"])
            time.sleep(0.4)
            self.robust_find_and_send_keys(self.driver_lf, By.XPATH, XPATHS_LF["login_pass"], creds_lf["pass"])
            time.sleep(0.4)
            try:
                pass_el = self.driver_lf.find_element(By.XPATH, XPATHS_LF["login_pass"])
                pass_el.send_keys(Keys.TAB)
            except Exception:
                pass
            time.sleep(0.8)
            self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["login_btn"])
            try:
                WebDriverWait(self.driver_lf, 30).until(EC.presence_of_element_located((By.XPATH, XPATHS_LF["nav_menu_btn"])))
            except TimeoutException:
                botoes = self.driver_lf.find_elements(By.XPATH, XPATHS_LF["login_btn"])
                if botoes and botoes[0].is_displayed():
                    try: botoes[0].click()
                    except: self.driver_lf.execute_script("arguments[0].click();", botoes[0])
                    WebDriverWait(self.driver_lf, 20).until(EC.presence_of_element_located((By.XPATH, XPATHS_LF["nav_menu_btn"])))
            print("Login LF OK.")
            
            print(f"Iniciando Portal WEG ({self.local_weg_escolhido}) com config de download...")
            self.download_dir = os.path.join(os.getcwd(), "pedidos_pdf")
            os.makedirs(self.download_dir, exist_ok=True)

            chrome_options = webdriver.ChromeOptions()
            chrome_options.page_load_strategy = 'eager'
            chrome_options.add_argument("--disable-background-timer-throttling")
            chrome_options.add_argument("--disable-backgrounding-occluded-windows")
            chrome_options.add_argument("--disable-renderer-backgrounding")
            chrome_options.add_argument("--disable-features=CalculateWindowOcclusion")
            chrome_options.add_argument("--kiosk-printing")
            prefs = {
                "download.default_directory": self.download_dir, 
                "download.prompt_for_download": False, 
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
                "savefile.default_directory": self.download_dir
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            self.driver_weg = webdriver.Chrome(options=chrome_options)
            try:
                self.driver_weg.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(window, 'print', { value: function(){ console.log('Print preview suprimido'); }, writable: false, configurable: false });"
                })
            except Exception:
                pass
            
            try:
                self.driver_weg.get("chrome://settings/")
                self.driver_weg.execute_script("chrome.settingsPrivate.setDefaultZoom(0.9);")
            except Exception:
                pass
            
            self.driver_weg.get(URL_PORTAL_WEG_JVE)
            self.robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["login_user"], creds_weg["user"])
            time.sleep(0.3)
            self.robust_find_and_send_keys(self.driver_weg, By.XPATH, XPATHS_WEG["login_pass"], creds_weg["pass"])
            time.sleep(0.5)
            self.robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["login_btn"])
            print("Login WEG OK.")
            
            time.sleep(0.8) 
            self.organizar_janelas_lado_a_lado()
            
            # Substitui o popup inicial pelo painel de controle
            self.criar_painel_controle()
            
            # Dispara leitura automática dos pedidos não lidos da WEG em segundo plano
            threading.Thread(target=lambda: self.atualizar_fila_pedidos(dar_f5=False), daemon=True).start()
            
            self.info_topmost("Pronto", "Logins efetuados. Digite o número do pedido no painel para começar.")
            
        except Exception as e:
            self.error_topmost("Erro no Login", f"Não foi possível fazer login.\nErro: {e}")
            self.fechar_tudo()

    # --- LÓGICA DE PROCESSAMENTO SUBSTITUINDO O LOOP ETERNO ---
    def processar_proximo_fila(self):
        """Processa diretamente o próximo pedido sugerido da fila (o mais antigo)."""
        pedidos = getattr(self, 'pedidos_nao_lidos', [])
        if not pedidos:
            self.warning_topmost("Fila Vazia", "Não há pedidos pendentes na fila.")
            return
        prox = pedidos[0]
        # Atualiza a caixa de texto com o número para transparência ao usuário
        self.entry_pedido.delete(0, tk.END)
        self.entry_pedido.insert(0, prox)
        self.iniciar_processamento_pedido(prox)

    def processar_pedido_digitado(self):
        """Processa exatamente o que estiver colocado/digitado na caixa de texto pelo usuário."""
        numero = self.entry_pedido.get().strip()
        if not numero:
            self.warning_topmost("Aviso", "Digite o número do pedido na caixa antes de continuar.")
            return
        self.iniciar_processamento_pedido(numero)

    def processar_pedido(self):
        """Alias para processar o pedido colocado na caixa de texto."""
        self.processar_pedido_digitado()

    def iniciar_processamento_pedido(self, numero):
        """Inicia a execução do pedido no LF e WEG em thread separada."""
        numero = str(numero).strip()
        if not numero: return
        
        self.numero_pedido_atual = numero
        self.itens_desconto_35 = []
        
        # Se o pedido estava na lista de não lidos, remove para avançar a fila
        if hasattr(self, 'pedidos_nao_lidos') and numero in self.pedidos_nao_lidos:
            self.pedidos_nao_lidos.remove(numero)
        
        # Desativa a interface para evitar duplo clique
        if hasattr(self, 'btn_proximo'):
            self.btn_proximo.config(state=tk.DISABLED, text=f"Processando {numero}...")
        if hasattr(self, 'btn_digitado'):
            self.btn_digitado.config(state=tk.DISABLED)
        if hasattr(self, 'btn_recomecar'):
            self.btn_recomecar.config(state=tk.DISABLED)
        if hasattr(self, 'btn_trocar'):
            self.btn_trocar.config(state=tk.DISABLED)
        if hasattr(self, 'btn_atualizar_fila'):
            self.btn_atualizar_fila.config(state=tk.DISABLED)
        self.entry_pedido.config(state=tk.DISABLED)
        
        t = threading.Thread(target=self._executar_processo_pedido, daemon=True)
        t.start()

    def _restaurar_painel(self):
        if hasattr(self, 'btn_proximo'):
            self.btn_proximo.config(state=tk.NORMAL)
        if hasattr(self, 'btn_digitado'):
            self.btn_digitado.config(state=tk.NORMAL)
        if hasattr(self, 'btn_recomecar'):
            self.btn_recomecar.config(state=tk.NORMAL)
        if hasattr(self, 'btn_trocar'):
            self.btn_trocar.config(state=tk.NORMAL)
        if hasattr(self, 'btn_atualizar_fila'):
            self.btn_atualizar_fila.config(state=tk.NORMAL)
        self.entry_pedido.config(state=tk.NORMAL)
        
        # Sincroniza a interface com a fila atualizada de pedidos
        self._sincronizar_ui_fila_pedidos()
        
        self.entry_pedido.focus()
        self.force_top_window()

    def _executar_processo_pedido(self):
        try:
            self.organizar_janelas_lado_a_lado()
            print(f"Buscando dados do pedido {self.numero_pedido_atual} no portal WEG...")
            dados_pedido_completo = self.buscar_dados_weg(self.numero_pedido_atual)
            self.dados_pedido_atual = dados_pedido_completo
            
            if not dados_pedido_completo:
                self.warning_topmost("Não Encontrado", f"Pedido {self.numero_pedido_atual} não localizado.")
                return # Interrompe o processo e volta pro painel
                
            # --- VERIFICAÇÃO DE CNPJ RESTRITO (WEG DOIS IRMÃOS) ---
            if self.local_weg_escolhido == "DI":
                cnpj_raw = dados_pedido_completo.get("cnpj", "").strip()
                cnpj_limpo = re.sub(r'[^\d]', '', cnpj_raw)
                
                # Verifica com regex se contém 0030-02 / 0034-28 (com ou sem pontuação/traço) ou se o CNPJ limpo termina nesses números
                eh_cnpj_restrito = bool(re.search(r'0030[-./]?02|0034[-./]?28', cnpj_raw)) or cnpj_limpo.endswith("003002") or cnpj_limpo.endswith("003428")
                
                if eh_cnpj_restrito:
                    deseja_continuar = self.confirm_topmost(
                        "Aviso - CNPJ Restrito (WEG Dois Irmãos)",
                        f"O pedido {self.numero_pedido_atual} possui CNPJ com final restrito:\n"
                        f"{cnpj_raw}\n\n"
                        "Deseja PROCESSAR este pedido assim mesmo?\n\n"
                        "[Sim] -> Continuar e processar pedido\n"
                        "[Não] -> Cancelar e ir para o próximo"
                    )
                    if not deseja_continuar:
                        print(f"Pedido {self.numero_pedido_atual} cancelado pelo usuário devido ao final do CNPJ ({cnpj_raw}).")
                        return
            # -----------------------------------------------------
                
            if dados_pedido_completo.get("aviso_extracao"):
                print(dados_pedido_completo["aviso_extracao"])
                
            if dados_pedido_completo.get("rfq"):
                self.show_rfq_popup(dados_pedido_completo["rfq"])
                self.aplicar_historico_rfq(dados_pedido_completo)
            
            print("Dados encontrados. Verificando aplicações...")
            
            # Se for pedido do comprador Genesis Deucher, todos os itens vêm do contrato e são tratados como Industrialização
            # Exceção: CNPJ 14.309.992/0001-48 (WEG DRIVES) mantém a aplicação original do pedido (Consumo ou Industrialização)
            cnpj_limpo_ped = re.sub(r'[^\d]', '', str(dados_pedido_completo.get("cnpj", "")))
            if dados_pedido_completo.get("is_genesis") and cnpj_limpo_ped != "14309992000148":
                print("  > Pedido Genesis Deucher: Forçando aplicação de todos os itens para 'Industrialização'")
                for item in dados_pedido_completo["itens"]:
                    item["aplicacao"] = "Industrialização"
            elif cnpj_limpo_ped == "14309992000148":
                print("  > CNPJ 14.309.992/0001-48 (WEG DRIVES): Mantendo aplicação original do pedido (Consumo/Industrialização).")
            
            # --- DETECÇÃO ANTECIPADA DE CÓDIGOS LF FALTANTES ---
            itens_sem_codigo = [it for it in dados_pedido_completo["itens"] if not it.get("codigo_lf")]
            if itens_sem_codigo:
                print(f"  > Detectados {len(itens_sem_codigo)} itens sem código LF. Abrindo tela em paralelo para o usuário...")
                self.codigos_preenchidos_event.clear()
                self.root.after(0, lambda: self.abrir_janela_codigos_pendentes(itens_sem_codigo))
            else:
                print("  > Todos os itens já possuem código LF mapeado.")
                self.codigos_preenchidos_event.set()
            # ----------------------------------------------------

            grupos_aplicacao = {}
            for item in dados_pedido_completo["itens"]:
                app_tipo = item.get("aplicacao", "Industrialização")
                if app_tipo not in grupos_aplicacao:
                    grupos_aplicacao[app_tipo] = []
                grupos_aplicacao[app_tipo].append(item)
                
            eh_misto = len(grupos_aplicacao) > 1
            if eh_misto:
                self.info_topmost("Pedido Misto", "Este pedido contém itens de Consumo e Industrialização. O robô irá separá-los em pedidos distintos.")
            
            # Ordenar para Consumo vir primeiro
            tipos_ordenados = sorted(grupos_aplicacao.keys(), key=lambda x: 0 if 'consumo' in x.lower() else 1)
            caminho_pdf = None
            
            for index, app_tipo in enumerate(tipos_ordenados):
                itens_grupo = grupos_aplicacao[app_tipo]
                print(f"\n--- Iniciando criação de pedido para: {app_tipo} ({len(itens_grupo)} itens) ---")
                
                dados_sub_pedido = dados_pedido_completo.copy()
                dados_sub_pedido["itens"] = itens_grupo
                dados_sub_pedido["aplicacao_final"] = app_tipo
                dados_sub_pedido["eh_misto"] = eh_misto
                
                sub_total = 0.0
                for item in itens_grupo:
                    qtd = self._converter_para_float(item['qtd'])
                    preco = self._converter_para_float(item['preco'])
                    sub_total += qtd * preco
                dados_sub_pedido["valor_total"] = sub_total

                try:
                    if len(self.driver_lf.find_elements(By.XPATH, XPATHS_LF["nav_pedido_link"])) == 0:
                        raise Exception("Menu fechado")
                    if not self.driver_lf.find_element(By.XPATH, XPATHS_LF["nav_pedido_link"]).is_displayed():
                        raise Exception("Menu fechado")
                except Exception:
                     self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_menu_btn"], wait_time=5)
                     time.sleep(0.1)
                
                self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["nav_pedido_link"])
                time.sleep(0.1)
                self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["novo_pedido_btn"])
                print("Página 'Novo Pedido' aberta. Preenchendo...")

                self.preencher_pedido_lf(dados_sub_pedido) 
                
                print("Cabeçalho salvo. Verificando checkbox de industrialização...")
                if self.local_weg_escolhido == "JVE" and MAPA_APLICACAO.get(app_tipo, "USO E CONSUMO/IMOBILIZADO") == "INDUSTRIALIZAÇÃO":
                    self.clicar_checkbox_industrializacao() 
                
                print("Iniciando preenchimento de itens no LF...")
                self.preencher_itens_lf(dados_sub_pedido) 
                
                if index == 0:
                    print("Voltando ao portal WEG para baixar o PDF.")
                    caminho_pdf = self.baixar_pdf_weg()
                    if not caminho_pdf:
                        self.error_topmost("Falha no Download", "Não foi possível baixar o PDF. O anexo não será feito.")
                
                if caminho_pdf:
                    print(f"PDF baixado. Anexando arquivo ao pedido: {caminho_pdf}")
                    self.anexar_pdf_lf(caminho_pdf) 
                    
                print(f"✅ Pedido para {app_tipo} concluído no LF.")
            
            # --- FECHAR O POP-UP DA WEG SOMENTE APÓS O TÉRMINO COMPLETO DO PEDIDO ---
            try:
                print("Pedido 100% processado. Fechando o pop-up do pedido WEG...")
                self.driver_weg.switch_to.window(self.driver_weg.window_handles[0])
                self.driver_weg.switch_to.default_content()
                self.robust_click(self.driver_weg, By.XPATH, XPATHS_WEG["popup_btn_fechar_X"], wait_time=10)
                time.sleep(0.5)
            except Exception as e:
                print(f"  > Aviso: Não consegui fechar o pop-up da WEG automaticamente: {e}")

            # Recarrega a página da WEG (F5) para atualizar status ('Não lida' vira 'Confirmado') e puxar próximos
            try:
                print("Atualizando página da WEG (F5) para obter próximos pedidos 'Não lido'...")
                self.driver_weg.refresh()
                try:
                    WebDriverWait(self.driver_weg, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//table//tbody/tr"))
                    )
                except Exception:
                    pass
                time.sleep(0.4)
                novos_nao_lidos = self.extrair_pedidos_nao_lidos_portal()
                self.pedidos_nao_lidos = novos_nao_lidos
                print(f"Pedidos 'Não lido' restantes na fila: {self.pedidos_nao_lidos}")
            except Exception as _e_ref:
                print(f"Aviso ao atualizar lista WEG pós-pedido: {_e_ref}")
            
            # --- MOSTRAR AVISO DE ITENS COM MAIS DE 35% DE DESCONTO NO FINAL ---
            if self.itens_desconto_35:
                def formatar_lista_portugues(lista):
                    if not lista: return ""
                    lst = []
                    for x in lista:
                        if str(x) not in lst: lst.append(str(x))
                    if len(lst) == 1: return lst[0]
                    elif len(lst) == 2: return f"{lst[0]} e {lst[1]}"
                    else: return ", ".join(lst[:-1]) + f" e {lst[-1]}"
                
                itens_list = [it["item_oc"] for it in self.itens_desconto_35]
                codigos_list = [it["codigo"] for it in self.itens_desconto_35]
                
                txt_itens = formatar_lista_portugues(itens_list)
                txt_codigos = formatar_lista_portugues(codigos_list)
                
                if len(itens_list) == 1:
                    mensagem = f"Pedido WEG {self.numero_pedido_atual} , item {txt_itens}, código {txt_codigos} passou dos 35%"
                else:
                    mensagem = f"Pedido WEG {self.numero_pedido_atual} , os itens {txt_itens}, códigos {txt_codigos} passaram dos 35%"
                
                self.mostrar_aviso_desconto_35(mensagem)
            # ------------------------------------------------------------------
            
            self.info_topmost("Sucesso", f"Pedido {self.numero_pedido_atual} processado com sucesso!")
            
        except Exception as e:
            self.error_topmost("Erro no Processamento", f"Ocorreu um erro ao processar o pedido.\nErro: {e}")
        finally:
            self.root.after(0, self._restaurar_painel)

    def _limpar_valor_numerico(self, texto_sujo):
        try:
            return re.sub(r'[^\d,]', '', texto_sujo)
        except Exception: return ""

    def _converter_para_float(self, texto_sujo):
        try:
            if isinstance(texto_sujo, (int, float)):
                return float(texto_sujo)
            s = str(texto_sujo).strip()
            s = re.sub(r'[^\d,.]', '', s)
            if not s:
                return 0.0
            if ',' in s and '.' in s:
                s = s.replace('.', '').replace(',', '.')
            elif ',' in s:
                s = s.replace(',', '.')
            return float(s)
        except Exception:
            return 0.0 

    def buscar_dados_weg(self, numero_pedido):
        print(f"Buscando pedido {numero_pedido} no portal WEG...")
        try:
            self.driver_weg.switch_to.default_content()
            link_xpath_dinamico = f"//a[contains(text(), '{numero_pedido}')]"
            links_visiveis = self.driver_weg.find_elements(By.XPATH, link_xpath_dinamico)
            
            clicado = False
            # 1. Se o pedido já estiver visível na tabela da tela atual, clica direto sem precisar pesquisar
            if links_visiveis and links_visiveis[0].is_displayed():
                print(f"  > Pedido {numero_pedido} já está visível na tabela. Clicando diretamente...")
                self.robust_click(self.driver_weg, By.XPATH, link_xpath_dinamico)
                clicado = True
            
            # 2. Se não estiver visível na tabela, pesquisa usando o campo de busca
            if not clicado:
                print(f"  > Pedido {numero_pedido} não visível diretamente. Usando campo 'Pesquisar'...")
                campo_busca = self.robust_wait_for(self.driver_weg, By.XPATH, XPATHS_WEG["campo_busca_pedido"])
                if campo_busca:
                    self.driver_weg.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_busca)
                    campo_busca.click()
                    campo_busca.send_keys(Keys.CONTROL + "a")
                    campo_busca.send_keys(Keys.BACKSPACE)
                    time.sleep(0.1)
                    campo_busca.send_keys(numero_pedido)
                    campo_busca.send_keys(Keys.RETURN)
                    print(f"Busca por {numero_pedido} enviada...")
                    time.sleep(0.4)
                
                self.robust_click(self.driver_weg, By.XPATH, link_xpath_dinamico)
                
            print("Link do pedido clicado. Aguardando pop-up (iframe)...")
        except Exception as e:
            print(f"Erro ao buscar ou clicar no pedido: {e}"); return None
        
        try:
            self.robust_switch_to_iframe(self.driver_weg, 
                                    (By.XPATH, XPATHS_WEG["popup_iframe"]),
                                    (By.XPATH, XPATHS_WEG["popup_iframe_check"]))
            print("Foco mudado para o iframe.")
        except Exception as e:
            print(f"Não consegui encontrar ou mudar para o iframe. {e}")
            self.error_topmost("Erro de Iframe", f"Não consegui 'focar' no pop-up (iframe).")
            return None

        dados_extraidos = {"itens": [], "aviso_extracao": ""}
        wait_weg = WebDriverWait(self.driver_weg, 10)
        try:
            print("Extraindo dados de dentro do iframe...")
            # Confirmação do número do pedido real WEG (45xxxxxxxx)
            try:
                body_text = self.driver_weg.find_element(By.TAG_NAME, "body").text
                m_ped_real = re.search(r'\b(45\d{8})\b', body_text)
                if m_ped_real:
                    num_real = m_ped_real.group(1)
                    if str(self.numero_pedido_atual).strip() != num_real:
                        print(f"  > Ajustando self.numero_pedido_atual de '{self.numero_pedido_atual}' para o pedido real '{num_real}'!")
                        self.numero_pedido_atual = num_real
            except Exception as _e_num:
                print(f"Aviso ao validar número real no iframe: {_e_num}")

            cnpj_element = self.robust_wait_for(self.driver_weg, By.XPATH, XPATHS_WEG["popup_cnpj"])
            dados_extraidos["cnpj"] = cnpj_element.text if cnpj_element else ""
            
            try:
                els_cli = self.driver_weg.find_elements(By.XPATH, "//section[1]//p | //section[1]//h1 | //section[1]//h2 | //section[1]//h3 | //*[contains(text(),'CNPJ')]/preceding-sibling::*")
                for ec in els_cli:
                    t = ec.text.strip()
                    if t and not t.isdigit() and len(t) > 3 and "CNPJ" not in t.upper() and "/" not in t:
                        dados_extraidos["cliente_nome"] = t
                        print(f"Cliente encontrado: {t}")
                        break
            except Exception:
                pass
            
            cond_pagto_element = self.robust_wait_for(self.driver_weg, By.XPATH, XPATHS_WEG["popup_cond_pagto"])
            cond_valor_bruto = cond_pagto_element.text.split('\n')[-1].strip() if cond_pagto_element else "45"
            dados_extraidos["cond_pagto"] = "60" if "60" in cond_valor_bruto else "45"
            
            cif_fob_element = self.robust_wait_for(self.driver_weg, By.XPATH, XPATHS_WEG["popup_cif_fob"])
            frete_valor_bruto = cif_fob_element.text if cif_fob_element else "CIF"
            dados_extraidos["cif_fob"] = "FOB" if "FOB" in frete_valor_bruto.upper() else "CIF"
            
            valor_total_el = self.robust_wait_for(self.driver_weg, By.XPATH, XPATHS_WEG["popup_valor_total"])
            dados_extraidos["valor_total"] = self._converter_para_float(valor_total_el.text) if valor_total_el else 0.0
            
            try:
                rfq_el = self.driver_weg.find_element(By.XPATH, XPATHS_WEG["popup_rfq"])
                dados_extraidos["rfq"] = rfq_el.text.strip()
                print(f"RFQ encontrada: {dados_extraidos['rfq']}")
            except Exception:
                dados_extraidos["rfq"] = ""

            try:
                comprador_el = self.driver_weg.find_element(By.XPATH, "//i[contains(@class, 'icon-user-alt')]/parent::* | /html/body/main/form[1]/div/div[3]/div[2]/div[1]/p[2]")
                dados_extraidos["comprador"] = comprador_el.text.strip()
                print(f"Comprador encontrado: {dados_extraidos['comprador']}")
            except Exception:
                dados_extraidos["comprador"] = ""

            try:
                unidade_el = self.driver_weg.find_element(By.XPATH, "//*[@id='MEComponentManager_MECollapse_1']//h1 | //h1[contains(., 'Betim')] | //*[contains(text(), 'Betim')]")
                dados_extraidos["unidade_header"] = unidade_el.text.strip()
            except Exception:
                dados_extraidos["unidade_header"] = ""
                
            dados_extraidos["is_genesis"] = "GENESIS" in dados_extraidos["comprador"].upper()
            if dados_extraidos["is_genesis"]:
                print("  > Pedido do comprador Genesis Deucher detectado!")

            print("Extraindo itens do pop-up...")
            elementos_itens = wait_weg.until(EC.presence_of_all_elements_located((By.XPATH, XPATHS_WEG["linhas_itens"])))
            aplicacao_mestra = None
            dados_extraidos["obs_itens"] = []
            
            for i in range(0, len(elementos_itens), 2):
                item_header_el = elementos_itens[i]
                item_details_el = elementos_itens[i+1]
                
                print(f"--- Processando item {i//2 + 1} ---")
                item_data = {}
                item_data["preco"] = self._limpar_valor_numerico(item_details_el.find_element(By.XPATH, XPATHS_WEG["item_preco"]).text)
                item_data["qtd"] = self._limpar_valor_numerico(item_details_el.find_element(By.XPATH, XPATHS_WEG["item_qtd"]).text)
                item_data["aplicacao"] = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_aplicacao"]).text
                
                try:
                    item_data["codigo_produto"] = item_details_el.find_element(By.XPATH, XPATHS_WEG["item_codigo"]).text.strip()
                except Exception:
                    item_data["codigo_produto"] = ""
                
                try:
                    item_data["numero_weg"] = item_header_el.find_element(By.XPATH, XPATHS_WEG["item_numero_weg"]).get_attribute("value")
                except Exception:
                    item_data["numero_weg"] = str(10 * (i // 2 + 1))
                
                try:
                    item_data["descricao"] = item_header_el.find_element(By.XPATH, XPATHS_WEG["item_desc_span"]).text
                except Exception:
                    item_data["descricao"] = f"Item {i//2 + 1}"

                if i == 0:
                    print("  > Extraindo dados de observação do PRIMEIRO item.")
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
                    dados_extraidos["aviso_extracao"] = "AVISO: Pedido contém itens de 'Consumo' e 'Industrialização' misturados!"
                
                print(f"  > Item N°{item_data['numero_weg']} (Cód: {item_data['codigo_produto']}) extraído: Qtd:{item_data['qtd']}, Preço:{item_data['preco']}")
                dados_extraidos["itens"].append(item_data)
                
            dados_extraidos["aplicacao_final"] = aplicacao_mestra
            print("Extração do pop-up concluída. (Permanecendo no iframe)")
            
            # Limpeza do CNPJ e identificação de regras tributárias
            cnpj_raw = dados_extraidos.get("cnpj", "")
            cnpj_limpo = re.sub(r'[^\d]', '', cnpj_raw)
            dados_extraidos["cnpj_limpo"] = cnpj_limpo
            
            # Pró-Emprego só existe para Joinville (SC)
            is_pro_emprego = (cnpj_limpo in CNPJS_PRO_EMPREGO) and (self.local_weg_escolhido == "JVE")
            dados_extraidos["is_pro_emprego"] = is_pro_emprego
            
            is_ind = "industrializa" in str(dados_extraidos.get("aplicacao_final", "")).lower()
            if is_ind:
                if self.local_weg_escolhido == "JVE" and (is_pro_emprego or (dados_extraidos["is_genesis"] and cnpj_limpo != "14309992000148")):
                    dados_extraidos["regra_icms"] = "isento"
                elif self.local_weg_escolhido == "DI":
                    dados_extraidos["regra_icms"] = "4%"
                else:
                    dados_extraidos["regra_icms"] = "12%"
            else:
                dados_extraidos["regra_icms"] = "17%"
                
            print(f"  > Dados fiscais: CNPJ={cnpj_limpo} | Aplicação={dados_extraidos['aplicacao_final']} | Pró-Emprego={is_pro_emprego} | Regra ICMS={dados_extraidos['regra_icms']}")
            
            # --- CORREÇÃO DE PREÇOS COM A TABELA ---
            teve_correcao = False
            eh_cnpj_preco_cliente = (cnpj_limpo == "14309992000148")
            if not dados_extraidos["rfq"] and self.tabela_precos:
                novo_total = 0.0
                for item in dados_extraidos["itens"]:
                    cod = item.get("codigo_produto", "").strip() or item["numero_weg"].strip()
                    cod_norm = self.normalizar_codigo(cod)
                    
                    if cod_norm in self.tabela_precos:
                        entry = self.tabela_precos[cod_norm]
                        preco_selecionado = None
                        codigo_lf_selecionado = None
                        origem_preco = ""
                        
                        item_app = str(item.get("aplicacao", "")).lower()
                        item_is_ind = "industrializa" in item_app if item_app else is_ind
                        
                        if item_is_ind:
                            ind_data = entry.get("industrializacao", {})
                            geral_data = entry.get("geral", {})
                            cons_data = entry.get("consumo", {})
                            codigo_lf_selecionado = ind_data.get("codigo_lf") or geral_data.get("codigo_lf") or cons_data.get("codigo_lf")
                            
                            if not eh_cnpj_preco_cliente:
                                if self.local_weg_escolhido == "JVE" and (is_pro_emprego or dados_extraidos["is_genesis"]):
                                    preco_selecionado = ind_data.get("preco_pro_emprego") or geral_data.get("preco_pro_emprego")
                                    origem_preco = "Coluna I (Pró-Emprego Joinville)"
                                else:
                                    preco_selecionado = ind_data.get("preco_ind_12") or geral_data.get("preco_ind_12") or ind_data.get("preco_pro_emprego") or geral_data.get("preco_pro_emprego")
                                    origem_preco = "Coluna K (INDL 12%)" if (ind_data.get("preco_ind_12") or geral_data.get("preco_ind_12")) else "Coluna I (Fallback)"
                        else:
                            cons_data = entry.get("consumo", {})
                            geral_data = entry.get("geral", {})
                            ind_data = entry.get("industrializacao", {})
                            codigo_lf_selecionado = cons_data.get("codigo_lf") or geral_data.get("codigo_lf") or ind_data.get("codigo_lf")
                            
                            if not eh_cnpj_preco_cliente:
                                preco_selecionado = cons_data.get("preco_consumo") or geral_data.get("preco_consumo") or cons_data.get("preco_pro_emprego") or geral_data.get("preco_pro_emprego")
                                origem_preco = "Coluna L (Consumo 17%)" if (cons_data.get("preco_consumo") or geral_data.get("preco_consumo")) else "Coluna I (Fallback)"
                            
                        if eh_cnpj_preco_cliente:
                            preco_selecionado = None
                            print(f"  > [CNPJ WEG DRIVES - 14309992000148] Item {item['numero_weg']} (Cód: {cod_norm}): Mantendo preço original do pedido ({item['preco']}).")

                        if preco_selecionado:
                            preco_orig_str = item["preco"]
                            item["preco_original"] = preco_orig_str
                            item["preco"] = preco_selecionado
                            item["foi_corrigido"] = True
                            teve_correcao = True
                            print(f"  > [TABELA] Item {item['numero_weg']} (Cód: {cod_norm}): Preço alterado de {preco_orig_str} para {preco_selecionado} via {origem_preco}")
                            
                        if codigo_lf_selecionado:
                            item["codigo_lf"] = codigo_lf_selecionado
                            print(f"  > [TABELA] Item {item['numero_weg']} (Cód: {cod_norm}): SKU LF definido como {codigo_lf_selecionado}")
                            
                    qtd = self._converter_para_float(item.get("qtd", 0))
                    prc = self._converter_para_float(item.get("preco", 0))
                    novo_total += (qtd * prc)
                    
                if teve_correcao:
                    dados_extraidos["valor_total_corrigido"] = novo_total
            
            # --- CONTRATO KANBAN WEG BLUMENAU / GRAVATAÍ ---
            eh_blumenau_pedido = (cnpj_limpo in ["07175725004238", "07175725001484"]) or ("BLUMENAU" in (str(dados_extraidos.get("recebedor", "")) + " " + str(dados_extraidos.get("ponto_descarga", "")) + " " + str(dados_extraidos.get("unidade_header", ""))).upper())
            eh_gravatai_pedido = (cnpj_limpo == "07175725003002") or ("GRAVATAI" in (str(dados_extraidos.get("recebedor", "")) + " " + str(dados_extraidos.get("ponto_descarga", "")) + " " + str(dados_extraidos.get("unidade_header", ""))).upper())

            if hasattr(self, "mapa_blumenau_sap") and self.mapa_blumenau_sap:
                for item in dados_extraidos.get("itens", []):
                    cod = item.get("codigo_produto", "").strip() or item["numero_weg"].strip()
                    cod_norm = self.normalizar_codigo(cod)
                    if cod_norm in self.mapa_blumenau_sap:
                        b_item = self.mapa_blumenau_sap[cod_norm]
                        if not item.get("codigo_lf") and b_item.get("codigo_lf"):
                            item["codigo_lf"] = b_item["codigo_lf"]
                            print(f"  > [KANBAN BLUMENAU] Item {item['numero_weg']} (Cód SAP: {cod_norm}): SKU LF identificado como {b_item['codigo_lf']}")
                        
                        if (eh_blumenau_pedido or eh_gravatai_pedido) and not dados_extraidos.get("rfq"):
                            preco_b = None
                            origem_b = ""
                            if eh_gravatai_pedido:
                                p_val = b_item.get("precos", {}).get("bruto_4")
                                if p_val:
                                    preco_b = f"{float(p_val):.2f}".replace('.', ',')
                                    origem_b = "Contrato Blumenau/Gravataí (Bruto 4%)"
                            else:
                                p_val = b_item.get("precos", {}).get("pro_emprego")
                                if p_val:
                                    preco_b = f"{float(p_val):.2f}".replace('.', ',')
                                    origem_b = "Contrato Blumenau (Pró-Emprego SC)"
                            
                            if preco_b:
                                item["preco_original"] = item.get("preco")
                                item["preco"] = preco_b
                                item["foi_corrigido"] = True
                                teve_correcao = True
                                print(f"  > [KANBAN BLUMENAU] Item {item['numero_weg']} (Cód SAP: {cod_norm}): Preço alterado para {preco_b} via {origem_b}")

            if teve_correcao:
                novo_total = sum(self._converter_para_float(it.get("qtd", 0)) * self._converter_para_float(it.get("preco", 0)) for it in dados_extraidos.get("itens", []))
                dados_extraidos["valor_total_corrigido"] = novo_total
            
            return dados_extraidos
            
        except Exception as e:
            print(f"Erro ao extrair dados DENTRO do pop-up: {e}")
            self.error_topmost("Erro na Extração", f"Não consegui ler os dados do pop-up.\nErro: {e}")
            self.driver_weg.switch_to.default_content(); return None
        
    def preencher_pedido_lf(self, dados_pedido):
        wait = WebDriverWait(self.driver_lf, 25)
        
        def abrir_popup_e_selecionar(valor_busca, termo_click=None, exact_match=False, btn_reabrir=None, selecionar_se_unico=False):
            self.dismiss_apex_alert_if_any(self.driver_lf)
            print(f"  > Popup: Buscando por '{valor_busca}' (Match exato: {exact_match})")
            texto_a_esperar = termo_click if termo_click else valor_busca
            base_xpath_popup_visivel = "//div[contains(@class,'ui-dialog-popuplov') and not(contains(@style,'display: none'))]"

            popup_aberto = False
            for tentativa in range(3):
                self.dismiss_apex_alert_if_any(self.driver_lf)
                self.driver_lf.switch_to.default_content()
                try:
                    WebDriverWait(self.driver_lf, 2.5).until(EC.visibility_of_element_located((By.XPATH, base_xpath_popup_visivel)))
                    popup_aberto = True
                    break
                except TimeoutException:
                    self.dismiss_apex_alert_if_any(self.driver_lf)
                    if btn_reabrir:
                        print(f"  > ⚠️ Popup não detectado (tentativa {tentativa+1}/3). Reabrindo...")
                        self.robust_switch_to_iframe(self.driver_lf, 
                                                     (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), 
                                                     (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
                        self.wait_for_apex_ready(self.driver_lf, timeout=5)
                        self.robust_click(self.driver_lf, btn_reabrir[0], btn_reabrir[1])
                        time.sleep(0.4)
                    else:
                        time.sleep(0.5)

            if not popup_aberto:
                self.driver_lf.switch_to.default_content()
                self.dismiss_apex_alert_if_any(self.driver_lf)
                try:
                    WebDriverWait(self.driver_lf, 5).until(EC.visibility_of_element_located((By.XPATH, base_xpath_popup_visivel)))
                    popup_aberto = True
                except TimeoutException:
                    resp = self.confirm_topmost(
                        "Intervenção Manual - Popup",
                        f"O robô não conseguiu abrir o popup para buscar:\n'{valor_busca}'\n\n"
                        "Por favor, abra a pesquisa manualmente no navegador e aperte 'Sim' para continuar."
                    )
                    if not resp:
                        raise Exception(f"Cancelado pelo usuário: popup não abriu para {valor_busca}")

            def obter_opcoes_popup():
                """Coleta as opções válidas exibidas atualmente no popup LOV."""
                opcoes = []
                try:
                    trs = self.driver_lf.find_elements(
                        By.XPATH, 
                        f"{base_xpath_popup_visivel}//div[contains(@class, 'a-PopupLOV-results')]//table//tbody//tr | {base_xpath_popup_visivel}//table//tbody//tr"
                    )
                    for tr in trs:
                        try:
                            if not tr.is_displayed():
                                continue
                            txt = tr.text.strip()
                            if not txt:
                                continue
                            txt_lower = txt.lower()
                            if any(msg in txt_lower for msg in ["no data found", "nenhum dado", "nenhum registro", "nenhum resultado", "carregando", "loading"]):
                                continue
                            if tr.find_elements(By.TAG_NAME, "th") and not tr.find_elements(By.TAG_NAME, "td"):
                                continue
                            opcoes.append(("tr", tr, txt))
                        except Exception:
                            pass
                    if opcoes:
                        return opcoes

                    lis = self.driver_lf.find_elements(
                        By.XPATH, 
                        f"{base_xpath_popup_visivel}//div[contains(@class, 'a-PopupLOV-results')]//ul/li | {base_xpath_popup_visivel}//ul/li"
                    )
                    for li in lis:
                        try:
                            if not li.is_displayed():
                                continue
                            txt = li.text.strip()
                            if not txt:
                                continue
                            txt_lower = txt.lower()
                            if any(msg in txt_lower for msg in ["no data found", "nenhum dado", "nenhum registro", "nenhum resultado", "carregando", "loading"]):
                                continue
                            opcoes.append(("li", li, txt))
                        except Exception:
                            pass
                except Exception:
                    pass
                return opcoes

            def clicar_opcao(tipo, elemento, texto):
                alvo = elemento
                if tipo == "tr":
                    tds = elemento.find_elements(By.TAG_NAME, "td")
                    if tds:
                        alvo = tds[0]
                self.driver_lf.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", alvo)
                time.sleep(0.1)
                try:
                    alvo.click()
                except Exception:
                    self.driver_lf.execute_script("arguments[0].click();", alvo)
                print(f"✅ Popup: Selecionado: {texto}")

            # Se configurado para selecionar direto quando for única opção (ex: transportadora)
            if selecionar_se_unico:
                t_check = time.time() + 1.5
                opcoes_iniciais = []
                while time.time() < t_check:
                    opcoes_iniciais = obter_opcoes_popup()
                    if len(opcoes_iniciais) > 0:
                        time.sleep(0.3)
                        opcoes_iniciais = obter_opcoes_popup()
                        break
                    time.sleep(0.15)
                
                if len(opcoes_iniciais) == 1:
                    tipo, el, txt = opcoes_iniciais[0]
                    print(f"  > ℹ️ Popup apresentou apenas 1 opção disponível: '{txt}'. Selecionando diretamente sem buscar...")
                    clicar_opcao(tipo, el, txt)
                    try:
                        WebDriverWait(self.driver_lf, 4).until(EC.invisibility_of_element_located((By.XPATH, base_xpath_popup_visivel)))
                    except Exception:
                        pass
                    return True

            try:
                xpath_campo_busca = f"{base_xpath_popup_visivel}//input[contains(@class, 'a-PopupLOV-search') and @type='text']"
                xpath_botao_busca = f"{base_xpath_popup_visivel}//button[contains(@class,'a-PopupLOV-doSearch')]"
                xpath_load_more = f"{base_xpath_popup_visivel}//button[contains(text(), 'Load More') or contains(text(), 'Carregar Mais') or contains(@class, 'a-GV-loadMoreButton')]"
                
                campo_busca = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_campo_busca)))
                botao_busca = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_busca)))
                
                campo_busca.clear()
                campo_busca.send_keys(valor_busca)
                botao_busca.click()
                
                xpath_resultado_especifico = ""
                if exact_match:
                    xpath_resultado_especifico = f"{base_xpath_popup_visivel}//ul/li[normalize-space(.)='{texto_a_esperar}'] | {base_xpath_popup_visivel}//table//tr/td[normalize-space(.)='{texto_a_esperar}']"
                else:
                    xpath_resultado_especifico = f"{base_xpath_popup_visivel}//ul/li[contains(., '{texto_a_esperar}')] | {base_xpath_popup_visivel}//table//tr/td[contains(., '{texto_a_esperar}')]"

                # Espera dinâmica rápida pelo resultado (máx 1.5s com polling de 0.15s)
                t_limite = time.time() + 1.5
                while time.time() < t_limite:
                    try:
                        res = self.driver_lf.find_elements(By.XPATH, xpath_resultado_especifico)
                        if res and res[0].is_displayed():
                            break
                    except Exception:
                        pass
                    time.sleep(0.15)

                encontrado = False
                tentativas_load = 0
                max_load = 5

                while not encontrado and tentativas_load <= max_load:
                    try:
                        resultados = self.driver_lf.find_elements(By.XPATH, xpath_resultado_especifico)
                        if len(resultados) > 0:
                            resultado = resultados[0]
                            self.driver_lf.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", resultado)
                            time.sleep(0.05)
                            resultado.click()
                            print(f"✅ Popup: Selecionado: {texto_a_esperar}")
                            encontrado = True
                            break
                    except Exception:
                        pass
                    
                    if not encontrado:
                        try:
                            btn_load = self.driver_lf.find_elements(By.XPATH, xpath_load_more)
                            if len(btn_load) > 0 and btn_load[0].is_displayed():
                                print("  > Clicando em 'Load More Rows'...")
                                self.driver_lf.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", btn_load[0])
                                time.sleep(0.3)
                                btn_load[0].click()
                                time.sleep(1.5)
                                tentativas_load += 1
                            else:
                                break
                        except Exception:
                            break

                # Se não encontrou e selecionar_se_unico está ativo, tenta limpar a busca para verificar se há apenas 1 opção
                if not encontrado and selecionar_se_unico:
                    print(f"  ⚠️ Valor '{valor_busca}' não encontrado. Limpando busca para reavaliar opções disponíveis...")
                    try:
                        campo_busca.clear()
                        campo_busca.send_keys(Keys.CONTROL + "a")
                        campo_busca.send_keys(Keys.BACKSPACE)
                        botao_busca.click()
                        time.sleep(0.8)
                        opcoes_pos = obter_opcoes_popup()
                        if len(opcoes_pos) == 1:
                            tipo, el, txt = opcoes_pos[0]
                            print(f"  > ℹ️ Apenas 1 opção disponível no popup ('{txt}'). Selecionando...")
                            clicar_opcao(tipo, el, txt)
                            encontrado = True
                    except Exception as e_clear:
                        print(f"  ⚠️ Erro ao tentar redefinir busca: {e_clear}")

                if not encontrado:
                    resp = self.confirm_topmost(
                        "Intervenção Manual - Popup",
                        f"O robô não encontrou automaticamente a opção:\n'{texto_a_esperar}'\n\n"
                        f"Por favor, selecione a opção desejada manualmente no popup do navegador e clique em 'Sim' para continuar.\n"
                        f"(Ou 'Não' para abortar este pedido)."
                    )
                    if resp:
                        print(f"  > Opção selecionada manualmente pelo usuário no popup. Prosseguindo...")
                        encontrado = True
                    else:
                        raise Exception(f"Item '{texto_a_esperar}' não encontrado no popup após tentar carregar mais.")

                try:
                    WebDriverWait(self.driver_lf, 4).until(EC.invisibility_of_element_located((By.XPATH, base_xpath_popup_visivel)))
                except Exception:
                    pass

                return True

            except TimeoutException:
                print(f"❌ Timeout ao tentar selecionar '{valor_busca}' no popup.")
                raise TimeoutException
            except Exception as e:
                print(f"❌ Erro ao selecionar '{valor_busca}' no popup:", e)
                raise e

        try:
            cnpj_limpo = re.sub(r'[^\d]', '', dados_pedido.get('cnpj', ''))
            cond_pagto = dados_pedido.get('cond_pagto', '45')
            frete = dados_pedido.get('cif_fob', 'CIF')
            valor_total = dados_pedido.get('valor_total', 0.0) 
            
            transportadora_texto = ""
            
            # --- LÓGICA DE TRANSPORTADORAS ---
            obs_itens_str = " ".join(dados_pedido.get("obs_itens", [])).upper()
            recebedor_str = str(dados_pedido.get("recebedor", "")).upper()
            ponto_descarga_str = str(dados_pedido.get("ponto_descarga", "")).upper()
            deposito_str = str(dados_pedido.get("deposito", "")).upper()
            cnpj_str = str(dados_pedido.get("cnpj", "")).upper()
            cnpj_limpo_full = re.sub(r'[^\d]', '', cnpj_str)
            unidade_header_str = str(dados_pedido.get("unidade_header", "")).upper()
            
            eh_betim = ("BETIM" in (obs_itens_str + " " + recebedor_str + " " + ponto_descarga_str + " " + deposito_str + " " + cnpj_str + " " + unidade_header_str)) or ("07175725004319" in cnpj_limpo_full)

            eh_blumenau = (cnpj_limpo in ["07175725004238", "07175725001484"]) or ("BLUMENAU" in (obs_itens_str + " " + recebedor_str + " " + ponto_descarga_str + " " + deposito_str + " " + cnpj_str + " " + unidade_header_str))
            eh_gravatai = (cnpj_limpo == "07175725003002") or ("GRAVATAI" in (obs_itens_str + " " + recebedor_str + " " + ponto_descarga_str + " " + deposito_str + " " + cnpj_str + " " + unidade_header_str))

            if self.local_weg_escolhido == "JVE" and eh_betim:
                transportadora_texto = "83630053000113"
                print("  > Pedido WEG Joinville com destino Betim (07.175.725/0043-19). Selecionando transportadora 83630053000113 (TRANSJOI).")
            elif eh_blumenau and frete == "CIF":
                transportadora_texto = "46918010000101"
                print("  > Pedido WEG Blumenau (CIF). Selecionando transportadora 46918010000101 (DISK E TENHA EXPRESS).")
            elif eh_blumenau and frete == "FOB":
                transportadora_texto = "83630053000113"
                print("  > Pedido WEG Blumenau (FOB). Selecionando transportadora 83630053000113 (TRANSJOI).")
            elif eh_gravatai:
                transportadora_texto = "03350624000550"
                print("  > Pedido WEG Gravataí (07.175.725/0030-02). Selecionando transportadora 03350624000550 (TRD).")
            elif cnpj_limpo == "14759173000283" and self.local_weg_escolhido == "DI":
                transportadora_texto = "63553243000177"
            elif cnpj_limpo == "10953379000108" and frete == "FOB": 
                transportadora_texto = "83630053000113"
            elif cnpj_limpo == "10885321000174" and frete == "FOB": # Caso Linhares
                transportadora_texto = "83630053000113"
            elif self.local_weg_escolhido == "DI" and frete == "FOB":
                transportadora_texto = "83630053000113"
            elif self.local_weg_escolhido == "JVE": 
                transportadora_texto = "46918010000101" if frete == "CIF" else "05399007000141"
            else: 
                transportadora_texto = "03350624000550"
            # ---------------------------------

            self.robust_switch_to_iframe(self.driver_lf, 
                                    (By.XPATH, XPATHS_LF["novo_pedido_iframe"]),
                                    (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
            print("Iniciando preenchimento do cabeçalho na ordem correta...")

            filial_valor = "22" if self.local_weg_escolhido == "JVE" else "23"
            el_filial = self.robust_wait_for(self.driver_lf, By.ID, XPATHS_LF["campo_filial"], wait_time=25)
            Select(el_filial).select_by_value(filial_valor)
            self.wait_for_apex_ready(self.driver_lf)
            print(f"✅ 1. Filial: Valor {filial_valor}")
            
            btn_cli = (By.XPATH, "//button[contains(@id,'P7_BEG_VW_CLIENTE_ID_lov_btn')]")
            self.robust_click(self.driver_lf, btn_cli[0], btn_cli[1])
            abrir_popup_e_selecionar(cnpj_limpo, btn_reabrir=btn_cli) 
            self.robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
            self.wait_for_apex_ready(self.driver_lf)
            try:
                nome_lf = self.driver_lf.find_element(By.ID, "P7_BEG_VW_CLIENTE_ID").get_attribute("value")
                if nome_lf and len(nome_lf.strip()) > 2:
                    dados_pedido["cliente_nome"] = nome_lf.strip()
                    if hasattr(self, "dados_pedido_atual") and self.dados_pedido_atual:
                        self.dados_pedido_atual["cliente_nome"] = nome_lf.strip()
            except Exception:
                pass

            aplicacao_weg = dados_pedido.get("aplicacao_final", "Industrialização")
            aplicacao_lf = MAPA_APLICACAO.get(aplicacao_weg, "USO E CONSUMO/IMOBILIZADO")
            el_fin = self.robust_wait_for(self.driver_lf, By.ID, XPATHS_LF["campo_finalidade"], wait_time=25)
            sel_fin = Select(el_fin)
            if sel_fin.first_selected_option.text.strip().upper() != aplicacao_lf.upper():
                sel_fin.select_by_visible_text(aplicacao_lf)
                self.wait_for_apex_ready(self.driver_lf)
                time.sleep(0.5)
            print(f"✅ 3. Finalidade: {aplicacao_lf}")

            self.wait_for_apex_ready(self.driver_lf)
            el_tipo = self.robust_wait_for(self.driver_lf, By.ID, XPATHS_LF["campo_tipo_nf"], wait_time=25)
            sel_tipo = Select(el_tipo)
            if sel_tipo.first_selected_option.text.strip().upper() != "VENDA":
                sel_tipo.select_by_visible_text("VENDA")
                self.wait_for_apex_ready(self.driver_lf)
            print("✅ 4. Tipo de NF selecionado: VENDA")
            
            self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_ordem_compra"], self.numero_pedido_atual)
            self.wait_for_apex_ready(self.driver_lf)
            print(f"✅ 5. N° Ordem Compra: {self.numero_pedido_atual}")

            btn_forma = (By.XPATH, "//button[contains(@id,'P7_BEG_VEN_FORMA_PAGAMENTO_ID_lov_btn')]")
            self.robust_click(self.driver_lf, btn_forma[0], btn_forma[1])
            abrir_popup_e_selecionar("CARTEIRA", exact_match=True, btn_reabrir=btn_forma) 
            self.robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
            self.wait_for_apex_ready(self.driver_lf)

            btn_cond = (By.XPATH, "//button[contains(@id,'P7_BEG_CONDICAO_PAGAMENTO_ID_lov_btn')]")
            self.robust_click(self.driver_lf, btn_cond[0], btn_cond[1])
            abrir_popup_e_selecionar(cond_pagto, exact_match=True, btn_reabrir=btn_cond) 
            self.robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
            self.wait_for_apex_ready(self.driver_lf)

            try:
                el_frete = self.robust_wait_for(self.driver_lf, By.ID, XPATHS_LF["campo_tipo_frete"])
                if el_frete: Select(el_frete).select_by_visible_text(frete)
                self.wait_for_apex_ready(self.driver_lf)
                print(f"✅ 8. Tipo Frete: {frete}")
            except Exception:
                print(f"  ! Não consegui selecionar Tipo de Frete '{frete}'.")

            # Dispensar alertas de transportadora preferencial antes e depois do clique na transportadora
            self.dismiss_apex_alert_if_any(self.driver_lf)
            btn_transp = (By.XPATH, "//button[contains(@id,'P7_BEG_VW_TRANSPORTADORA_ID_lov_btn')]")
            self.robust_click(self.driver_lf, btn_transp[0], btn_transp[1])
            abrir_popup_e_selecionar(transportadora_texto, btn_reabrir=btn_transp, selecionar_se_unico=True) 
            self.robust_switch_to_iframe(self.driver_lf, (By.XPATH, XPATHS_LF["novo_pedido_iframe"]), (By.ID, XPATHS_LF["novo_pedido_iframe_check"]))
            self.wait_for_apex_ready(self.driver_lf)
            
            if frete == "CIF" and valor_total < 400.0:
                print(f"✅ Pedido CIF (R${valor_total}) < R$400. Marcando 'Pedido Complementar'...")
                self.robust_check_checkbox(self.driver_lf, By.ID, XPATHS_LF["radio_ped_complementar_sim"])
                print("  > 'SIM' marcado.")
            else:
                print(f"  > Pedido não se qualifica para 'Complementar'. Pulando.")
            
            # --- LÓGICA DE OBSERVAÇÕES ATUALIZADA ---
            obs_nf = "\n".join(dados_pedido.get("obs_itens", []))
            
            if aplicacao_lf == "INDUSTRIALIZAÇÃO" and self.local_weg_escolhido != "JVE":
                ttd_msg = "Diferimento de ICMS TTD 490\nTermo de concessão nº 205000004022828"
                obs_nf = f"{obs_nf}\n\n{ttd_msg}" if obs_nf else ttd_msg
            
            if dados_pedido.get("rfq"):
                obs_pedido = "Preço conforme RFQ, aut. RAFAEL."
            else:
                obs_pedido = "Pedido acordo WEG"
                
            if dados_pedido.get("eh_misto"):
                obs_pedido += "\nAtender pedido parcial"
            
            if cnpj_limpo == "10885321000174" and frete == "FOB":
                texto_linhares = (
                    "MATERIAL DEVE SER DESPACHADO O MAIS BREVE POSSÍVEL MESMO QUE\n"
                    "PARCIALMENTE SE ATENTAR ÁS ORIENTAÇÕES DE FATURAMENTO CONTIDAS NO PEDIDO\n"
                    "DE COMPRA ENVIAR XML PARA O E-MAIL nfe@weg.net E AGUARDAR O ACEITE\n"
                    "DESPACHAR PELA TRANSPORTADORA TRANSJOI TRANSPORTES (47) 4009-5600 EMAIL:\n"
                    "sac@transjoi.com.br / cs.weg@transjoi.com.br INFORMAR NO CAMPO\n"
                    "OBSERVAÇÕES DA NOTA FISCAL REDESPACHO PARA: RODOMOVEIS TRANSPORTES E\n"
                    "MADEIRAS LTDA CNPJ: 76.816.792/0001-39 Endereço: Rua Carlos Blanck, 63\n"
                    "Bairro: Vila Lalau - Jaraguá do Sul / SC Contato Responsável: Cintia\n"
                    "Monique Dias e Angelo Ribeiro Fone: 47- 3371 4522 / 9940 1400 e 47 -\n"
                    "99945-8943 / 3633-3161"
                )
                obs_nf = f"{obs_nf}\n\n{texto_linhares}" if obs_nf else texto_linhares
                obs_pedido = f"{obs_pedido}\n\n{texto_linhares}"
                
            self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_obs_nf"], obs_nf.strip())
            print(f"✅ 10. Observação NF preenchida.")
            
            self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["campo_obs_pedido"], obs_pedido.strip())
            print(f"✅ 11. Observação Pedido preenchida.")
            # -----------------------------------------
            
            print("Cabeçalho preenchido. Clicando em 'Incluir'...")
            time.sleep(0.1)
            self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_salvar_cabecalho"])

            self.driver_lf.switch_to.default_content()
            self.wait_for_apex_ready(self.driver_lf)
            print("✅ Foco retornado para a página principal do LF. Cabeçalho salvo.")

        except Exception as e:
            self.driver_lf.switch_to.default_content() 
            raise e
            
    def preencher_itens_lf(self, dados_pedido):
        wait_lf = WebDriverWait(self.driver_lf, 25)
        
        if not dados_pedido.get("itens"):
             print("  > Nenhum item extraído da WEG para preencher.")
             return

        # Sincronização: se houver códigos pendentes, aguarda o usuário terminar de preencher
        if not self.codigos_preenchidos_event.is_set():
            print("  > Aguardando usuário preencher os códigos LF faltantes...")
            self._atualizar_status_janela_codigos("⏳ Robô concluiu o cabeçalho no LF e está aguardando seus códigos...")
            self.codigos_preenchidos_event.wait()
            print("  > Códigos liberados pelo usuário! Iniciando inclusão dos itens...")

        print("Iniciando adição de itens...")
        for item_weg in dados_pedido["itens"]:
            codigo_lf = item_weg.get("codigo_lf")
            
            if not codigo_lf:
                print(f"  > Pulando o item {item_weg['numero_weg']} (sem código LF).")
                continue
            
            codigo_lf = str(codigo_lf).strip()
            if '/' in codigo_lf:
                codigo_lf = codigo_lf.split('/')[0].strip()
            if len(codigo_lf) < 6 and codigo_lf.isdigit():
                codigo_lf = codigo_lf.zfill(6)
            
            try:
                print("Aguardando botão '+' (Novo Item Pedido) carregar...")
                self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_adicionar_item"])
                print(f"Clicando em '+' para adicionar item {item_weg['numero_weg']} (Cód LF: {codigo_lf})...")
                
                self.robust_switch_to_iframe(self.driver_lf,
                                        (By.XPATH, XPATHS_LF["item_popup_iframe"]),
                                        (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                print("  > Foco no iframe 'Itens do Pedido'.")
                
                self.robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_sku_btn"])
                
                self.driver_lf.switch_to.default_content()
                print("  > Foco retornado para página principal (para buscar SKU).")
                
                base_xpath_popup_visivel = "//div[contains(@class,'ui-dialog-popuplov') and not(contains(@style,'display: none'))]"
                xpath_campo_busca_sku = f"{base_xpath_popup_visivel}//input[contains(@class, 'a-PopupLOV-search') and @type='text']"
                
                search_el = self.robust_wait_for(self.driver_lf, By.XPATH, xpath_campo_busca_sku, EC.visibility_of_element_located)
                if search_el:
                    search_el.clear()
                    search_el.send_keys(codigo_lf)
                    search_el.send_keys(Keys.RETURN)
                print(f"  > Buscando SKU {codigo_lf}...")
                
                xpath_resultado_sku = f"{base_xpath_popup_visivel}{XPATHS_LF['item_sku_search_result']}"
                self.robust_click(self.driver_lf, By.XPATH, xpath_resultado_sku)
                print(f"  > SKU {codigo_lf} selecionado.")
                
                self.robust_switch_to_iframe(self.driver_lf,
                                        (By.XPATH, XPATHS_LF["item_popup_iframe"]),
                                        (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                print("  > Foco no iframe. Preenchendo campos...")
                
                # Verificar desconto unitário com base em P8_VL_UNITARIO
                try:
                    el_vl_unitario = WebDriverWait(self.driver_lf, 10).until(
                        EC.presence_of_element_located((By.ID, "P8_VL_UNITARIO"))
                    )
                    valor_sistema_str = ""
                    for _ in range(30):
                        valor_sistema_str = el_vl_unitario.get_attribute("value")
                        if valor_sistema_str and valor_sistema_str.strip() != "":
                            break
                        time.sleep(0.05)
                    
                    if valor_sistema_str:
                        val_sistema = self._converter_para_float(valor_sistema_str)
                        val_inserido = self._converter_para_float(item_weg['preco'])
                        if val_sistema > 0:
                            desconto = 1.0 - (val_inserido / val_sistema)
                            print(f"    > Preço Sistema: {val_sistema:.2f} | Preço Inserido: {val_inserido:.2f} | Desconto: {desconto*100:.2f}%")
                            if desconto > 0.35:
                                print(f"    ⚠️ Desconto acima de 35% detectado no item {item_weg['numero_weg']}!")
                                self.itens_desconto_35.append({
                                    "item_oc": item_weg.get("numero_weg", "10"),
                                    "codigo": item_weg.get("codigo_produto", "")
                                })
                except Exception as e_vl:
                    print(f"    ⚠️ Erro ao verificar desconto do item: {e_vl}")

                self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_ordem_compra"], self.numero_pedido_atual)
                print(f"    > N° Ordem Compra: {self.numero_pedido_atual}")

                self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_item_oc"], item_weg['numero_weg'])
                print(f"    > N° Item OC: {item_weg['numero_weg']}")
                
                self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_qtd"], item_weg['qtd'])
                print(f"    > Qtd: {item_weg['qtd']}")
                
                self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["item_popup_preco"], item_weg['preco'])
                print(f"    > Preço: {item_weg['preco']}")
                
                print("  > Clicando em 'Incluir' (item)...")
                self.robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_btn_salvar"])

                self.driver_lf.switch_to.default_content()
                print("  > Foco na página principal. Verificando inclusão do item...")
                
                # Verificação inteligente de fechamento do popup ou detecção de erro de lote mínimo
                salvo = False
                for _ in range(30): # até 3 segundos com verificação rápida
                    time.sleep(0.1)
                    try:
                        self.driver_lf.switch_to.default_content()
                        iframes_abertos = self.driver_lf.find_elements(By.XPATH, XPATHS_LF["item_popup_iframe"])
                        if not iframes_abertos or not iframes_abertos[0].is_displayed():
                            salvo = True
                            break
                    except (StaleElementReferenceException, NoSuchElementException):
                        salvo = True
                        break
                    
                    # Se o iframe ainda está visível, verifica se houve erro de lote mínimo
                    deu_lote, lote_val, msg_lote = self.detectar_erro_lote_minimo()
                    if deu_lote:
                        print(f"  ⚠️ Restrição de Lote Mínimo detectada: '{msg_lote}' (Lote sugerido: {lote_val})")
                        res_dialogo = self.mostrar_dialogo_lote_minimo(item_weg, lote_val or "")
                        acao = res_dialogo.get("acao")
                        lote_definido = res_dialogo.get("lote_minimo") or lote_val or "10"
                        
                        if acao == "avisar_lu":
                            res_lu = self.mostrar_popup_mensagem_lu(item_weg, lote_definido)
                            if res_lu.get("acao") == "fazer_lote_minimo":
                                acao = "fazer_lote_minimo"
                            else:
                                try:
                                    self.driver_lf.switch_to.default_content()
                                    btn_fechar_item = self.driver_lf.find_elements(By.CSS_SELECTOR, ".ui-dialog-titlebar-close")
                                    if btn_fechar_item: btn_fechar_item[0].click()
                                except Exception:
                                    pass
                                raise RecomecarPedido(f"Pedido cancelado para avisar Lu sobre Lote Mínimo no item {item_weg.get('numero_weg')}.")
                        
                        if acao == "fazer_lote_minimo":
                            print(f"  > Ajustando quantidade para o lote mínimo ({lote_definido})...")
                            self.dismiss_apex_alert_if_any(self.driver_lf)
                            
                            self.robust_switch_to_iframe(self.driver_lf,
                                                    (By.XPATH, XPATHS_LF["item_popup_iframe"]),
                                                    (By.ID, XPATHS_LF["item_popup_iframe_check"]))
                            
                            qtd_el = self.driver_lf.find_element(By.ID, XPATHS_LF["item_popup_qtd"])
                            qtd_el.clear()
                            time.sleep(0.1)
                            qtd_el.send_keys(str(lote_definido))
                            time.sleep(0.2)
                            
                            self.robust_click(self.driver_lf, By.ID, XPATHS_LF["item_popup_btn_salvar"])
                            
                            item_weg["qtd_original"] = item_weg.get("qtd")
                            item_weg["qtd"] = str(lote_definido)
                            item_weg["lote_minimo_ajustado"] = True
                            item_weg["foi_corrigido"] = True
                            self.dados_pedido_atual["valor_total_corrigido"] = True
                            
                            self.driver_lf.switch_to.default_content()
                            wait_lf.until(
                                EC.invisibility_of_element_located((By.XPATH, XPATHS_LF["item_popup_iframe"]))
                            )
                            salvo = True
                            print(f"  ✅ Item {codigo_lf} salvo com lote mínimo ({lote_definido})!")
                            break
                        elif acao == "cancelar":
                            try:
                                self.driver_lf.switch_to.default_content()
                                btn_fechar_item = self.driver_lf.find_elements(By.CSS_SELECTOR, ".ui-dialog-titlebar-close")
                                if btn_fechar_item: btn_fechar_item[0].click()
                            except Exception:
                                pass
                            raise RecomecarPedido("Pedido cancelado pelo usuário devido a lote mínimo.")

                if not salvo:
                    self.driver_lf.switch_to.default_content()
                    wait_lf.until(
                        EC.invisibility_of_element_located(
                            (By.XPATH, XPATHS_LF["item_popup_iframe"])
                        )
                    )
                print(f"  > Item {codigo_lf} salvo. Pop-up fechado.")

            except Exception as e:
                print(f"ERRO ao processar item {item_weg.get('numero_weg', '')}: {e}")
                self.driver_lf.switch_to.default_content() 
                raise e

        print("--- Todos os itens adicionados ---")

    def clicar_checkbox_industrializacao(self):
        print("  > Pedido de JVE e INDUSTRIALIZAÇÃO. Procurando checkbox TTD...")
        try:
            self.robust_check_checkbox(self.driver_lf, By.XPATH, XPATHS_LF["checkbox_industrializacao"], wait_time=5)
        except Exception as e:
            print(f"  > ERRO ao tentar marcar o checkbox TTD: {e}")

    def formatar_br(self, val, casas=2):
        return f"{float(val):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def aplicar_correcoes_dom_impressao(self, driver):
        is_ind = "industrializa" in str(self.dados_pedido_atual.get("aplicacao_final", "")).lower()
        cnpj_limpo = re.sub(r'[^\d]', '', str(self.dados_pedido_atual.get("cnpj", "")))
        is_pro_emprego = cnpj_limpo in CNPJS_PRO_EMPREGO
        is_genesis = bool(self.dados_pedido_atual.get("is_genesis", False))
        
        script_base = """
        window.replaceInDOM = function(oldStr, newStr) {
            let html = document.body.innerHTML;
            const escaped = oldStr.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
            let pattern = escaped;
            if (/^\\d/.test(oldStr)) pattern = "(?<!\\\\d)" + pattern;
            if (/\\d$/.test(oldStr)) pattern = pattern + "(?!\\\\d)";
            const regex = new RegExp(pattern, "g");
            document.body.innerHTML = html.replace(regex, newStr);
        };
        
        window.replaceItemTotal = function(pNovo, tNovo) {
            let html = document.body.innerHTML;
            let pEscaped = pNovo.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
            let regex = new RegExp("(" + pEscaped + "(?!\\\\d)(?:[\\\\s\\\\S]{1,150}?))(\\\\d{1,3}(?:\\\\.\\\\d{3})*,\\\\d{2,4})", "g");
            document.body.innerHTML = html.replace(regex, "$1" + tNovo);
        };
        
        window.replaceOrderTotal = function(tNovo) {
            let html = document.body.innerHTML;
            let regex = /(Valor total do pedido(?:[\\s\\S]{1,150}?))(\\d{1,3}(?:\\.\\d{3})*,\\d{2,4})/gi;
            document.body.innerHTML = html.replace(regex, "$1" + tNovo);
        };
        """
        driver.execute_script(script_base)
        
        for item in self.dados_pedido_atual["itens"]:
            if item.get("foi_corrigido") or item.get("lote_minimo_ajustado"):
                qtd_num = self._converter_para_float(item['qtd'])
                preco_novo_num = self._converter_para_float(item['preco'])
                
                # Se a quantidade foi alterada pelo lote mínimo:
                if item.get("lote_minimo_ajustado"):
                    qtd_orig_num = self._converter_para_float(item.get('qtd_original', qtd_num))
                    print(f"  > PDF: Ajustando quantidade do item {item.get('numero_weg')} de {qtd_orig_num} para {qtd_num} (Lote Mínimo)...")
                    for casas in [0, 2, 3]:
                        q_orig = self.formatar_br(qtd_orig_num, casas)
                        q_nova = self.formatar_br(qtd_num, casas)
                        driver.execute_script(f"window.replaceInDOM('{q_orig}', '{q_nova}');")
                
                # Se o preço foi alterado
                if item.get("preco_original"):
                    preco_orig_num = self._converter_para_float(item['preco_original'])
                    p_orig_4 = self.formatar_br(preco_orig_num, 4)
                    p_novo_4 = self.formatar_br(preco_novo_num, 4)
                    p_orig_2 = self.formatar_br(preco_orig_num, 2)
                    p_novo_2 = self.formatar_br(preco_novo_num, 2)
                    driver.execute_script(f"window.replaceInDOM('{p_orig_4}', '{p_novo_4}');")
                    driver.execute_script(f"window.replaceInDOM('{p_orig_2}', '{p_novo_2}');")
                
                total_item_novo = qtd_num * preco_novo_num
                t_novo_4 = self.formatar_br(total_item_novo, 4)
                t_novo_2 = self.formatar_br(total_item_novo, 2)
                p_novo_4 = self.formatar_br(preco_novo_num, 4)
                p_novo_2 = self.formatar_br(preco_novo_num, 2)
                
                driver.execute_script(f"window.replaceItemTotal('{p_novo_4}', '{t_novo_4}');")
                driver.execute_script(f"window.replaceItemTotal('{p_novo_2}', '{t_novo_2}');")
                
        # Substituir o valor total final da nota no rodapé
        if "valor_total_corrigido" in self.dados_pedido_atual or any(it.get("lote_minimo_ajustado") for it in self.dados_pedido_atual.get("itens", [])):
            sum_novo = 0.0
            for item in self.dados_pedido_atual.get("itens", []):
                qtd_num = self._converter_para_float(item['qtd'])
                preco_novo_num = self._converter_para_float(item['preco'])
                sum_novo += qtd_num * preco_novo_num
                
            tot_novo_4 = self.formatar_br(sum_novo, 4)
            tot_novo_2 = self.formatar_br(sum_novo, 2)
            driver.execute_script(f"window.replaceOrderTotal('{tot_novo_4}');")
            driver.execute_script(f"window.replaceOrderTotal('{tot_novo_2}');")
            
        # Regras de ICMS no PDF:
        if self.local_weg_escolhido == "JVE":
            if is_ind:
                if is_pro_emprego or (is_genesis and cnpj_limpo != "14309992000148"):
                    # 1. Industrialização Joinville com Pró-Emprego (ou Genesis) -> ISENTO
                    print("  > PDF: Aplicando regra ICMS ISENTO (Industrialização Pró-Emprego Joinville)...")
                    js_icms = """
                    let html = document.body.innerHTML;
                    const regexFull = /-?(?:17|12|7|4),\\d{2}(?:\\s|&nbsp;)*%(?:<(?!\\/?(?:td|tr|th|table|tbody|thead))[^>]+>|\\s|&nbsp;)*(?:j[áa]|n[ãa]o)\\s*incluso/gi;
                    html = html.replace(regexFull, 'Isento');
                    
                    const regexPart = /-?(?:17|12|7|4),\\d{2}(?:\\s|&nbsp;)*%/gi;
                    html = html.replace(regexPart, 'Isento');
                    
                    const regexStranded = /Isento(?:<(?!\\/?(?:td|tr|th|table|tbody|thead))[^>]+>|\\s|&nbsp;)*(?:j[áa]|n[ãa]o)\\s*incluso/gi;
                    html = html.replace(regexStranded, 'Isento');
                    
                    document.body.innerHTML = html;
                    """
                    driver.execute_script(js_icms)
                else:
                    # 2. Industrialização Joinville SEM Pró-Emprego (outros CNPJs) -> 12% ICMS DESTACADO
                    print("  > PDF: Aplicando regra ICMS 12% DESTACADO (Industrialização Outros CNPJs Joinville)...")
                    js_icms_12 = """
                    let html = document.body.innerHTML;
                    // Substitui 17%, -17%, 4%, etc. por 12,00 %
                    const regexFull = /-?(?:17|7|4),\\d{2}(?:\\s|&nbsp;)*%/gi;
                    html = html.replace(regexFull, '12,00 %');
                    
                    // Substitui Isento na coluna ICMS por 12,00 % (se houver)
                    const regexIsento = /Isento(?=(?:<(?!\\/?(?:td|tr|th|table|tbody|thead))[^>]+>|\\s|&nbsp;)*(?:j[áa]|n[ãa]o)\\s*incluso)/gi;
                    html = html.replace(regexIsento, '12,00 %');
                    
                    document.body.innerHTML = html;
                    """
                    driver.execute_script(js_icms_12)
            else:
                # 3. Consumo Joinville -> Mantém o ICMS destacado de Consumo (17% ou original)
                print("  > PDF: Pedido de Consumo Joinville. Mantendo ICMS original do pedido.")
        elif self.local_weg_escolhido == "DI":
            # 4. Dois Irmãos:
            # "só sai insento quando é industrialização em Joinville"
            # "era pra ter saido os 4% de icms normal"
            print("  > PDF: Pedido WEG Dois Irmãos (DI). Mantendo regra de 4% de ICMS normal (não aplica Isento).")
            if is_ind:
                js_icms_di = """
                let html = document.body.innerHTML;
                // Se houver Isento indevido na coluna de ICMS em pedido de DI, substitui por 4,00 %
                const regexIsento = /Isento(?=(?:<(?!\\/?(?:td|tr|th|table|tbody|thead))[^>]+>|\\s|&nbsp;)*(?:j[áa]|n[ãa]o)\\s*incluso)/gi;
                html = html.replace(regexIsento, '4,00 %');
                document.body.innerHTML = html;
                """
                driver.execute_script(js_icms_di)
            else:
                print("  > PDF: Pedido de Consumo Dois Irmãos. Mantendo ICMS original do pedido.")

    def testar_apenas_download_pdf(self):
        """Modo de teste: apenas abre o pedido na WEG, baixa o PDF com as correções e não mexe no LF."""
        num_pedido = self.entry_pedido.get().strip()
        if not num_pedido:
            self.error_topmost("Aviso", "Por favor, digite o número do pedido.")
            return

        self.numero_pedido_atual = num_pedido
        self.btn_processar.config(state=tk.DISABLED)
        if hasattr(self, 'btn_testar_pdf'):
            self.btn_testar_pdf.config(state=tk.DISABLED)

        def _worker_teste():
            try:
                print(f"\n--- [MODO TESTE] Processando apenas download do PDF: {num_pedido} ---")
                self.dados_pedido_atual = self.buscar_dados_weg(num_pedido)

                # Baixa o PDF
                caminho_pdf = self.baixar_pdf_weg()

                if caminho_pdf and os.path.exists(caminho_pdf):
                    self.info_topmost("Sucesso no Teste", f"✅ PDF baixado e corrigido com sucesso!\n\nSalvo em:\n{os.path.abspath(caminho_pdf)}")
                    try:
                        os.startfile(os.path.dirname(os.path.abspath(caminho_pdf)))
                    except Exception:
                        pass
                else:
                    self.error_topmost("Falha no Download", "❌ Não foi possível baixar o PDF. Verifique os logs no console.")

            except Exception as e:
                self.error_topmost("Erro no Teste", f"Erro durante o teste de download: {e}")
                import traceback
                traceback.print_exc()

            finally:
                def _rest():
                    self.btn_processar.config(state=tk.NORMAL)
                    if hasattr(self, 'btn_testar_pdf'):
                        self.btn_testar_pdf.config(state=tk.NORMAL)
                self.root.after(0, _rest)

        threading.Thread(target=_worker_teste, daemon=True).start()

    def _encontrar_pdf_baixado(self, num_pedido):
        """Procura o PDF do pedido na pasta de downloads, suportando os nomes nativos do ME."""
        caminho_padrao = os.path.join(self.download_dir, f"{num_pedido}.pdf")
        if os.path.exists(caminho_padrao) and os.path.getsize(caminho_padrao) > 1000:
            return caminho_padrao

        # Busca por qualquer PDF que contenha o número do pedido no nome
        padrao = os.path.join(self.download_dir, f"*{num_pedido}*.pdf")
        arquivos = glob.glob(padrao)
        if arquivos:
            arquivos.sort(key=os.path.getmtime, reverse=True)
            mais_recente = arquivos[0]
            if os.path.getsize(mais_recente) > 1000:
                try:
                    import shutil
                    shutil.copy2(mais_recente, caminho_padrao)
                    return caminho_padrao
                except Exception:
                    return mais_recente
        return None

    def baixar_pdf_weg(self):
        print(f"Iniciando download do PDF para o pedido {self.numero_pedido_atual}...")
        num_pedido = str(self.numero_pedido_atual).strip()
        caminho_padrao = os.path.join(self.download_dir, f"{num_pedido}.pdf")

        # Remove arquivo antigo do mesmo pedido para não reutilizar PDF desatualizado
        if os.path.exists(caminho_padrao):
            try:
                os.remove(caminho_padrao)
                print(f"  > Arquivo antigo '{num_pedido}.pdf' removido.")
            except Exception:
                pass
            
        try:
            janela_principal = self.driver_weg.window_handles[0]
            self.driver_weg.switch_to.window(janela_principal)
            time.sleep(0.3)
            
            print("  > Focando no iframe do pedido WEG...")
            self.robust_switch_to_iframe(self.driver_weg,
                                    (By.XPATH, XPATHS_WEG["popup_iframe"]),
                                    (By.XPATH, XPATHS_WEG["popup_iframe_check"]))
            print("  > Foco no iframe do pedido WEG.")
            
            # Clicar no botão 'Opções'
            print("  > Clicando em 'Opções'...")
            btn_opc = WebDriverWait(self.driver_weg, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_WEG["btn_opcoes"]))
            )
            self.driver_weg.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_opc)
            time.sleep(0.3)
            try:
                btn_opc.click()
            except Exception:
                self.driver_weg.execute_script("arguments[0].click();", btn_opc)
            
            time.sleep(0.6)
            
            # Clicar em 'Imprimir'
            print("  > Clicando em 'Imprimir'...")
            btn_imp = WebDriverWait(self.driver_weg, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_WEG["btn_imprimir_menu"]))
            )
            
            handles_antes = list(self.driver_weg.window_handles)
            
            # Clica via JS para disparar o download
            self.driver_weg.execute_script("arguments[0].click();", btn_imp)
            
            # Aguarda o download automático ou abertura da aba de impressão (até 10 segundos)
            janela_impressao = None
            is_ind = "industrializa" in str(self.dados_pedido_atual.get("aplicacao_final", "")).lower()
            is_genesis = bool(self.dados_pedido_atual.get("is_genesis", False))
            teve_correcao = bool(self.dados_pedido_atual.get("valor_total_corrigido")) or any(it.get("lote_minimo_ajustado") for it in self.dados_pedido_atual.get("itens", []))
            precisa_corrigir = teve_correcao or is_ind or is_genesis

            for _ in range(40):
                time.sleep(0.15)
                if not precisa_corrigir:
                    pdf_encontrado = self._encontrar_pdf_baixado(num_pedido)
                    if pdf_encontrado and os.path.exists(pdf_encontrado):
                        print(f"  > PDF detectado e validado no disco: {pdf_encontrado}")
                        return pdf_encontrado
                
                handles_atuais = self.driver_weg.window_handles
                if len(handles_atuais) > len(handles_antes):
                    janela_impressao = [h for h in handles_atuais if h not in handles_antes][0]
                    break
            
            # Se abriu aba de impressão e não baixou direto via navegador:
            if janela_impressao:
                self.driver_weg.switch_to.window(janela_impressao)
                print("  > Foco na aba de impressão.")
                
                try:
                    WebDriverWait(self.driver_weg, 10).until(
                        lambda d: d.execute_script("return document.readyState") in ["interactive", "complete"] and (d.execute_script("return !!document.body && document.body.innerHTML.length > 50") or True)
                    )
                except Exception:
                    pass
                time.sleep(0.3)
                
                if precisa_corrigir:
                    print("  > Aplicando correções no DOM da página de impressão...")
                    self.aplicar_correcoes_dom_impressao(self.driver_weg)

                print("  > Gerando PDF via CDP...")
                print_settings = {
                    "printBackground": True,
                    "paperWidth": 8.27, 
                    "paperHeight": 11.69, 
                    "scale": 1.0 
                }
                result = self.driver_weg.execute_cdp_cmd("Page.printToPDF", print_settings)
                pdf_data = base64.b64decode(result['data'])
                with open(caminho_padrao, 'wb') as f:
                    f.write(pdf_data)

                print(f"  > Download via CDP concluído! Arquivo salvo em: {caminho_padrao}")
                
                self.driver_weg.close()
                self.driver_weg.switch_to.window(janela_principal)
                return caminho_padrao

            # Última checagem no disco
            for _ in range(8):
                time.sleep(0.5)
                pdf_encontrado = self._encontrar_pdf_baixado(num_pedido)
                if pdf_encontrado:
                    return pdf_encontrado

            return None

        except Exception as e:
            print(f"Aviso durante baixar_pdf_weg: {e}")
            pdf_encontrado = self._encontrar_pdf_baixado(num_pedido)
            if pdf_encontrado:
                return pdf_encontrado
            return None
    
    def anexar_pdf_lf(self, caminho_pdf):
        print(f"Iniciando anexo do PDF '{caminho_pdf}' no portal LF...")
        wait_lf = WebDriverWait(self.driver_lf, 25)
        
        try:
            self.driver_lf.switch_to.window(self.driver_lf.window_handles[0])
            self.driver_lf.switch_to.default_content()
            
            print("  > Procurando botão 'Novo' (para Anexo)...")
            self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["btn_anexar"])
            
            print("  > Aguardando iframe 'Anexos Pedido'...")
            self.robust_switch_to_iframe(self.driver_lf,
                                    (By.XPATH, XPATHS_LF["anexo_iframe"]),
                                    (By.ID, XPATHS_LF["anexo_iframe_check"]))
            print("  > Foco no iframe de anexo.")
            
            print(f"  > Preenchendo descrição com: {self.numero_pedido_atual}")
            self.robust_find_and_send_keys(self.driver_lf, By.ID, XPATHS_LF["anexo_descricao_input"], self.numero_pedido_atual)
            
            print("  > Procurando campo de upload (input type=file)...")
            caminho_abs_pdf = os.path.abspath(caminho_pdf)
            print(f"  > Enviando caminho do arquivo: {caminho_abs_pdf}")
            
            campo_upload = self.robust_wait_for(self.driver_lf, By.XPATH, XPATHS_LF["anexo_upload_input"], EC.presence_of_element_located)
            if campo_upload:
                campo_upload.send_keys(caminho_abs_pdf)
            
            print("  > Clicando em 'Incluir' (anexo)...")
            self.robust_click(self.driver_lf, By.XPATH, XPATHS_LF["anexo_btn_salvar"])
            
            self.driver_lf.switch_to.default_content()
            print("  > Aguardando pop-up de anexo fechar...")
            wait_lf.until(
                EC.invisibility_of_element_located(
                    (By.XPATH, XPATHS_LF["anexo_iframe"])
                )
            )
            print("✅ Anexo concluído com sucesso.")
            
        except Exception as e:
            self.driver_lf.switch_to.default_content() 
            raise e
            
    def fechar_tudo(self):
        print("Encerrando automação.")
        if self.driver_lf: self.driver_lf.quit()
        if self.driver_weg: self.driver_weg.quit()
        self.root.destroy()

# --- Ponto de entrada principal da aplicação ---
if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoPedidosApp(root)
    root.mainloop()