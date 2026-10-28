import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import threading
import openpyxl
import os
import sys
import subprocess
import time

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

try:
    import historico_rfq
except Exception as _e_hist:
    historico_rfq = None
    print(f"Aviso: módulo historico_rfq não carregado ({_e_hist}). RFQs não serão salvas no histórico.")

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    import fitz
except Exception:
    fitz = None

try:
    import pypdf
except Exception:
    pypdf = None

# --- Redirecionamento de Logs (Thread-Safe) ---
_orig_print = print

def print(*args, **kwargs):
    _orig_print(*args, **kwargs)
    message = " ".join(map(str, args))
    try:
        if 'root' in globals() and root:
            root.after(0, lambda: _write_to_log(message))
    except:
        pass

def _write_to_log(message):
    try:
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, f"{message}\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)
    except:
        pass

# --- Variáveis Globais ---
rfq_spreadsheet_path = None
product_spreadsheet_path = None

# --- Dicionário de Credenciais ---
LOGIN_CREDENTIALS = {
    "Santa Catarina (SC)": {"user": os.environ.get("WEG_JVE_USER", "weg12"), "pass": os.environ.get("WEG_JVE_PASS", "EmbracoWEG157128@")},
    "Dois Irmãos (DI)": {"user": os.environ.get("WEG_DI_USER", "3C8BF357"), "pass": os.environ.get("WEG_DI_PASS", "LF_GG10@2027a")}
}

# --- Funções de Conversão de PDF ---
def convert_pdf_to_excel(pdf_path, excel_path=None):
    """
    Converte automaticamente um arquivo PDF contendo tabelas/dados para um arquivo Excel (.xlsx).
    Suporta múltiplos mecanismos de extração (pdfplumber, PyMuPDF/fitz, pypdf).
    """
    if not excel_path:
        base, _ = os.path.splitext(pdf_path)
        excel_path = base + "_convertido.xlsx"

    all_rows = []

    # 1. Tenta pdfplumber
    if pdfplumber is not None:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table:
                                if row and any(cell is not None and str(cell).strip() != '' for cell in row):
                                    cleaned = [str(cell).replace('\n', ' ').strip() if cell is not None else '' for cell in row]
                                    all_rows.append(cleaned)
                    else:
                        text = page.extract_text()
                        if text:
                            for line in text.split('\n'):
                                if line.strip():
                                    all_rows.append(line.strip().split())
        except Exception as e:
            print(f"Aviso: pdfplumber falhou ({e}). Tentando PyMuPDF...")

    # 2. Tenta PyMuPDF (fitz)
    if not all_rows and fitz is not None:
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                try:
                    tabs = page.find_tables()
                    if tabs and tabs.tables:
                        for tab in tabs:
                            for row in tab.extract():
                                if row and any(cell is not None and str(cell).strip() != '' for cell in row):
                                    cleaned = [str(cell).replace('\n', ' ').strip() if cell is not None else '' for cell in row]
                                    all_rows.append(cleaned)
                except Exception:
                    pass
                    
                if not all_rows:
                    text = page.get_text("text")
                    if text:
                        for line in text.split('\n'):
                            if line.strip():
                                all_rows.append(line.strip().split())
        except Exception as e:
            print(f"Aviso: PyMuPDF falhou ({e}).")

    # 3. Tenta pypdf
    if not all_rows and pypdf is not None:
        try:
            reader = pypdf.PdfReader(pdf_path)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        if line.strip():
                            all_rows.append(line.strip().split())
        except Exception:
            pass

    if not all_rows:
        raise Exception("Nenhum dado ou tabela foi localizado no arquivo PDF. Certifique-se de ter 'pdfplumber' ou 'PyMuPDF' instalado.")

    df = pd.DataFrame(all_rows)
    df.to_excel(excel_path, index=False, header=False)
    return excel_path

def convert_custom_pdf():
    pdf_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF para converter para Excel",
        filetypes=(("Arquivos PDF", "*.pdf"),)
    )
    if pdf_path:
        try:
            print(f"📄 Convertendo PDF selecionado: {pdf_path}...")
            excel_path = convert_pdf_to_excel(pdf_path)
            messagebox.showinfo("Sucesso", f"PDF convertido para Excel com sucesso!\n\nArquivo gerado:\n{excel_path}")
            print(f"✅ PDF convertido para Excel: {excel_path}")
        except Exception as e:
            messagebox.showerror("Erro na Conversão", f"Não foi possível converter o PDF: {e}")

# --- Funções da Interface Gráfica (GUI) ---
def select_rfq_file():
    global rfq_spreadsheet_path
    path = filedialog.askopenfilename(
        title="Selecione a Planilha de RFQs (Excel ou PDF)",
        filetypes=(("Arquivos Suportados", "*.xlsx;*.xls;*.pdf"), ("Planilhas Excel", "*.xlsx;*.xls"), ("Arquivos PDF", "*.pdf"))
    )
    if path:
        rfq_spreadsheet_path = path
        filename = os.path.basename(path)
        label_rfq_path.config(text=f"🟢 RFQ: {filename}", fg="#15803D")
        check_files_ready()

def select_product_file():
    global product_spreadsheet_path
    path = filedialog.askopenfilename(
        title="Selecione o arquivo de Produtos (Excel ou PDF)",
        filetypes=(("Arquivos Suportados", "*.xlsx;*.xls;*.pdf"), ("Planilhas Excel", "*.xlsx;*.xls"), ("Arquivos PDF", "*.pdf"))
    )
    if path:
        product_spreadsheet_path = path
        filename = os.path.basename(path)
        label_product_path.config(text=f"🟢 Produtos: {filename}", fg="#15803D")
        check_files_ready()

def check_files_ready():
    if rfq_spreadsheet_path and product_spreadsheet_path:
        btn_start.config(state=tk.NORMAL)
        btn_start_single.config(state=tk.NORMAL)
        lbl_status_badge.config(text="🟢 PRONTO PARA INICIAR", bg="#DCFCE7", fg="#166534")

def on_login_type_change(event):
    selected_type = login_type_combo.get()
    credentials = LOGIN_CREDENTIALS.get(selected_type, {"user": "", "pass": ""})
    user_var.set(credentials["user"])
    pass_var.set(credentials["pass"])

def set_gui_state(state):
    is_normal = (state == tk.NORMAL)
    btn_start.config(state=state)
    btn_start_single.config(state=state)
    btn_select_rfq.config(state=state)
    btn_select_product.config(state=state)
    login_type_combo.config(state="readonly" if is_normal else tk.DISABLED)
    user_entry.config(state=state)
    pass_entry.config(state=state)
    row_test_entry.config(state=state)
    entry_col_link.config(state=state)
    entry_col_code.config(state=state)
    entry_row_start.config(state=state)
    
    if not is_normal:
        lbl_status_badge.config(text="🟡 EM EXECUÇÃO...", bg="#FEF08A", fg="#854D0E")
    else:
        if rfq_spreadsheet_path and product_spreadsheet_path:
            lbl_status_badge.config(text="🟢 PRONTO PARA INICIAR", bg="#DCFCE7", fg="#166534")
        else:
            lbl_status_badge.config(text="⚪ AGUARDANDO ARQUIVOS", bg="#E2E8F0", fg="#475569")

def clear_log():
    log_text.config(state=tk.NORMAL)
    log_text.delete("1.0", tk.END)
    log_text.config(state=tk.DISABLED)

def start_automation_thread(single_mode=False):
    if not rfq_spreadsheet_path or not product_spreadsheet_path:
        messagebox.showwarning("Arquivos Faltando", "Por favor, selecione as duas planilhas.")
        return

    username = user_var.get()
    password = pass_var.get()
    login_type = login_type_combo.get() 
    
    col_link_str = col_link_var.get().strip().upper()
    col_code_str = col_code_var.get().strip().upper()
    
    try:
        row_start = int(row_start_var.get().strip())
    except ValueError:
        messagebox.showwarning("Erro", "Linha Inicial deve ser um número válido.")
        return

    if not username or not password:
        messagebox.showwarning("Credenciais Faltando", "Por favor, preencha o usuário e a senha.")
        return
        
    test_row_num = 0
    if single_mode:
        try:
            test_row_num = int(row_test_var.get())
            if test_row_num < row_start:
                messagebox.showwarning("Linha Inválida", f"Por favor, insira um número de linha maior ou igual a {row_start}.")
                return
        except ValueError:
            messagebox.showwarning("Entrada Inválida", "Por favor, insira um número válido para a linha de teste.")
            return

    set_gui_state(tk.DISABLED)
    clear_log()
    
    automation_thread = threading.Thread(target=run_automation, args=(single_mode, username, password, test_row_num, login_type, col_link_str, col_code_str, row_start))
    automation_thread.start()

def run_automation(single_mode, username, password, test_row_num, login_type, col_link_str, col_code_str, row_start):
    driver = None
    try:
        options = Options()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(options=options)
        main_automation_logic(rfq_spreadsheet_path, product_spreadsheet_path, driver, single_mode, username, password, test_row_num, login_type, col_link_str, col_code_str, row_start)
    except Exception as e:
        if "main_automation_logic" not in str(e):
            messagebox.showerror("Erro na Automação", f"Ocorreu um erro inesperado: {e}")
            print(f"Erro Detalhado: {e}")
    finally:
        if driver:
            driver.quit()
        root.after(0, lambda: set_gui_state(tk.NORMAL))

def load_product_data_from_pdf(pdf_path):
    product_data = {}
    if not pdfplumber:
        return product_data
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for t in tables:
                    header_col_map = {}
                    for row in t:
                        if not row or not any(row): continue
                        row_str = [str(c).strip().lower() if c else '' for c in row]
                        
                        if any('código' in s or 'cod' in s for s in row_str):
                            for c, v in enumerate(row_str):
                                if 'código' in v or 'cod' in v: header_col_map['code'] = c
                                elif 'descri' in v: header_col_map['desc'] = c
                                elif 'marca' in v: header_col_map['brand'] = c
                                elif 'ref' in v: header_col_map['ref'] = c
                                elif 'umed' in v or 'unid' in v: header_col_map['umed'] = c
                                elif 'ncm' in v: header_col_map['ncm'] = c
                                elif ('unit' in v or 'preço' in v or 'preco' in v) and 'vlr' not in v: header_col_map['unit'] = c
                                elif v == '%' and 'icms' not in header_col_map: header_col_map['icms'] = c

                        if 'code' in header_col_map:
                            if 'unit' not in header_col_map: header_col_map['unit'] = 10
                            raw_val = row[header_col_map['code']]
                            if not raw_val: continue
                            code_raw = str(raw_val).strip()
                            if '/' in code_raw:
                                code_raw = code_raw.split('/')[0]
                            code_norm = re.sub(r'[^\d]', '', code_raw)
                            if not code_norm or not code_norm.isdigit(): continue
                            if len(code_norm) < 6: code_norm = code_norm.zfill(6)
                            
                            def get_c(key):
                                idx = header_col_map.get(key)
                                if idx is not None and idx < len(row) and row[idx]:
                                    return str(row[idx]).replace('\n', ' ').strip()
                                return ''
                            
                            desc = get_c('desc')
                            brand = get_c('brand')
                            ref = get_c('ref')
                            umed = get_c('umed').upper()
                            ncm = get_c('ncm')
                            unit_raw = get_c('unit')
                            icms = get_c('icms')
                            
                            unit_clean = unit_raw.replace('.', '').replace(',', '.') if ',' in unit_raw else unit_raw
                            try:
                                val_float = float(unit_clean)
                                unit_formatted = f'{val_float:.2f}'.replace('.', ',')
                            except:
                                unit_formatted = unit_raw
                            
                            product_data[code_norm] = {
                                'OBSERVACAO': f'DESC: {desc}, MARCA: {brand}, REF: {ref}',
                                'UMED': umed,
                                'NCM': ncm,
                                'UNITARIO': unit_formatted,
                                'ICMS': icms
                            }
    except Exception as e:
        print(f"Erro ao ler PDF diretamente: {e}")
    return product_data

def is_pdf_file(path):
    if not path or not os.path.exists(path):
        return False
    if str(path).lower().endswith('.pdf'):
        return True
    try:
        with open(path, 'rb') as f:
            header = f.read(4)
            return header == b'%PDF'
    except:
        return False

def safe_read_excel(path):
    try:
        return pd.read_excel(path, header=None, dtype=str).fillna('')
    except Exception as e1:
        try:
            return pd.read_excel(path, header=None, dtype=str, engine='openpyxl').fillna('')
        except Exception:
            try:
                return pd.read_excel(path, header=None, dtype=str, engine='xlrd').fillna('')
            except Exception:
                try:
                    return pd.read_csv(path, header=None, dtype=str).fillna('')
                except Exception:
                    raise e1

def load_product_data_from_spreadsheet(path, login_type):
    if is_pdf_file(path):
        print(f"\n--- Lendo PDF de Produtos Diretamente: {os.path.basename(path)} ---")
        p_data = load_product_data_from_pdf(path)
        if p_data:
            print(f"{len(p_data)} produtos carregados diretamente do PDF.")
            return p_data

    product_data = {}
    try:
        df = safe_read_excel(path)
        print("\n--- Lendo Planilha de Produtos ---")
        
        icms_base_col_index = None
        icms_special_col_index = None
        header_row_found = False

        print("Procurando linha de cabeçalho...")
        for index, row in df.iterrows():
            row_str_values = [str(val).lower() for val in row.values]
            if any('código' in s or 'cod' in s for s in row_str_values) and ('%' in row_str_values or any('icms' in s for s in row_str_values)):
                percent_cols = [i for i, val in enumerate(row.values) if str(val).strip() == '%']
                if len(percent_cols) >= 3:
                    icms_base_col_index = percent_cols[0]
                    icms_special_col_index = percent_cols[2]
                    header_row_found = True
                    break
                elif len(percent_cols) >= 1:
                    icms_base_col_index = percent_cols[0]
                    icms_special_col_index = percent_cols[0]
                    header_row_found = True
                    break
                elif any('icms' in s for s in row_str_values):
                    header_row_found = True
                    break
        
        if not header_row_found:
            filename = os.path.basename(path)
            msg = (
                f"O cabeçalho de produtos não foi encontrado no arquivo:\n'{filename}'\n\n"
                f"💡 Motivo provável:\n"
                f"Este arquivo parece ser a Planilha de RFQs (com a lista de cotações) e foi selecionado no botão 'Planilha de Produtos'.\n\n"
                f"Como resolver:\n"
                f"1. No 1º botão (📄 Selecionar Planilha de RFQs), selecione: '{filename}'.\n"
                f"2. No 2º botão (📊 Selecionar Produtos), selecione o PDF ou planilha de produtos."
            )
            messagebox.showerror("Planilha Invertida / Erro de Formato", msg)
            return None

        for index, row in df.iterrows():
            code_raw = str(row[0]).strip()
            if '/' in code_raw:
                code_raw = code_raw.split('/')[0]
            
            code_normalized = re.sub(r'[^\d]', '', code_raw)
            if not code_normalized or not code_normalized.isdigit():
                continue
            
            if len(code_normalized) < 6:
                code_normalized = code_normalized.zfill(6)
                
            code_key = code_normalized

            def get_val(r, idx):
                return str(r.iloc[idx]) if idx is not None and len(r) > idx else ''

            description = ' '.join(filter(None, [get_val(row, 1), get_val(row, 2), get_val(row, 3)])).strip()
            brand = ' '.join(filter(None, [get_val(row, 4), get_val(row, 5)])).strip()
            ref_forn = ' '.join(filter(None, [get_val(row, 6), get_val(row, 7)])).strip()
            unitario_str = ''.join(filter(None, [get_val(row, 15), get_val(row, 16)])).strip()
            
            base_icms_value = get_val(row, icms_base_col_index).strip() if icms_base_col_index is not None else get_val(row, 25).strip()
            special_check_value = get_val(row, icms_special_col_index).strip() if icms_special_col_index is not None else ''
            
            try:
                valor_numerico_especial = float(special_check_value.replace(',', '.')) if special_check_value else 0
            except:
                valor_numerico_especial = 0
            
            icms = "8,8" if valor_numerico_especial != 0 else base_icms_value
            if not icms: icms = "0"
            
            ncm_raw = get_val(row, 9).strip()
            ncm_val = ncm_raw.split()[0].replace('-', '.') if ncm_raw else ''

            product_data[code_key] = {
                'OBSERVACAO': f"DESC: {description}, MARCA: {brand}, REF: {ref_forn}",
                'UMED': get_val(row, 8).strip().upper(),
                'NCM': ncm_val,
                'UNITARIO': unitario_str.replace(',', '.').strip(),
                'ICMS': icms
            }
        print(f"{len(product_data)} produtos carregados.")
    except Exception as e:
        messagebox.showerror("Erro ao Ler Planilha de Produtos", f"Erro: {e}")
        return None
    return product_data

def process_rfq_codes(codes_str):
    lines = [c.strip() for c in str(codes_str).replace('\r', '\n').split('\n') if c.strip()]
    processed_items = []
    i = 0
    while i < len(lines):
        current = lines[i]
        line = current
        nota = ""

        # --- Observação de LOTE (ex.: "123456-15dd lote 50 pçs" ou linha só com "Lote mínimo 50") ---
        m_lote = re.search(r'\blote', line, re.IGNORECASE)
        if m_lote:
            nota = line[m_lote.start():].strip()
            line = line[:m_lote.start()].strip(' -–:;,()/')
            if not line:
                # Linha só de observação: não é um item, anexa a nota ao item anterior
                if processed_items:
                    processed_items[-1]['nota'] = (processed_items[-1].get('nota', '') + ' | ' + nota).strip(' |')
                i += 1
                continue

        # --- Prazo: aceita "-15dd", " 15dd", " 15 dd" e "15 dias" ---
        days = ""
        m_days = re.search(r'(\d+)\s*(?:dd|dias?)\b', line, re.IGNORECASE)
        if m_days:
            days = m_days.group(1)
            line = (line[:m_days.start()] + ' ' + line[m_days.end():]).strip(' -–')

        parts = [p.strip() for p in line.split('-') if p.strip()]
        code = ""
        
        possible_code_parts = []
        for p in parts:
            if p.isdigit() and len(p) <= 3 and not possible_code_parts:
                pass
            else:
                possible_code_parts.append(p)
                
        if not possible_code_parts and parts:
            code = parts[0]
        else:
            code = "-".join(possible_code_parts)
            
        pure_code = code.replace('.', '').replace('-', '').strip()
        if pure_code.isdigit() and len(pure_code) in [8, 9] and (i + 1) < len(lines):
            nxt_line = lines[i+1].strip()
            if (len(pure_code) == 8 and len(nxt_line) == 2 and nxt_line.isdigit()) or \
               (len(pure_code) == 9 and len(nxt_line) == 1 and nxt_line.isdigit()):
                code = pure_code + nxt_line
                i += 1
                
        processed_items.append({
            'original': current,
            'code': code,
            'days': days,
            'nota': nota
        })
        i += 1
        
    return processed_items

def extrair_textos_lote(*textos):
    """Retorna as linhas (de qualquer célula da linha da planilha) que mencionam 'lote'."""
    encontrados = []
    for t in textos:
        if t is None:
            continue
        for ln in str(t).replace('\r', '\n').split('\n'):
            ln = ln.strip()
            if ln and re.search(r'\blote', ln, re.IGNORECASE) and ln not in encontrados:
                encontrados.append(ln)
    return encontrados

def extrair_texto_item_portal(driver, relative_index):
    """Lê o texto do item na tela da RFQ (descrição/código WEG), sem os textos dos selects."""
    script = """
        var el = document.getElementById(arguments[0]);
        if (!el) return '';
        var c = el.cloneNode(true);
        c.querySelectorAll('select,option,script,style,input,textarea,button').forEach(function(n){ n.remove(); });
        var t = (c.textContent || '').replace(/\\s+/g, ' ').trim();
        if (t.length < 20 && el.parentElement) {
            var p = el.parentElement.cloneNode(true);
            p.querySelectorAll('select,option,script,style,input,textarea,button').forEach(function(n){ n.remove(); });
            t = (p.textContent || '').replace(/\\s+/g, ' ').trim();
        }
        return t;
    """
    try:
        txt = driver.execute_script(script, f"tdItem_{relative_index}") or ""
        return str(txt)[:800]
    except Exception:
        return ""

def find_product_match(code_to_find, all_products):
    normalized_code = code_to_find
    if '/' in normalized_code:
        normalized_code = normalized_code.split('/')[0]
        
    if len(normalized_code) < 6 and normalized_code.isdigit():
        normalized_code = normalized_code.zfill(6)

    if normalized_code in all_products: return normalized_code
    for length in [9, 8, 7]:
        if len(normalized_code) < length: continue
        prefix = normalized_code[:length]
        matches = [k for k in all_products if k.startswith(prefix)]
        if len(matches) == 1: return matches[0]
        if len(matches) > 1:
            suffix = normalized_code[length:]
            for m in matches:
                if m.endswith(suffix): return m
    return None

def selecionar_unidade_segura(driver, element_id, unidade_preferida):
    try:
        select_element = Select(driver.find_element(By.ID, element_id))
        tentativas = [
            {"PECA": "PEC", "PEÇA": "PEC", "JOGO": "JG", "UNIDADE": "UN", "UNIDA DE": "UN"}.get(unidade_preferida, unidade_preferida),
            "PEC",
            "UN"
        ]
        sucesso = False
        for tentativa in tentativas:
            try:
                select_element.select_by_value(tentativa)
                sucesso = True
                break
            except:
                continue
        if not sucesso:
            try:
                select_element.select_by_index(1)
            except: pass
    except Exception as e:
        print(f"  - Erro no seletor de unidade: {e}")

def col_letter_to_index(letter):
    letter = letter.upper()
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def show_custom_choice_dialog(title, prompt, aviso=None):
    event = threading.Event()
    result = {"choice": "cancel"}
    
    def create_dialog():
        dialog = tk.Toplevel(root)
        dialog.title(title)
        altura = 190
        if aviso:
            altura += 60 + 18 * min(str(aviso).count('\n') + 1, 10)
        dialog.geometry(f"480x{altura}")
        dialog.configure(bg="#F8FAFC")
        dialog.resizable(False, False)
        dialog.transient(root)
        dialog.grab_set()
        dialog.attributes('-topmost', True)
        
        # Centralizar em relação ao root
        x = root.winfo_x() + (root.winfo_width() // 2) - 240
        y = root.winfo_y() + (root.winfo_height() // 2) - (altura // 2)
        dialog.geometry(f"+{x}+{y}")
        
        header_frame = tk.Frame(dialog, bg="#1E293B", pady=8, padx=15)
        header_frame.pack(fill='x')
        lbl_head = tk.Label(header_frame, text=f"⚡ {title}", font=("Segoe UI", 10, "bold"), fg="#FFFFFF", bg="#1E293B")
        lbl_head.pack(anchor="w")
        
        body_frame = tk.Frame(dialog, bg="#F8FAFC", padx=20, pady=15)
        body_frame.pack(fill='both', expand=True)

        label = tk.Label(body_frame, text=prompt, wraplength=440, justify="center", font=("Segoe UI", 9), bg="#F8FAFC", fg="#0F172A")
        label.pack(pady=(0, 15))

        if aviso:
            aviso_frame = tk.Frame(body_frame, bg="#FEF3C7", highlightbackground="#F59E0B", highlightthickness=1, padx=10, pady=6)
            aviso_frame.pack(fill='x', pady=(0, 12))
            tk.Label(aviso_frame, text="⚠️ ATENÇÃO - LOTE INFORMADO NA PLANILHA (verifique antes de confirmar):", font=("Segoe UI", 9, "bold"), bg="#FEF3C7", fg="#92400E", wraplength=420, justify="left").pack(anchor="w")
            tk.Label(aviso_frame, text=str(aviso), font=("Segoe UI", 9), bg="#FEF3C7", fg="#78350F", wraplength=420, justify="left").pack(anchor="w", pady=(2, 0))
        
        btn_frame = tk.Frame(body_frame, bg="#F8FAFC")
        btn_frame.pack(fill='x')
        
        def on_confirm_next():
            result["choice"] = "confirm_next"
            dialog.destroy()
            event.set()
            
        def on_next():
            result["choice"] = "next"
            dialog.destroy()
            event.set()
            
        def on_cancel():
            result["choice"] = "cancel"
            dialog.destroy()
            event.set()
            
        btn_confirm = tk.Button(btn_frame, text="✔ Confirmar e Próximo", command=on_confirm_next, bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=6, cursor="hand2", activebackground="#15803D", activeforeground="white")
        btn_confirm.pack(side='left', expand=True, padx=4)
        
        btn_next = tk.Button(btn_frame, text="⏩ Apenas Próximo", command=on_next, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=6, cursor="hand2", activebackground="#1D4ED8", activeforeground="white")
        btn_next.pack(side='left', expand=True, padx=4)
        
        btn_cancel = tk.Button(btn_frame, text="✖ Cancelar", command=on_cancel, bg="#DC2626", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=6, cursor="hand2", activebackground="#B91C1C", activeforeground="white")
        btn_cancel.pack(side='left', expand=True, padx=4)
        
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)

    root.after(0, create_dialog)
    event.wait()
    return result["choice"]

# --- Lógica de Automação Web ---
def main_automation_logic(rfq_path, product_path, driver, single_mode, username, password, test_row_num, login_type, col_link_str, col_code_str, row_start):
    product_data = load_product_data_from_spreadsheet(product_path, login_type)
    if not product_data: return

    tasks = []
    
    link_idx = col_letter_to_index(col_link_str)
    code_idx = col_letter_to_index(col_code_str)

    try:
        wb = openpyxl.load_workbook(rfq_path)
        sheet = wb.active
        for r_idx, row in enumerate(sheet.iter_rows(min_row=row_start, values_only=False), start=row_start):
            if link_idx >= len(row) or code_idx >= len(row):
                continue
                
            cell_link = row[link_idx]
            cell_code = row[code_idx]
            
            link = cell_link.hyperlink.target if cell_link.hyperlink else (str(cell_link.value).strip() if cell_link.value and str(cell_link.value).strip().startswith('http') else None)
            codes = str(cell_code.value).strip() if cell_code.value else ""
            
            if link and link.startswith('http'):
                codes_lower = codes.lower()
                if not codes_lower or codes_lower == '?' or codes_lower == 'nt' or 'declinar' in codes_lower or 'cotação' in codes_lower or 'cotacao' in codes_lower:
                    continue

                # Número da RFQ (ex.: célula "RFQ: 1093929") e ID da cotação (link ...?Cot=22818686)
                valores_linha = [str(c.value) for c in row if c.value is not None]
                rfq_num = ""
                for v in valores_linha:
                    m_rfq = re.search(r'RFQ\s*:?\s*(\d{4,})', v, re.IGNORECASE)
                    if m_rfq:
                        rfq_num = m_rfq.group(1)
                        break
                m_cot = re.search(r'Cot=(\d+)', link, re.IGNORECASE)
                cot_id = m_cot.group(1) if m_cot else ""

                tasks.append({
                    'link': link, 'codes': codes, 'row_num': r_idx,
                    'rfq': rfq_num, 'cotacao_id': cot_id,
                    'lote_textos': extrair_textos_lote(*valores_linha)
                })
    except Exception as e:
        messagebox.showerror("Erro Planilha RFQ", f"{e}")
        return

    if not tasks:
        messagebox.showinfo("Aviso", "Nenhuma RFQ válida encontrada.")
        return

    tasks_to_process = [t for t in tasks if t['row_num'] == test_row_num] if single_mode else tasks
    if single_mode and not tasks_to_process:
        messagebox.showinfo("Erro", f"Linha {test_row_num} inválida.")
        return

    wait = WebDriverWait(driver, 3)
    
    for task_idx, task in enumerate(tasks_to_process):
        print(f"\nProcessando RFQ Linha {task['row_num']}...")
        driver.get(task['link'])

        # Login
        try:
            wait.until(EC.presence_of_element_located((By.ID, 'LoginName'))).send_keys(username)
            driver.find_element(By.ID, 'RAWSenha').send_keys(password)
            driver.find_element(By.ID, 'SubmitAuth').click()
        except: pass

        # Cabeçalho
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'Frete'))).clear()
            driver.find_element(By.ID, 'Frete').send_keys("0")
            
            # --- CORREÇÃO: TIPO DE FRETE (Incoterms) ---
            try:
                Select(driver.find_element(By.ID, 'selectIcoterms')).select_by_value('CIF')
            except Exception as e:
                print(f"Erro ao selecionar Tipo Frete (CIF): {e}")

            # --- CONDIÇÃO DE PAGAMENTO (45 dias) ---
            try:
                Select(driver.find_element(By.ID, 'selectCondicaopagamento')).select_by_value('C045')
            except Exception as e:
                print(f"Erro ao selecionar Condição de Pagamento: {e}")

            # --- Tenta preencher contato se estiver vazio ---
            try:
                contato = driver.find_element(By.ID, 'inputNomeContato')
                if not contato.get_attribute('value'):
                    contato.send_keys("LUCIANO SOUZA SANTOS")
                    driver.find_element(By.ID, 'inputNumFoneCota').send_keys("47992705317")
            except: pass

            driver.find_element(By.ID, 'ValidadePropostaAux').clear()
            driver.find_element(By.ID, 'ValidadePropostaAux').send_keys((datetime.now() + timedelta(days=15)).strftime('%d/%m/%Y'))
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            
            for f_id in ['LimiteFat', 'CustoFinanceiro']:
                el = driver.find_element(By.ID, f_id)
                el.clear()
                el.send_keys("0")
            
            Select(driver.find_element(By.ID, 'MoedaCot')).select_by_value('BRL')
            
            try:
                th_elem = driver.find_element(By.XPATH, '//th[contains(., "Nome requisitante")]')
                req_text = th_elem.get_attribute("innerHTML").split("<br>")[0].strip()
                frete_field = driver.find_element(By.ID, 'atrib_LOCALFRETE_1_1_0_0')
                frete_field.clear()
                frete_field.send_keys(req_text)
            except: pass

        except Exception as e:
            print(f"Erro cabeçalho: {e}")
            if not messagebox.askyesno("Erro", f"Falha no cabeçalho. Continuar?"): break
            continue

        # --- Número da RFQ (se não veio da planilha, tenta ler da própria página) ---
        if not task.get('rfq'):
            try:
                body_txt = driver.find_element(By.TAG_NAME, 'body').text
                m_rfq_pg = re.search(r'RFQ\s*:?\s*(\d{4,})', body_txt, re.IGNORECASE)
                if m_rfq_pg:
                    task['rfq'] = m_rfq_pg.group(1)
            except Exception:
                pass
        print(f"RFQ: {task.get('rfq') or '(não identificada)'} | Cotação: {task.get('cotacao_id') or '-'}")

        itens_registro = []

        # --- Processamento dos Itens ---
        try:
            total_itens_elem = wait.until(EC.presence_of_element_located((By.ID, "tdTotalItens")))
            qtd_itens = int(re.search(r'\d+', total_itens_elem.text).group(0))
            parsed_items = process_rfq_codes(task['codes'])

            print(f"Total de itens na RFQ: {qtd_itens}")

            for i in range(1, qtd_itens + 1):
                relative_index = (i - 1) % 10 + 1
                
                try:
                    wait.until(EC.element_to_be_clickable((By.ID, f'Preco{relative_index}')))
                except:
                    print(f"Timeout esperando item {relative_index} carregar.")

                parsed_info = parsed_items[i-1] if (i-1) < len(parsed_items) else {'code': '', 'days': '', 'nota': '', 'original': ''}
                raw_code_info = parsed_info['code']
                custom_days = parsed_info['days']
                
                code_lower = raw_code_info.lower()
                
                should_decline = False
                matched_key = None
                product_info = None

                if 'nt' in code_lower or 'declinar' in code_lower or not raw_code_info:
                    should_decline = True
                else:
                    clean_code = re.sub(r'\s+', '', raw_code_info).replace('.0', '')
                    if '/' in clean_code:
                        clean_code = clean_code.split('/')[0]
                    if len(clean_code) < 6 and clean_code.isdigit():
                        clean_code = clean_code.zfill(6)
                        
                    matched_key = find_product_match(clean_code, product_data)
                    
                    if matched_key:
                        product_info = product_data[matched_key]
                        print(f"Item {i}: Match '{clean_code}' -> '{matched_key}'")
                    else:
                        print(f"Item {i}: Produto '{clean_code}' não encontrado. Declinando.")
                        should_decline = True

                if parsed_info.get('nota'):
                    print(f"Item {i}: ⚠️ Observação de lote na planilha: {parsed_info['nota']}")

                # --- Registro para o histórico de RFQs ---
                try:
                    texto_portal = extrair_texto_item_portal(driver, relative_index)
                    m_cod_weg = re.search(r'C[óo]d(?:igo)?[^:\d]{0,20}:\s*([\w.\-/]+)', texto_portal, re.IGNORECASE)
                    itens_registro.append({
                        'item_rfq': i,
                        'linha_planilha': parsed_info.get('original', ''),
                        'codigo_planilha': raw_code_info,
                        'codigo_lf': matched_key if (product_info and not should_decline) else None,
                        'declinado': bool(should_decline),
                        'preco': str(product_info.get('UNITARIO', '')) if product_info else '',
                        'descricao_lf': str(product_info.get('OBSERVACAO', '')) if product_info else '',
                        'umed': str(product_info.get('UMED', '')) if product_info else '',
                        'ncm': str(product_info.get('NCM', '')) if product_info else '',
                        'icms': str(product_info.get('ICMS', '')) if product_info else '',
                        'prazo': custom_days or '3',
                        'nota_lote': parsed_info.get('nota', ''),
                        'codigo_weg': m_cod_weg.group(1) if m_cod_weg else '',
                        'texto_rfq': texto_portal,
                    })
                except Exception as e_reg:
                    print(f"  - Aviso: não consegui registrar o item {i} no histórico: {e_reg}")

                if should_decline:
                    try:
                        btn_nao = driver.find_element(By.ID, f'btnNaoResponder_{relative_index}')
                        txt_just = driver.find_element(By.ID, f'txtJustificativaRecusa_{relative_index}')
                        
                        if not txt_just.is_displayed():
                            if btn_nao.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", btn_nao)
                                driver.execute_script("arguments[0].click();", btn_nao)
                            wait.until(EC.visibility_of(txt_just))
                            
                        txt_just.clear()
                        txt_just.send_keys("nt")
                    except Exception as e:
                        print(f"Erro ao declinar item {i}: {e}")

                elif product_info:
                    try:
                        # Verificar se o item estava recusado e reativá-lo para preenchimento
                        btn_nao = driver.find_element(By.ID, f'btnNaoResponder_{relative_index}')
                        txt_just = driver.find_element(By.ID, f'txtJustificativaRecusa_{relative_index}')
                        
                        if txt_just.is_displayed():
                            if btn_nao.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", btn_nao)
                                driver.execute_script("arguments[0].click();", btn_nao)
                            wait.until(EC.invisibility_of_element_located((By.ID, f'txtJustificativaRecusa_{relative_index}')))
                    except Exception as e:
                        print(f"  - Aviso: Erro ao reativar item {i} (pode já estar ativo): {e}")

                    try:
                        # Moeda
                        try:
                            moeda_select = driver.find_element(By.XPATH, f'//*[@id="tdItem_{relative_index}"]//select[contains(., "BRL")]')
                            Select(moeda_select).select_by_value('BRL')
                        except: pass

                        # Preço
                        price_val = str(product_info['UNITARIO']).replace('.', ',')
                        p_field = driver.find_element(By.ID, f'Preco{relative_index}')
                        driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", p_field, price_val)

                        # Prazo
                        prazo = "3"
                        if custom_days:
                            prazo = custom_days
                        elif 'dd' in code_lower:
                            try:
                                prazo = [p for p in code_lower.split('-') if 'dd' in p][0].replace('dd','').strip()
                            except: pass
                        prazo_field = driver.find_element(By.ID, f'Prazo{relative_index}')
                        prazo_field.clear()
                        prazo_field.send_keys(prazo)
                        
                        # NBM (NCM)
                        try:
                            nbm_field = driver.find_element(By.ID, f'NBM{relative_index}')
                        except Exception:
                            try:
                                nbm_field = driver.find_element(By.XPATH, f'//*[@id="tdItem_{relative_index}"]//input[contains(@id, "NBM") or contains(@name, "NBM")]')
                            except Exception:
                                nbm_field = None
                        
                        if nbm_field:
                            nbm_field.clear()
                            nbm_field.send_keys(product_info['NCM'])
                        
                        # ICMS
                        icms_val = str(product_info['ICMS']).replace('%', '').strip()
                        icms_field = driver.find_element(By.ID, f'ICMS{relative_index}')
                        icms_field.clear()
                        icms_field.send_keys(icms_val)
                        
                        # Observação
                        try:
                            obs_field = None
                            # 1. Busca direta por name="Observacao{relative_index}"
                            obs_candidatos = driver.find_elements(By.NAME, f'Observacao{relative_index}')
                            if obs_candidatos:
                                obs_field = obs_candidatos[0]
                            
                            # 2. Busca dentro do container tdItem_{relative_index} por name
                            if not obs_field:
                                obs_candidatos = driver.find_elements(By.XPATH, f'//*[@id="tdItem_{relative_index}"]//input[@name="Observacao{relative_index}"]')
                                if obs_candidatos:
                                    obs_field = obs_candidatos[0]
                            
                            # 3. XPath exato fornecido pelo novo layout (td[5])
                            if not obs_field:
                                obs_candidatos = driver.find_elements(By.XPATH, f'//*[@id="tdItem_{relative_index}"]/table/tbody/tr[3]/td/table/tbody/tr[3]/td[5]/table/tbody/tr/td[2]/input')
                                if obs_candidatos:
                                    obs_field = obs_candidatos[0]
                            
                            # 4. Busca pelo label "Obs:" dentro do item
                            if not obs_field:
                                obs_candidatos = driver.find_elements(By.XPATH, f'//*[@id="tdItem_{relative_index}"]//td[contains(normalize-space(text()), "Obs")]/following::input[1]')
                                if obs_candidatos:
                                    obs_field = obs_candidatos[0]

                            # Validação de segurança: garantir que NÃO é o campo NBM
                            if obs_field:
                                field_id = obs_field.get_attribute("id") or ""
                                if field_id.startswith("NBM"):
                                    print(f"  - Aviso: Elemento encontrado para Observação coincidiu com NBM ({field_id}). Ignorando.")
                                    obs_field = None

                            if obs_field:
                                obs_field.clear()
                                obs_texto = str(product_info.get('OBSERVACAO', '')).strip()
                                obs_field.send_keys(obs_texto[:100])
                            else:
                                print(f"  - Aviso: Campo Observação não localizado para o item {i}.")
                        except Exception as e:
                            print(f"  - Aviso: Erro ao preencher Observação no item {i}. {e}")
                        
                        # Subst. Tributária (Robusto)
                        try:
                            item_container = driver.find_element(By.ID, f"tdItem_{relative_index}")
                            st_select_element = item_container.find_element(By.XPATH, ".//*[contains(text(), 'Subst')]/following::select[1]")
                            
                            select_obj = Select(st_select_element)
                            try: select_obj.select_by_visible_text('não')
                            except:
                                try: select_obj.select_by_visible_text('Não')
                                except: 
                                    st_select_element.click()
                                    st_select_element.send_keys("n")
                                    st_select_element.send_keys(Keys.TAB)

                            valor_st_field = driver.find_element(By.ID, f'ValorSubstituicaoTributaria{relative_index}')
                            valor_st_field.clear()
                            valor_st_field.send_keys("0")
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", valor_st_field)
                        except Exception as e:
                            print(f"  - Aviso: Erro ao preencher Subst. Tributária no item {i}. {e}")

                        # Unidade
                        selecionar_unidade_segura(driver, f'UnidadeResp{relative_index}', product_info['UMED'])

                        # IPI e Outros
                        ipi_field = driver.find_element(By.ID, f'IPI{relative_index}')
                        ipi_field.clear()
                        ipi_field.send_keys("0")
                        
                        Select(driver.find_element(By.ID, f'IPIIncluso{relative_index}')).select_by_value('N')
                        Select(driver.find_element(By.ID, f'ICMSIncluso{relative_index}')).select_by_value('S')
                        
                        lote_field = driver.find_element(By.ID, f'LoteMinimo{relative_index}')
                        lote_field.clear()
                        lote_field.send_keys("1")

                    except Exception as e:
                        print(f"Erro preenchendo item {i}: {e}")

                # Paginação
                if i % 10 == 0 and i < qtd_itens:
                    print(f"--- Fim da página. Indo para próxima... ---")
                    try:
                        elemento_pagina_atual = driver.find_element(By.ID, 'Preco1')
                        next_btn_img = driver.find_element(By.XPATH, '//img[contains(@src, "bot_pg_nxt.gif")]')
                        driver.execute_script("arguments[0].click();", next_btn_img.find_element(By.XPATH, './..'))
                        
                        try: WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
                        except: pass

                        wait.until(EC.staleness_of(elemento_pagina_atual))
                        wait.until(EC.visibility_of_element_located((By.ID, 'Preco1')))
                        print("Nova página carregada.")

                    except Exception as e_pag:
                        print(f"ERRO CRÍTICO NA PAGINAÇÃO: {e_pag}")
                        if not messagebox.askyesno("Erro Paginação", "Falha ao mudar de página. Continuar manualmente?"):
                            break
            
        except Exception as e_items:
            print(f"Erro geral no loop de itens: {e_items}")

        # --- Salva a RFQ no histórico (para depois localizar os códigos no pedido) ---
        chave_hist = ""
        if historico_rfq and itens_registro:
            try:
                chave_hist = historico_rfq.salvar_rfq({
                    'rfq': task.get('rfq', ''),
                    'cotacao_id': task.get('cotacao_id', ''),
                    'link': task.get('link', ''),
                    'portal': login_type,
                    'planilha': os.path.basename(rfq_path),
                    'linha_planilha': task['row_num'],
                    'texto_planilha': task.get('codes', ''),
                    'avisos_lote': task.get('lote_textos', []),
                    'status': 'preenchida',
                    'itens': itens_registro,
                })
                if chave_hist:
                    qtd_cod = sum(1 for it in itens_registro if it.get('codigo_lf'))
                    print(f"💾 RFQ {chave_hist} salva no histórico ({qtd_cod} itens com código).")
                else:
                    print("⚠️ RFQ sem número identificado — não foi salva no histórico.")
            except Exception as e_hist:
                print(f"⚠️ Erro ao salvar RFQ no histórico: {e_hist}")

        aviso_lote = "\n".join(task.get('lote_textos', [])) or None
        if aviso_lote:
            print(f"⚠️ LOTE informado na linha {task['row_num']}: {aviso_lote}")

        if single_mode:
            if aviso_lote:
                messagebox.showwarning("Atenção - Lote", f"A linha {task['row_num']} da planilha menciona LOTE. Verifique antes de confirmar:\n\n{aviso_lote}")
            messagebox.showinfo("Teste", "Teste finalizado.")
            break
        
        # Obter escolha do usuário para a próxima ação
        escolha = show_custom_choice_dialog("Ação para RFQ", f"RFQ da linha {task['row_num']} foi totalmente preenchida.\nEscolha o que deseja fazer:", aviso=aviso_lote)
        
        if escolha == "confirm_next":
            if historico_rfq and chave_hist:
                try:
                    historico_rfq.atualizar_status(chave_hist, 'confirmada')
                except Exception as e_hist:
                    print(f"⚠️ Erro ao atualizar status no histórico: {e_hist}")
            print("Confirmando RFQ no portal...")
            try:
                confirm_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "MEComponentManager_MEButton_6"))
                )
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});", confirm_btn)
                time.sleep(0.3)
                confirm_btn.click()
                print("   > Botão 'Confirmar' clicado.")
            except:
                try:
                    confirm_btn = driver.find_element(By.XPATH, '//*[@id="MEComponentManager_MEButton_6"]')
                    driver.execute_script("arguments[0].click();", confirm_btn)
                    print("   > Botão 'Confirmar' clicado via fallback.")
                except Exception as e_click:
                    print(f"   ⚠️ Não foi possível clicar no botão 'Confirmar': {e_click}")
            
            # Tratar popup de confirmação
            time.sleep(1.0)
            popup_sucesso = False
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                print(f"   > Alerta do navegador aceito: '{alert_text}'")
                popup_sucesso = True
            except:
                pass
            
            if not popup_sucesso:
                try:
                    ActionChains(driver).send_keys(Keys.ENTER).perform()
                    print("   > Tecla ENTER enviada para confirmação do popup.")
                except:
                    pass
            time.sleep(1.5) # Aguarda processamento do envio no portal
            
        elif escolha == "next":
            print("Pulando para a próxima RFQ sem confirmar/enviar.")
            
        else: # "cancel"
            print("Automação cancelada pelo usuário.")
            break

    messagebox.showinfo("Fim", "Processo finalizado.")

# --- Configuração da Interface Gráfica Tkinter ---
root = tk.Tk()
root.title("Automação de Cotações RFQ v8.0")
root.geometry("840x820")
root.minsize(800, 780)
root.configure(bg="#F1F5F9")

style = ttk.Style(root)
style.theme_use('clam')

# Estilos Globais TTK
style.configure(".", background="#F1F5F9", font=("Segoe UI", 9), foreground="#0F172A")
style.configure("TFrame", background="#F1F5F9")
style.configure("TLabelframe", background="#F1F5F9", bordercolor="#CBD5E1", lightcolor="#E2E8F0", darkcolor="#CBD5E1")
style.configure("TLabelframe.Label", background="#F1F5F9", foreground="#1E293B", font=("Segoe UI", 10, "bold"))
style.configure("TLabel", background="#F1F5F9", foreground="#0F172A")
style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#0F172A")
style.configure("TCombobox", fieldbackground="#FFFFFF")

# Container Principal com padding
main_frame = ttk.Frame(root, padding="15")
main_frame.pack(expand=True, fill='both')

# --- Header Bar ---
header_frame = tk.Frame(main_frame, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=15, pady=12)
header_frame.pack(fill='x', pady=(0, 10))

title_sub_frame = tk.Frame(header_frame, bg="#FFFFFF")
title_sub_frame.pack(side='left')

title_label = tk.Label(title_sub_frame, text="⚡ Automação de Cotações RFQ - Versão 2.8", font=("Segoe UI", 16, "bold"), fg="#0F172A", bg="#FFFFFF")
title_label.pack(anchor='w')

subtitle_label = tk.Label(title_sub_frame, text="Leitura Direta de Excel & PDF - Versão 2.8 (Histórico de RFQs + Aviso de Lote)", font=("Segoe UI", 9, "bold"), fg="#2563EB", bg="#FFFFFF")
subtitle_label.pack(anchor='w')

# Badge de Status Geral
lbl_status_badge = tk.Label(header_frame, text="⚪ AGUARDANDO ARQUIVOS", font=("Segoe UI", 8, "bold"), bg="#E2E8F0", fg="#475569", padx=10, pady=4)
lbl_status_badge.pack(side='right')

# --- Card 1: Seleção de Arquivos ---
file_frame = ttk.LabelFrame(main_frame, text=" 📁 1. Seleção de Arquivos (Excel ou PDF) ", padding="12")
file_frame.pack(fill='x', pady=(0, 8))

# Subframe RFQ File
rfq_row = tk.Frame(file_frame, bg="#F1F5F9")
rfq_row.pack(fill='x', pady=2)
btn_select_rfq = tk.Button(rfq_row, text="📄 Selecionar Planilha de RFQs (.xlsx)", command=select_rfq_file, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, cursor="hand2", activebackground="#1D4ED8", activeforeground="white")
btn_select_rfq.pack(side='left', padx=(0, 10))
label_rfq_path = tk.Label(rfq_row, text="🔴 Nenhuma planilha de RFQs selecionada", font=("Segoe UI", 9, "italic"), fg="#64748B", bg="#F1F5F9")
label_rfq_path.pack(side='left', fill='x', expand=True)

# Subframe Product File
prod_row = tk.Frame(file_frame, bg="#F1F5F9")
prod_row.pack(fill='x', pady=(6, 2))
btn_select_product = tk.Button(prod_row, text="📊 Selecionar Produtos (Excel ou PDF)", command=select_product_file, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, cursor="hand2", activebackground="#1D4ED8", activeforeground="white")
btn_select_product.pack(side='left', padx=(0, 10))
label_product_path = tk.Label(prod_row, text="🔴 Nenhuma planilha ou PDF de produtos selecionado", font=("Segoe UI", 9, "italic"), fg="#64748B", bg="#F1F5F9")
label_product_path.pack(side='left', fill='x', expand=True)

# --- Grid 2 Colunas para Configurações e Credenciais ---
middle_grid = tk.Frame(main_frame, bg="#F1F5F9")
middle_grid.pack(fill='x', pady=(0, 8))
middle_grid.columnconfigure(0, weight=1)
middle_grid.columnconfigure(1, weight=1)
# Left Column: Configurações da Planilha
config_frame = ttk.LabelFrame(middle_grid, text=" ⚙️ 2. Configurações da Planilha ", padding="12")
config_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

ttk.Label(config_frame, text="Coluna Links (Letra):").grid(row=0, column=0, padx=4, pady=4, sticky='w')
col_link_var = tk.StringVar(value="A")
entry_col_link = ttk.Entry(config_frame, textvariable=col_link_var, width=6)
entry_col_link.grid(row=0, column=1, padx=4, pady=4, sticky='w')

ttk.Label(config_frame, text="Coluna Códigos (Letra):").grid(row=1, column=0, padx=4, pady=4, sticky='w')
col_code_var = tk.StringVar(value="I")
entry_col_code = ttk.Entry(config_frame, textvariable=col_code_var, width=6)
entry_col_code.grid(row=1, column=1, padx=4, pady=4, sticky='w')

ttk.Label(config_frame, text="Linha Inicial na Planilha:").grid(row=2, column=0, padx=4, pady=4, sticky='w')
row_start_var = tk.StringVar(value="1")
entry_row_start = ttk.Entry(config_frame, textvariable=row_start_var, width=6)
entry_row_start.grid(row=2, column=1, padx=4, pady=4, sticky='w')

# Right Column: Login
login_frame = ttk.LabelFrame(middle_grid, text=" 🔑 3. Credenciais do Portal ", padding="12")
login_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
login_frame.columnconfigure(1, weight=1)

ttk.Label(login_frame, text="Tipo de Login:").grid(row=0, column=0, padx=4, pady=4, sticky="w")
login_type_combo = ttk.Combobox(login_frame, values=list(LOGIN_CREDENTIALS.keys()), state="readonly")
login_type_combo.grid(row=0, column=1, padx=4, pady=4, sticky="ew")
login_type_combo.bind("<<ComboboxSelected>>", on_login_type_change)

user_var = tk.StringVar()
pass_var = tk.StringVar()
ttk.Label(login_frame, text="Usuário:").grid(row=1, column=0, padx=4, pady=4, sticky="w")
user_entry = ttk.Entry(login_frame, textvariable=user_var)
user_entry.grid(row=1, column=1, padx=4, pady=4, sticky="ew")
ttk.Label(login_frame, text="Senha:").grid(row=2, column=0, padx=4, pady=4, sticky="w")
pass_entry = ttk.Entry(login_frame, textvariable=pass_var, show="*")
pass_entry.grid(row=2, column=1, padx=4, pady=4, sticky="ew")

login_type_combo.current(0)
on_login_type_change(None)

# --- Card 3: Execução da Automação ---
action_frame = ttk.LabelFrame(main_frame, text=" 🚀 4. Executar Automação ", padding="12")
action_frame.pack(fill='x', pady=(0, 8))

single_row_frame = tk.Frame(action_frame, bg="#F1F5F9")
single_row_frame.pack(fill='x', pady=(0, 8))
ttk.Label(single_row_frame, text="Nº da Linha para Teste:").pack(side='left', padx=(0, 6))
row_test_var = tk.StringVar()
row_test_entry = ttk.Entry(single_row_frame, textvariable=row_test_var, width=10)
row_test_entry.pack(side='left', padx=(0, 10))

btn_start_single = tk.Button(single_row_frame, text="🧪 Testar Apenas esta Linha", command=lambda: start_automation_thread(single_mode=True), state=tk.DISABLED, bg="#475569", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2", activebackground="#334155", activeforeground="white")
btn_start_single.pack(side='left')

btn_start = tk.Button(action_frame, text="▶  INICIAR AUTOMAÇÃO COMPLETA", command=lambda: start_automation_thread(single_mode=False), state=tk.DISABLED, bg="#16A34A", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", pady=8, cursor="hand2", activebackground="#15803D", activeforeground="white")
btn_start.pack(fill='x')

# --- Card 4: Frame de Logs ---
log_frame = ttk.LabelFrame(main_frame, text=" 📋 5. Console de Logs ", padding="10")
log_frame.pack(fill='both', expand=True)

log_header_frame = tk.Frame(log_frame, bg="#F1F5F9")
log_header_frame.pack(fill='x', pady=(0, 4))
btn_clear_log = tk.Button(log_header_frame, text="🧹 Limpar Logs", command=clear_log, bg="#E2E8F0", fg="#334155", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=2, cursor="hand2", activebackground="#CBD5E1")
btn_clear_log.pack(side='right')

log_container = tk.Frame(log_frame, bg="#0F172A")
log_container.pack(fill='both', expand=True)

log_text = tk.Text(log_container, height=9, state=tk.DISABLED, wrap=tk.WORD, bg="#0F172A", fg="#38BDF8", font=("Consolas", 9), relief="flat", padx=8, pady=8)
log_text.pack(side=tk.LEFT, fill='both', expand=True)

scrollbar = ttk.Scrollbar(log_container, command=log_text.yview)
scrollbar.pack(side=tk.RIGHT, fill='y')
log_text.config(yscrollcommand=scrollbar.set)

root.mainloop()