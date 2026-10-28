"""
importar_rfqs_antigas.py

Utilitário opcional para importar planilhas antigas de RFQ e associar com produtos/PDF
para preencher o arquivo historico_rfq.json retrospectivamente.
"""
import os
import re
import glob
import openpyxl
import pandas as pd
import historico_rfq

def extrair_dados_rfq_antiga(caminho_rfq):
    wb = openpyxl.load_workbook(caminho_rfq, data_only=False)
    ws = wb.active
    
    registros = []
    
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=False), start=1):
        if not row:
            continue
        vals = [str(c.value) if c.value is not None else '' for c in row]
        
        # Link na coluna A
        link_cell = row[0]
        link = link_cell.hyperlink.target if link_cell.hyperlink else (str(link_cell.value) if str(link_cell.value).startswith('http') else '')
        
        # Código na coluna I (índice 8)
        code_val = str(row[8].value or '') if len(row) > 8 else ''
        
        # RFQ na linha
        rfq_num = ''
        for v in vals:
            m = re.search(r'RFQ\s*:?\s*(\d{4,})', v, re.I)
            if m:
                rfq_num = m.group(1)
                break
                
        if not link and not rfq_num:
            continue
            
        m_cot = re.search(r'Cot=(\d+)', link, re.I)
        cot_id = m_cot.group(1) if m_cot else ''
        
        # Textos de lote
        lotes = [ln.strip() for v in vals for ln in v.split('\n') if 'lote' in ln.lower()]
        
        registros.append({
            'linha': r_idx,
            'rfq': rfq_num,
            'cotacao_id': cot_id,
            'link': link,
            'codes_raw': code_val,
            'lotes': lotes
        })
        
    return registros

def main():
    print("=== Importador de Planilhas Antigas de RFQ ===")
    planilhas_rfq = glob.glob("WEG DI*.xlsx")
    if not planilhas_rfq:
        print("Nenhuma planilha 'WEG DI*.xlsx' encontrada.")
        return
        
    total_importadas = 0
    for p in planilhas_rfq:
        nome = os.path.basename(p)
        print(f"\nLendo: {nome}")
        regs = extrair_dados_rfq_antiga(p)
        print(f"  > {len(regs)} linhas com RFQ encontradas.")
        
        for r in regs:
            # Processa códigos
            lines = [c.strip() for c in r['codes_raw'].replace('\r', '\n').split('\n') if c.strip()]
            itens = []
            for idx, c in enumerate(lines, start=1):
                c_clean = c.replace('.0', '').strip()
                part0 = c_clean.split('-')[0].split(' ')[0].strip()
                clean = re.sub(r'[^\d]', '', part0)
                if len(clean) < 6 and clean.isdigit():
                    clean = clean.zfill(6)
                itens.append({
                    'item_rfq': idx,
                    'codigo_planilha': c,
                    'codigo_lf': clean if clean else None,
                    'declinado': False if clean else True,
                    'preco': '',
                    'descricao_lf': '',
                    'texto_rfq': c,
                    'nota_lote': ' | '.join(r['lotes']) if r['lotes'] else ''
                })
                
            chave = historico_rfq.salvar_rfq({
                'rfq': r['rfq'],
                'cotacao_id': r['cotacao_id'],
                'link': r['link'],
                'portal': 'Dois Irmãos (DI)',
                'planilha': nome,
                'linha_planilha': r['linha'],
                'texto_planilha': r['codes_raw'],
                'avisos_lote': r['lotes'],
                'status': 'importada_antiga',
                'itens': itens
            })
            if chave:
                total_importadas += 1
                
    print(f"\n[OK] Concluido! {total_importadas} RFQs salvas no historico ({historico_rfq.HISTORICO_PATH}).")

if __name__ == '__main__':
    main()
