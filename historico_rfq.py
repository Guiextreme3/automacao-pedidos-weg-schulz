"""
Histórico de RFQs respondidas.

- O RFQ.py grava aqui cada RFQ preenchida (código LF, preço, descrição, texto do item no portal...).
- O WEG Pedidos Novo.py consulta este histórico pelo número da RFQ que aparece no pedido e
  tenta descobrir qual código LF corresponde a cada item do pedido (por código WEG, preço e descrição).

Os dados ficam em 'historico_rfq.json', na mesma pasta dos scripts.
"""
import os
import sys
import re
import json
import shutil
import threading
import unicodedata
from datetime import datetime


def _base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


HISTORICO_PATH = os.path.join(_base_dir(), "historico_rfq.json")
_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------
def normalizar_rfq(valor):
    """Extrai só o número da RFQ. Ex.: 'RFQ: 1093929' -> '1093929'."""
    if valor is None:
        return ""
    s = str(valor)
    m = re.search(r'RFQ\D{0,5}(\d{4,})', s, re.I)
    if not m:
        m = re.search(r'(\d{4,})', s)
    if not m:
        return ""
    num = m.group(1).lstrip('0')
    return num or m.group(1)


def para_float(valor):
    """Converte '1.234,56', '1234.56', '12,5' etc. para float. Retorna None se não conseguir."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = re.sub(r'[^\d,.]', '', str(valor))
    if not s:
        return None
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    elif s.count('.') > 1:
        # Ex.: '1.234.56' (vírgula já trocada por ponto) -> mantém só o último ponto como decimal
        s = s.replace('.', '', s.count('.') - 1)
    try:
        return float(s)
    except ValueError:
        return None


_STOPWORDS = {
    "DE", "DA", "DO", "DAS", "DOS", "COM", "SEM", "PARA", "POR", "EM", "NA", "NO", "E", "A", "O",
    "UN", "UND", "PC", "PCS", "PECA", "PECAS", "CX", "KG", "MT", "M", "MM", "TIPO", "REF", "MARCA",
    "DESC", "COD", "CODIGO", "ITEM", "QTD", "QUANTIDADE", "PRECO", "UNITARIO", "PRAZO", "OBS",
}


def _tokens(texto):
    t = unicodedata.normalize('NFKD', str(texto or "")).encode('ascii', 'ignore').decode().upper()
    return {tk for tk in re.findall(r'[A-Z0-9]+', t) if len(tk) >= 2 and tk not in _STOPWORDS}


def similaridade_descricao(desc_pedido, texto_candidato):
    """Fração (0 a 1) das palavras da descrição do pedido que aparecem no candidato.
    Aceita abreviações (ex.: 'ROLAM' casa com 'ROLAMENTO')."""
    tp = _tokens(desc_pedido)
    tc = _tokens(texto_candidato)
    if not tp or not tc:
        return 0.0
    acertos = 0
    for a in tp:
        if a in tc:
            acertos += 1
        elif len(a) >= 3 and any((b.startswith(a) or a.startswith(b)) for b in tc if len(b) >= 3):
            acertos += 1
    return acertos / len(tp)


def _codigo_aparece(codigo, texto):
    cod = re.sub(r'\D', '', str(codigo or "")).lstrip('0')
    if len(cod) < 4:
        return False
    return re.search(r'(?<!\d)0*' + re.escape(cod) + r'(?!\d)', str(texto or "")) is not None


# ---------------------------------------------------------------------------
# Leitura / gravação
# ---------------------------------------------------------------------------
def carregar_historico():
    if not os.path.exists(HISTORICO_PATH):
        return {}
    with open(HISTORICO_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data if isinstance(data, dict) else {}


def _carregar_para_gravar():
    """Carrega o histórico; se o arquivo estiver corrompido, faz backup antes para não perder nada."""
    try:
        return carregar_historico()
    except Exception:
        backup = HISTORICO_PATH + f".corrompido_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        try:
            shutil.copy2(HISTORICO_PATH, backup)
            print(f"⚠️ historico_rfq.json estava inválido. Backup salvo em: {backup}")
        except Exception:
            pass
        return {}


def _gravar(data):
    tmp = HISTORICO_PATH + ".tmp"
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, HISTORICO_PATH)


def chave_registro(rfq_numero, cotacao_id=""):
    n = normalizar_rfq(rfq_numero)
    if n:
        return n
    return f"COT-{cotacao_id}" if cotacao_id else ""


def salvar_rfq(registro):
    """Grava/atualiza uma RFQ no histórico. Retorna a chave usada (ou '' se não deu pra identificar)."""
    chave = chave_registro(registro.get("rfq"), registro.get("cotacao_id"))
    if not chave:
        return ""
    registro["rfq"] = normalizar_rfq(registro.get("rfq"))
    registro["atualizado_em"] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with _lock:
        data = _carregar_para_gravar()
        anterior = data.get(chave) or {}
        if anterior.get("status") == "confirmada" and registro.get("status") != "confirmada":
            # Não "rebaixa" uma RFQ que já foi confirmada antes
            registro["status"] = "confirmada"
        registro.setdefault("criado_em", anterior.get("criado_em") or registro["atualizado_em"])
        data[chave] = registro
        _gravar(data)
    return chave


def atualizar_status(chave, status):
    if not chave:
        return
    with _lock:
        data = _carregar_para_gravar()
        if chave in data:
            data[chave]["status"] = status
            data[chave]["atualizado_em"] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            _gravar(data)


def buscar_rfq(rfq_numero):
    n = normalizar_rfq(rfq_numero)
    if not n:
        return None
    try:
        data = carregar_historico()
    except Exception as e:
        print(f"⚠️ Não foi possível ler o historico_rfq.json: {e}")
        return None
    if n in data:
        return data[n]
    for reg in data.values():
        if normalizar_rfq(reg.get("rfq")) == n:
            return reg
    return None


# ---------------------------------------------------------------------------
# Casamento itens do pedido x itens da RFQ
# ---------------------------------------------------------------------------
def casar_itens_pedido(itens_pedido, registro_rfq):
    """
    Para cada item do pedido, procura o item da RFQ que mais encaixa.

    Retorna uma lista (mesma ordem de itens_pedido) de dicts:
      {
        "status": "auto" | "duvida" | "sem_candidato",
        "escolhido": <item da RFQ> ou None,
        "candidatos": [ {"item": <item da RFQ>, "score": float, "preco_ok": bool,
                         "codigo_ok": bool, "similaridade": float}, ... ]  (melhor primeiro)
      }
    """
    candidatos_rfq = [c for c in (registro_rfq or {}).get("itens", []) if c.get("codigo_lf")]
    resultados = []

    for item in itens_pedido:
        preco_ped = para_float(item.get("preco"))
        cod_weg_ped = item.get("codigo_produto", "")
        desc_ped = item.get("descricao", "")

        avaliados = []
        for c in candidatos_rfq:
            preco_c = para_float(c.get("preco"))
            preco_ok = bool(preco_ped and preco_c and abs(preco_ped - preco_c) <= max(0.011, 0.005 * preco_c))
            preco_perto = bool(preco_ped and preco_c and abs(preco_ped - preco_c) <= 0.05 * preco_c)
            texto_c = " ".join(str(c.get(k, "")) for k in ("texto_rfq", "codigo_weg", "descricao_lf"))
            codigo_ok = _codigo_aparece(cod_weg_ped, " ".join(str(c.get(k, "")) for k in ("texto_rfq", "codigo_weg")))
            sim = similaridade_descricao(desc_ped, texto_c)
            score = (50 if codigo_ok else 0) + (35 if preco_ok else (15 if preco_perto else 0)) + 30 * sim
            avaliados.append({"item": c, "score": score, "preco_ok": preco_ok, "preco_perto": preco_perto,
                              "codigo_ok": codigo_ok, "similaridade": sim})

        avaliados.sort(key=lambda x: x["score"], reverse=True)

        status, escolhido = "sem_candidato", None
        if avaliados:
            status = "duvida"
            escolhido = avaliados[0]["item"] if avaliados[0]["score"] > 0 else None

            por_codigo = [a for a in avaliados if a["codigo_ok"]]
            por_preco = [a for a in avaliados if a["preco_ok"]]

            if len(por_codigo) == 1 and (por_codigo[0]["preco_perto"] or por_codigo[0]["similaridade"] >= 0.3):
                status, escolhido = "auto", por_codigo[0]["item"]
            elif len(por_preco) == 1 and (por_preco[0]["similaridade"] >= 0.35 or len(avaliados) == 1):
                status, escolhido = "auto", por_preco[0]["item"]
            elif len(por_preco) > 1:
                ordenados = sorted(por_preco, key=lambda a: a["similaridade"], reverse=True)
                if ordenados[0]["similaridade"] >= 0.5 and ordenados[0]["similaridade"] - ordenados[1]["similaridade"] >= 0.25:
                    status, escolhido = "auto", ordenados[0]["item"]

        resultados.append({"status": status, "escolhido": escolhido, "candidatos": avaliados})

    # Se o mesmo item da RFQ foi escolhido automaticamente para mais de um item do pedido, pede confirmação
    usados = {}
    for r in resultados:
        if r["status"] == "auto":
            usados.setdefault(id(r["escolhido"]), []).append(r)
    for lista in usados.values():
        if len(lista) > 1:
            for r in lista:
                r["status"] = "duvida"

    return resultados
