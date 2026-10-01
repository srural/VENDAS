import os
import json
import time
import datetime
import threading

LOGS_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'logs'))
LOGS_FILE = os.path.join(LOGS_DIR, 'sefaz_transmissions.json')
MAX_LOG_ENTRIES = 1000

_lock = threading.Lock()

def _ensure_logs_dir():
    os.makedirs(LOGS_DIR, exist_ok=True)
    if not os.path.exists(LOGS_FILE):
        with open(LOGS_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f)

def log_sefaz_transmission(
    tipo_doc="NFE",          # 'NFE' ou 'NFCE'
    modelo="55",             # '55' ou '65'
    ambiente=1,              # 1=Producao, 2=Homologacao
    servico="Autorizacao",   # 'Autorizacao', 'RetAutorizacao', 'StatusServico', 'Inutilizacao', 'RecepcaoEvento'
    numero_doc=None,         # Numero da NFe / NFCe / Pedido
    chave_nfe=None,          # Chave de 44 digitos
    url=None,                # WebService URL
    status_http=200,         # Status code HTTP
    c_stat="",               # cStat retornado pela SEFAZ
    x_motivo="",             # xMotivo retornado
    n_prot="",               # Protocolo de autorizacao
    tempo_ms=0,              # Duracao da chamada em ms
    erro=None,               # Mensagem de erro se houver
    request_xml=None,        # XML enviado
    response_xml=None        # XML recebido
):
    """
    Registra uma comunicação com os WebServices da SEFAZ para auditoria e diagnósticos no painel.
    """
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")
    amb_int = int(ambiente) if str(ambiente) in ("1", "2") else 2
    amb_nome = "Produção" if amb_int == 1 else "Homologação"
    c_stat_str = str(c_stat or "").strip()
    is_autorizada = c_stat_str in ("100", "150")

    entry = {
        "id": f"log_{int(time.time() * 1000)}",
        "timestamp": now_str,
        "tipo_doc": str(tipo_doc or "NFE").upper(),
        "modelo": str(modelo or "55"),
        "ambiente": amb_int,
        "ambiente_nome": amb_nome,
        "servico": str(servico or "Autorizacao"),
        "numero_doc": str(numero_doc or ""),
        "chave_nfe": str(chave_nfe or ""),
        "url": str(url or ""),
        "status_http": int(status_http or 0),
        "c_stat": c_stat_str or ("200" if is_autorizada else ("999" if erro else "---")),
        "x_motivo": str(x_motivo or erro or "Sem mensagem"),
        "n_prot": str(n_prot or ""),
        "sucesso": is_autorizada or (status_http == 200 and not erro and c_stat_str in ("100", "150", "107")),
        "tempo_ms": int(tempo_ms or 0),
        "erro": str(erro) if erro else None,
        "request_xml": str(request_xml or ""),
        "response_xml": str(response_xml or "")
    }

    with _lock:
        try:
            _ensure_logs_dir()
            data = []
            if os.path.exists(LOGS_FILE):
                try:
                    with open(LOGS_FILE, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except Exception:
                    data = []

            # Inserir no início (ordem cronológica decrescente)
            data.insert(0, entry)
            if len(data) > MAX_LOG_ENTRIES:
                data = data[:MAX_LOG_ENTRIES]

            with open(LOGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[SEFAZ LOGGER ERROR] Erro ao gravar log: {e}")

    return entry

def get_sefaz_logs(tipo_doc=None, modelo=None, ambiente=None, c_stat=None, search=None, limit=100, offset=0):
    """
    Retorna logs da SEFAZ filtrados com suporte a paginação e busca textual.
    """
    with _lock:
        _ensure_logs_dir()
        try:
            with open(LOGS_FILE, 'r', encoding='utf-8') as f:
                logs = json.load(f)
        except Exception:
            logs = []

    filtered = []
    clean_search = str(search or "").strip().lower()

    for item in logs:
        if tipo_doc and str(tipo_doc).upper() != "ALL" and item.get("tipo_doc") != str(tipo_doc).upper():
            continue
        if modelo and str(modelo) != "ALL" and item.get("modelo") != str(modelo):
            continue
        if ambiente and str(ambiente) not in ("ALL", ""):
            if str(item.get("ambiente")) != str(ambiente):
                continue
        if c_stat and str(c_stat) not in ("ALL", ""):
            if c_stat == "OK" and not item.get("sucesso"):
                continue
            elif c_stat == "ERR" and item.get("sucesso"):
                continue
            elif c_stat not in ("OK", "ERR") and item.get("c_stat") != str(c_stat):
                continue

        if clean_search:
            match_fields = [
                item.get("chave_nfe", ""),
                item.get("numero_doc", ""),
                item.get("c_stat", ""),
                item.get("x_motivo", ""),
                item.get("n_prot", ""),
                item.get("servico", "")
            ]
            if not any(clean_search in str(val).lower() for val in match_fields):
                continue

        filtered.append(item)

    total = len(filtered)
    paged = filtered[offset:offset + limit]

    return {
        "success": True,
        "total": total,
        "logs": paged
    }

def clear_sefaz_logs():
    """
    Limpa o arquivo de logs da SEFAZ.
    """
    with _lock:
        _ensure_logs_dir()
        with open(LOGS_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f)
    return {"success": True, "message": "Logs da SEFAZ limpos com sucesso"}
