import os
import configparser
import psycopg2
from psycopg2.extras import RealDictCursor

PG_HOST = os.environ.get("PG_HOST", "localhost")
PG_PORT = int(os.environ.get("PG_PORT", 5432))
PG_USER = os.environ.get("PG_USER", "postgres")
PG_PASS = os.environ.get("PG_PASS", "masterkey")
PG_DB = os.environ.get("PG_DB", "vendas_db")

CONFIG_INI_PATH = os.path.join(os.path.dirname(__file__), '..', 'VB', 'Frente', 'config.ini')

def get_db_connection():
    return psycopg2.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASS,
        dbname=PG_DB,
        cursor_factory=RealDictCursor
    )

def init_cfgpdv_table():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS "CfgPdv" (
                    "Secao" TEXT NOT NULL,
                    "Chave" TEXT NOT NULL,
                    "Valor" TEXT,
                    PRIMARY KEY ("Secao", "Chave")
                )
            """)
            conn.commit()

            cursor.execute('ALTER TABLE "CfgPdv" ADD COLUMN IF NOT EXISTS "NroNfe" TEXT;')
            cursor.execute('ALTER TABLE "CfgPdv" ADD COLUMN IF NOT EXISTS "SerieNfe" TEXT;')
            cursor.execute('ALTER TABLE "CfgPdv" ADD COLUMN IF NOT EXISTS "id_empresa" BIGINT DEFAULT 1;')
            cursor.execute('UPDATE "CfgPdv" SET "id_empresa" = 1 WHERE "id_empresa" IS NULL;')
            conn.commit()

            cursor.execute('SELECT COUNT(*) as cnt FROM "CfgPdv"')
            count = cursor.fetchone()['cnt']

            if count == 0 and os.path.exists(CONFIG_INI_PATH):
                config = configparser.ConfigParser()
                config.optionxform = str
                config.read(CONFIG_INI_PATH, encoding='latin1')

                to_insert = []
                for section in config.sections():
                    for key, val in config.items(section):
                        to_insert.append((section, key, val))

                for item in to_insert:
                    cursor.execute(
                        'INSERT INTO "CfgPdv" ("Secao", "Chave", "Valor", "id_empresa") VALUES (%s, %s, %s, 1) ON CONFLICT ("Secao", "Chave") DO UPDATE SET "Valor" = EXCLUDED."Valor"',
                        item
                    )
                conn.commit()
    finally:
        conn.close()

def get_pdv_config(id_empresa=None):
    init_cfgpdv_table()
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            if id_empresa is not None:
                cursor.execute('SELECT "Secao", "Chave", "Valor" FROM "CfgPdv" WHERE "id_empresa" = %s OR "id_empresa" IS NULL', (id_empresa,))
            else:
                cursor.execute('SELECT "Secao", "Chave", "Valor" FROM "CfgPdv"')
            rows = cursor.fetchall()
    finally:
        conn.close()

    cfg = {
        "Config": {},
        "Sat": {},
        "Nfce": {},
        "Nfe": {},
        "Certificado": {},
        "boleto": {}
    }

    for r in rows:
        sec = r['Secao']
        key = r['Chave']
        val = r['Valor'] if r['Valor'] is not None else ''
        if sec not in cfg:
            cfg[sec] = {}
        cfg[sec][key] = val

    defaults = {
        "Config": {
            "Porta": "TELA", "PortaNf": "c:\\lixo.txt", "Leitor": "TECLADO",
            "LocalEst": "LOJA", "BaixaAuto": "1", "MostraEstoque": "1",
            "MostraCusto": "0", "Desconto": "0", "Copias": "2", "QtdLinhas": "7", "Mercado": "1"
        },
        "Sat": {
            "NroCaixa": "1", "LocalAcbr": "c:\\AcbrMonitorPlus",
            "CnpjDesenvolvedor": "23103347000165", "Versao": "0.08"
        },
        "Nfce": {
            "CodigoID": "000001", "CodigoSeg": "1317bbaf-264c-41d9-9c71-c75215211f2a",
            "SerieNfce": "1", "NroNfce": "16", "Ambiente": "2", "Chave": ""
        },
        "Nfe": {
            "SerieNfe": "1", "NroNfe": "1", "Ambiente": "2", "Chave": ""
        },
        "Certificado": {
            "Caminho": "C:\\Certificados\\certificado.pfx", "Senha": "",
            "Tipo": "A1", "Validade": "2027-12-31", "CNPJCertificado": "11054174000153"
        },
        "boleto": {
            "Banco": "", "AGENCIA": "", "CONTA": "", "Convenio": ""
        }
    }

    for sec_name, sec_keys in defaults.items():
        if sec_name not in cfg:
            cfg[sec_name] = {}
        for k, v in sec_keys.items():
            if k not in cfg[sec_name]:
                cfg[sec_name][k] = v

    return cfg

def save_pdv_config(data):
    init_cfgpdv_table()
    id_emp = data.get("id_empresa")
    if id_emp is None:
        id_emp = 1
    else:
        try:
            id_emp = int(id_emp)
        except (ValueError, TypeError):
            id_emp = 1

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            for sec_name, sec_data in data.items():
                if sec_name == "id_empresa":
                    continue
                if isinstance(sec_data, dict):
                    for key, val in sec_data.items():
                        cursor.execute("""
                            INSERT INTO "CfgPdv" ("Secao", "Chave", "Valor", "id_empresa")
                            VALUES (%s, %s, %s, %s)
                            ON CONFLICT ("Secao", "Chave") DO UPDATE SET "Valor" = EXCLUDED."Valor", "id_empresa" = EXCLUDED."id_empresa"
                        """, (sec_name, key, str(val), id_emp))

            # Atualiza também as colunas explícitas SerieNfe e NroNfe caso venham na seção Nfe
            nfe_sec = data.get("Nfe", {})
            if isinstance(nfe_sec, dict):
                serie_val = nfe_sec.get("SerieNfe")
                nro_val = nfe_sec.get("NroNfe")
                if serie_val is not None:
                    cursor.execute('UPDATE "CfgPdv" SET "SerieNfe" = %s, "id_empresa" = %s WHERE "Secao" = \'Nfe\'', (str(serie_val), id_emp))
                if nro_val is not None:
                    cursor.execute('UPDATE "CfgPdv" SET "NroNfe" = %s, "id_empresa" = %s WHERE "Secao" = \'Nfe\'', (str(nro_val), id_emp))
            conn.commit()
    finally:
        conn.close()

    return get_pdv_config(id_empresa=id_emp)


def get_and_increment_nfe_number(id_empresa=1):
    """
    Obtém a SerieNfe e NroNfe atuais da tabela CfgPdv para a NF-e Modelo 55,
    retorna a série e o número atual para uso na nota, e incrementa o NroNfe em +1 na CfgPdv.
    """
    cfg = get_pdv_config(id_empresa)
    nfe_sec = cfg.get("Nfe", {})
    
    serie = str(nfe_sec.get("SerieNfe") or nfe_sec.get("Serie") or "1").strip()
    nro_str = str(nfe_sec.get("NroNfe") or nfe_sec.get("Nro") or "1").strip()
    
    try:
        nro_atual = int(nro_str)
    except (ValueError, TypeError):
        nro_atual = 1

    nro_proximo = nro_atual + 1
    
    save_pdv_config({
        "id_empresa": id_empresa,
        "Nfe": {
            "SerieNfe": serie,
            "NroNfe": str(nro_proximo)
        }
    })
    
    return serie, nro_atual


def get_and_increment_nfce_number(id_empresa=1):
    """
    Obtém SerieNfce, NroNfce, CodigoID (cIdToken) e CodigoSeg (CSC) da tabela CfgPdv para a NFC-e Modelo 65,
    retorna os dados atuais para a nota e QR Code v2.00, e incrementa o NroNfce em +1 na CfgPdv.
    """
    cfg = get_pdv_config(id_empresa)
    nfce_sec = cfg.get("Nfce", {})
    
    serie = str(nfce_sec.get("SerieNfce") or nfce_sec.get("Serie") or "1").strip()
    nro_str = str(nfce_sec.get("NroNfce") or nfce_sec.get("Nro") or "1").strip()
    codigo_id = str(nfce_sec.get("CodigoID") or nfce_sec.get("cIdToken") or "000001").strip()
    codigo_seg = str(nfce_sec.get("CodigoSeg") or nfce_sec.get("csc") or "1317bbaf-264c-41d9-9c71-c75215211f2a").strip()

    try:
        nro_atual = int(nro_str)
    except (ValueError, TypeError):
        nro_atual = 1

    nro_proximo = nro_atual + 1
    
    save_pdv_config({
        "id_empresa": id_empresa,
        "Nfce": {
            "SerieNfce": serie,
            "NroNfce": str(nro_proximo),
            "CodigoID": codigo_id,
            "CodigoSeg": codigo_seg
        }
    })
    
    return serie, nro_atual, codigo_id, codigo_seg

