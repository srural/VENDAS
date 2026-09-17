import os
import configparser
from app.database import get_connection

NFE_INI_PATH = os.path.join(os.path.dirname(__file__), '..', 'VB', 'nfe.ini')

def normalize_empresa_row(row):
    if not row:
        return None
    d = dict(row)
    norm = {k.lower(): v for k, v in d.items()}

    keys_map = {
        'id_empresa': ['id_empresa', 'id', 'idempresa'],
        'RazaoSocial': ['RazaoSocial', 'razaosocial', 'NomeEmpresa', 'nomeempresa', 'nome', 'razao'],
        'Fantasia': ['Fantasia', 'fantasia', 'NomeFantasia', 'nomefantasia'],
        'CNPJ': ['CNPJ', 'cnpj', 'CGC', 'cgc'],
        'InscEst': ['InscEst', 'inscest', 'IE', 'ie'],
        'Logradouro': ['Logradouro', 'logradouro', 'Endereco', 'endereco', 'rua'],
        'Nro': ['Nro', 'nro', 'Numero', 'numero'],
        'Bairro': ['Bairro', 'bairro'],
        'Cidade': ['Cidade', 'cidade', 'Municipio', 'municipio'],
        'UF': ['UF', 'uf', 'Estado', 'estado'],
        'CEP': ['CEP', 'cep'],
        'Fone': ['Fone', 'fone', 'Telefone', 'telefone'],
        'CodigoIBGE': ['CodigoIBGE', 'codigoibge', 'CodIbge', 'codibge', 'ibge'],
        'RegimeTrib': ['RegimeTrib', 'regimetrib', 'Regime', 'regime', 'crt'],
        'PIS': ['PIS', 'pis'],
        'AliqPIS': ['AliqPIS', 'aliqpis'],
        'COFINS': ['COFINS', 'cofins'],
        'AliqCOFINS': ['AliqCOFINS', 'aliqcofins'],
        'SitTrib': ['SitTrib', 'sittrib'],
        'CFOP': ['CFOP', 'cfop'],
        'SitTribFixo': ['SitTribFixo', 'sittribfixo'],
        'Deducao': ['Deducao', 'deducao'],
        'Ativo': ['Ativo', 'ativo'],
        'DtCadastro': ['DtCadastro', 'dtcadastro']
    }
    result = {}
    for std_key, aliases in keys_map.items():
        val = None
        for a in aliases:
            if a in d:
                val = d[a]
                break
            elif a.lower() in norm:
                val = norm[a.lower()]
                break
        result[std_key] = val
    
    if result.get('id_empresa') is None and 'id_empresa' in d:
        result['id_empresa'] = d['id_empresa']
    return result

def init_empresa_table():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS "EMP" (
                    "id_empresa" BIGINT PRIMARY KEY,
                    "RazaoSocial" TEXT NOT NULL,
                    "Fantasia" TEXT,
                    "CNPJ" TEXT,
                    "InscEst" TEXT,
                    "Logradouro" TEXT,
                    "Nro" TEXT,
                    "Bairro" TEXT,
                    "Cidade" TEXT,
                    "UF" VARCHAR(2),
                    "CEP" TEXT,
                    "Fone" TEXT,
                    "CodigoIBGE" TEXT,
                    "RegimeTrib" TEXT,
                    "PIS" TEXT,
                    "AliqPIS" TEXT,
                    "COFINS" TEXT,
                    "AliqCOFINS" TEXT,
                    "SitTrib" TEXT,
                    "CFOP" TEXT,
                    "SitTribFixo" TEXT,
                    "Deducao" TEXT,
                    "Ativo" INTEGER DEFAULT 1,
                    "DtCadastro" TEXT
                )
            """)
            conn.commit()

            # Ensure all columns exist in EMP
            expected_columns = [
                ("RazaoSocial", "TEXT"),
                ("Fantasia", "TEXT"),
                ("CNPJ", "TEXT"),
                ("InscEst", "TEXT"),
                ("Logradouro", "TEXT"),
                ("Nro", "TEXT"),
                ("Bairro", "TEXT"),
                ("Cidade", "TEXT"),
                ("UF", "VARCHAR(2)"),
                ("CEP", "TEXT"),
                ("Fone", "TEXT"),
                ("CodigoIBGE", "TEXT"),
                ("RegimeTrib", "TEXT"),
                ("PIS", "TEXT"),
                ("AliqPIS", "TEXT"),
                ("COFINS", "TEXT"),
                ("AliqCOFINS", "TEXT"),
                ("SitTrib", "TEXT"),
                ("CFOP", "TEXT"),
                ("SitTribFixo", "TEXT"),
                ("Deducao", "TEXT"),
                ("Ativo", "INTEGER DEFAULT 1"),
                ("DtCadastro", "TEXT")
            ]
            for col_name, col_type in expected_columns:
                try:
                    cursor.execute(f'ALTER TABLE "EMP" ADD COLUMN IF NOT EXISTS "{col_name}" {col_type}')
                except Exception:
                    pass
            conn.commit()

            cursor.execute('SELECT COUNT(*) as cnt FROM "EMP"')
            count = cursor.fetchone()['cnt']

            if count == 0:
                razao = "EMPRESA MATRIZ"
                fantasia = "MATRIZ"
                cnpj = ""
                uf = "SP"
                ibge = ""
                logradouro = ""
                nro = ""
                bairro = ""
                cidade = ""
                ie = ""
                cep = ""
                fone = ""
                regime = "1"
                pis = "01"
                aliq_pis = "0,65"
                cofins = "01"
                aliq_cofins = "3,00"
                sit_trib = "102"
                cfop = "5102"
                sit_trib_fixo = "102"
                deducao = "0,00"

                if os.path.exists(NFE_INI_PATH):
                    try:
                        config = configparser.ConfigParser()
                        config.optionxform = str
                        config.read(NFE_INI_PATH, encoding='latin1')
                        if config.has_section('EMITENTE'):
                            e = config['EMITENTE']
                            razao = e.get('RAZAOSOCIAL', razao)
                            fantasia = e.get('FANTASIA', fantasia)
                            cnpj = e.get('CNPJ', cnpj)
                            uf = e.get('UF', uf)
                            ibge = e.get('CODIGOIBGE', ibge)
                            logradouro = e.get('LOGRADOURO', logradouro)
                            nro = e.get('NRO', nro)
                            bairro = e.get('BAIRRO', bairro)
                            cidade = e.get('CIDADE', cidade)
                            ie = e.get('INSCEST', ie)
                            cep = e.get('CEP', cep)
                            fone = e.get('FONE', fone)
                            regime = e.get('REGIMETRIB', regime)
                            pis = e.get('Pis', pis)
                            aliq_pis = e.get('aliqPis', aliq_pis)
                            cofins = e.get('Cofins', cofins)
                            aliq_cofins = e.get('aliqCofins', aliq_cofins)
                            sit_trib = e.get('SitTrib', sit_trib)
                        if config.has_section('FIXO'):
                            f = config['FIXO']
                            cfop = f.get('Cfop', cfop)
                            sit_trib_fixo = f.get('SitTribFixo', sit_trib_fixo)
                            deducao = f.get('Deducao', deducao)
                    except Exception as ex:
                        print("Notice: Error reading nfe.ini fallback to defaults:", ex)

                cursor.execute("""
                    INSERT INTO "EMP" (
                        "id_empresa", "RazaoSocial", "Fantasia", "CNPJ", "InscEst", "Logradouro", "Nro", "Bairro",
                        "Cidade", "UF", "CEP", "Fone", "CodigoIBGE", "RegimeTrib", "PIS", "AliqPIS", "COFINS", "AliqCOFINS",
                        "SitTrib", "CFOP", "SitTribFixo", "Deducao", "Ativo", "DtCadastro"
                    ) VALUES (
                        1, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 1, CURRENT_DATE::text
                    )
                """, (
                    razao, fantasia, cnpj, ie, logradouro, nro, bairro, cidade, uf, cep, fone,
                    ibge, regime, pis, aliq_pis, cofins, aliq_cofins, sit_trib, cfop, sit_trib_fixo, deducao
                ))
                conn.commit()
    finally:
        conn.close()

def migrate_all_tables_multiempresa():
    init_empresa_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public' AND table_name != 'EMP'")
            tables = [r['table_name'] for r in cursor.fetchall()]

            altered_count = 0
            for table in tables:
                cursor.execute("""
                    SELECT column_name FROM information_schema.columns 
                    WHERE table_schema='public' AND table_name=%s AND column_name='id_empresa'
                """, (table,))
                has_col = cursor.fetchone()
                if not has_col:
                    try:
                        cursor.execute(f'ALTER TABLE "{table}" ADD COLUMN id_empresa BIGINT DEFAULT 1')
                        cursor.execute(f'UPDATE "{table}" SET id_empresa = 1 WHERE id_empresa IS NULL')
                        altered_count += 1
                    except Exception as err:
                        print(f"Notice: Could not add id_empresa to {table}: {err}")

            conn.commit()
    finally:
        conn.close()

def get_empresas(include_inactive=False):
    init_empresa_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            sql = 'SELECT * FROM "EMP"'
            if not include_inactive:
                sql += ' WHERE "Ativo" = 1 OR "Ativo" IS NULL'
            sql += ' ORDER BY "id_empresa" ASC'
            cursor.execute(sql)
            return [normalize_empresa_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_empresa_by_id(id_empresa):
    init_empresa_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "EMP" WHERE "id_empresa" = %s', (id_empresa,))
            row = cursor.fetchone()
            return normalize_empresa_row(row) if row else None
    finally:
        conn.close()

def save_empresa(data):
    init_empresa_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            id_emp = data.get('id_empresa')
            razao = data.get('RazaoSocial', '').strip()
            fantasia = data.get('Fantasia', '').strip()
            cnpj = data.get('CNPJ', '').strip()
            ie = data.get('InscEst', '').strip()
            logradouro = data.get('Logradouro', '').strip()
            nro = data.get('Nro', '').strip()
            bairro = data.get('Bairro', '').strip()
            cidade = data.get('Cidade', '').strip()
            uf = data.get('UF', 'SP').strip().upper()
            cep = data.get('CEP', '').strip()
            fone = data.get('Fone', '').strip()
            ibge = data.get('CodigoIBGE', '').strip()
            regime = data.get('RegimeTrib', '3').strip()
            pis = data.get('PIS', '01').strip()
            aliq_pis = data.get('AliqPIS', '0,65').strip()
            cofins = data.get('COFINS', '01').strip()
            aliq_cofins = data.get('AliqCOFINS', '3').strip()
            sit_trib = data.get('SitTrib', '').strip()
            cfop = data.get('CFOP', '').strip()
            sit_trib_fixo = data.get('SitTribFixo', '').strip()
            deducao = data.get('Deducao', '33,33').strip()
            ativo = 1 if str(data.get('Ativo', '1')) in ['1', '-1', 'true', 'True'] else 0

            if id_emp and str(id_emp).isdigit() and int(id_emp) > 0:
                id_emp = int(id_emp)
                cursor.execute("""
                    UPDATE "EMP" SET
                        "RazaoSocial" = %s, "Fantasia" = %s, "CNPJ" = %s, "InscEst" = %s, "Logradouro" = %s, "Nro" = %s,
                        "Bairro" = %s, "Cidade" = %s, "UF" = %s, "CEP" = %s, "Fone" = %s, "CodigoIBGE" = %s, "RegimeTrib" = %s,
                        "PIS" = %s, "AliqPIS" = %s, "COFINS" = %s, "AliqCOFINS" = %s, "SitTrib" = %s, "CFOP" = %s,
                        "SitTribFixo" = %s, "Deducao" = %s, "Ativo" = %s
                    WHERE "id_empresa" = %s
                """, (
                    razao, fantasia, cnpj, ie, logradouro, nro, bairro, cidade, uf, cep, fone,
                    ibge, regime, pis, aliq_pis, cofins, aliq_cofins, sit_trib, cfop, sit_trib_fixo, deducao, ativo, id_emp
                ))
            else:
                cursor.execute('SELECT COALESCE(MAX("id_empresa"), 0) + 1 as next_id FROM "EMP"')
                next_id = cursor.fetchone()['next_id']
                cursor.execute("""
                    INSERT INTO "EMP" (
                        "id_empresa", "RazaoSocial", "Fantasia", "CNPJ", "InscEst", "Logradouro", "Nro", "Bairro", "Cidade", "UF",
                        "CEP", "Fone", "CodigoIBGE", "RegimeTrib", "PIS", "AliqPIS", "COFINS", "AliqCOFINS", "SitTrib",
                        "CFOP", "SitTribFixo", "Deducao", "Ativo", "DtCadastro"
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURRENT_DATE::text
                    )
                """, (
                    next_id, razao, fantasia, cnpj, ie, logradouro, nro, bairro, cidade, uf, cep, fone,
                    ibge, regime, pis, aliq_pis, cofins, aliq_cofins, sit_trib, cfop, sit_trib_fixo, deducao, ativo
                ))
                id_emp = next_id

            conn.commit()
            return get_empresa_by_id(id_emp)
    finally:
        conn.close()
