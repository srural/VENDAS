"""
Módulo de Sincronização e Migração Automática de Estrutura e Dados de Banco de Dados (PostgreSQL)
Permite comparar e sincronizar tabelas, colunas, tipos de dados, valores padrão, índices e dados
entre a base de TESTE/DESENVOLVIMENTO e a base de PRODUÇÃO.
"""

import sys
import os
import argparse
import datetime
import psycopg2
from psycopg2.extras import RealDictCursor, execute_values
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

# Fix Windows console encoding
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Cores para terminal
class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def get_connection(host="localhost", port=5432, user="postgres", password="masterkey", dbname="vendas_db"):
    """Cria conexão com o banco especificado."""
    return psycopg2.connect(
        host=host,
        port=int(port),
        user=user,
        password=password,
        dbname=dbname,
        cursor_factory=RealDictCursor
    )


def database_exists(dbname, host="localhost", port=5432, user="postgres", password="masterkey"):
    """Verifica se o banco de dados existe no servidor."""
    conn = psycopg2.connect(
        host=host,
        port=int(port),
        user=user,
        password=password,
        dbname="postgres"
    )
    conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM pg_database WHERE datname = %s", (dbname,))
            return cur.fetchone() is not None
    finally:
        conn.close()


def create_database(dbname, host="localhost", port=5432, user="postgres", password="masterkey"):
    """Cria um banco de dados novo se não existir."""
    conn = psycopg2.connect(
        host=host,
        port=int(port),
        user=user,
        password=password,
        dbname="postgres"
    )
    conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
    try:
        with conn.cursor() as cur:
            cur.execute(f'CREATE DATABASE "{dbname}" WITH OWNER "{user}" ENCODING "UTF8"')
            print(f"{Colors.GREEN}[+] Banco de dados '{dbname}' criado com sucesso!{Colors.ENDC}")
            return True
    finally:
        conn.close()


def list_available_tables(host="localhost", port=5432, user="postgres", password="masterkey", dbname="vendas_db"):
    """Lista todas as tabelas com contagem de colunas e linhas aproximadas."""
    conn = get_connection(host, port, user, password, dbname)
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT 
                    t.table_name,
                    COUNT(c.column_name) as columns_count
                FROM information_schema.tables t
                LEFT JOIN information_schema.columns c 
                    ON t.table_name = c.table_name AND t.table_schema = c.table_schema
                WHERE t.table_schema = 'public' AND t.table_type = 'BASE TABLE'
                GROUP BY t.table_name
                ORDER BY t.table_name
            """)
            rows = cur.fetchall()
            
            tables = []
            for r in rows:
                tbl_name = r["table_name"]
                # Obter contagem de registros
                try:
                    cur.execute(f'SELECT COUNT(*) as cnt FROM "{tbl_name}"')
                    cnt = cur.fetchone()["cnt"]
                except Exception:
                    cnt = 0
                tables.append({
                    "name": tbl_name,
                    "columns_count": r["columns_count"],
                    "rows_count": cnt
                })
            return tables
    finally:
        conn.close()


def get_full_schema_metadata(conn, table_filter=None):
    """
    Coleta todos os metadados da estrutura do banco:
    - Tabelas
    - Colunas (nome, tipo, tamanho, precisão, escala, nullable, default)
    - Chaves Primárias
    - Índices
    - Sequences
    """
    metadata = {
        "tables": {},
        "primary_keys": {},
        "indexes": {},
        "sequences": []
    }

    with conn.cursor() as cur:
        # 0. Sequences
        cur.execute("""
            SELECT sequence_name 
            FROM information_schema.sequences 
            WHERE sequence_schema = 'public'
            ORDER BY sequence_name
        """)
        for r in cur.fetchall():
            metadata["sequences"].append(r["sequence_name"])

        # 1. Tabelas
        cur.execute("""
            SELECT table_name 
            FROM information_schema.tables 
            WHERE table_schema = 'public' AND table_type = 'BASE TABLE'
            ORDER BY table_name
        """)
        for r in cur.fetchall():
            tbl = r["table_name"]
            if table_filter is None or tbl in table_filter:
                metadata["tables"][tbl] = {}

        # 2. Colunas
        cur.execute("""
            SELECT 
                table_name,
                column_name,
                data_type,
                udt_name,
                character_maximum_length,
                numeric_precision,
                numeric_scale,
                is_nullable,
                column_default
            FROM information_schema.columns
            WHERE table_schema = 'public'
            ORDER BY table_name, ordinal_position
        """)
        for col in cur.fetchall():
            tbl = col["table_name"]
            col_name = col["column_name"]
            if tbl in metadata["tables"]:
                metadata["tables"][tbl][col_name] = col

        # 3. Chaves Primárias
        cur.execute("""
            SELECT 
                tc.table_name, 
                kcu.column_name
            FROM information_schema.table_constraints tc
            JOIN information_schema.key_column_usage kcu
              ON tc.constraint_name = kcu.constraint_name 
             AND tc.table_schema = kcu.table_schema
            WHERE tc.constraint_type = 'PRIMARY KEY' 
              AND tc.table_schema = 'public'
            ORDER BY tc.table_name, kcu.ordinal_position
        """)
        for pk in cur.fetchall():
            tbl = pk["table_name"]
            if tbl in metadata["tables"]:
                if tbl not in metadata["primary_keys"]:
                    metadata["primary_keys"][tbl] = []
                metadata["primary_keys"][tbl].append(pk["column_name"])

        # 4. Índices
        cur.execute("""
            SELECT tablename, indexname, indexdef
            FROM pg_indexes
            WHERE schemaname = 'public'
        """)
        for idx in cur.fetchall():
            tbl = idx["tablename"]
            if tbl in metadata["tables"]:
                if tbl not in metadata["indexes"]:
                    metadata["indexes"][tbl] = {}
                metadata["indexes"][tbl][idx["indexname"]] = idx["indexdef"]

    return metadata


def format_sql_data_type(col_meta):
    """Converte os metadados do information_schema para a declaração de tipo SQL correspondente."""
    dt = col_meta["data_type"].upper()
    udt = col_meta.get("udt_name", "").lower()
    char_len = col_meta.get("character_maximum_length")
    num_prec = col_meta.get("numeric_precision")
    num_scale = col_meta.get("numeric_scale")

    if dt in ("CHARACTER VARYING", "VARCHAR"):
        return f"VARCHAR({char_len})" if char_len else "VARCHAR"
    elif dt in ("CHARACTER", "CHAR"):
        return f"CHAR({char_len})" if char_len else "CHAR(1)"
    elif dt in ("NUMERIC", "DECIMAL"):
        if num_prec and num_scale:
            return f"NUMERIC({num_prec}, {num_scale})"
        elif num_prec:
            return f"NUMERIC({num_prec})"
        return "NUMERIC"
    elif dt == "USER-DEFINED":
        return udt
    elif dt == "ARRAY":
        return f"{udt[1:]}[]" if udt.startswith("_") else f"{udt}[]"
    elif "TIMESTAMP" in dt:
        return "TIMESTAMP"
    elif "TIME" in dt:
        return "TIME"
    return dt


def compare_schemas(source_meta, target_meta, source_name="Origem", target_name="Destino"):
    """
    Compara a estrutura do banco de Origem com o Destino e gera a lista de DDLs necessários.
    """
    diffs = {
        "missing_tables": [],
        "missing_columns": [],
        "type_mismatches": [],
        "missing_indexes": [],
        "sql_statements": []
    }

    # 0. Sequences Faltantes no Destino
    src_seqs = source_meta.get("sequences", [])
    tgt_seqs = target_meta.get("sequences", [])
    for seq_name in src_seqs:
        if seq_name not in tgt_seqs:
            diffs["sql_statements"].append({
                "type": "CREATE_SEQUENCE",
                "table": "-",
                "description": f"Criar sequence '{seq_name}'",
                "sql": f'CREATE SEQUENCE IF NOT EXISTS "{seq_name}";'
            })

    # 1. Tabelas Faltantes no Destino
    for table_name, cols in source_meta["tables"].items():
        if table_name not in target_meta["tables"]:
            diffs["missing_tables"].append(table_name)
            
            # Sequences referenciadas
            for col_name, col_meta in cols.items():
                col_def_val = str(col_meta.get("column_default") or '')
                if "nextval" in col_def_val:
                    import re
                    seq_match = re.search(r"nextval\('\"?([^'\":]+)\"?'", col_def_val)
                    if seq_match:
                        s_name = seq_match.group(1)
                        diffs["sql_statements"].append({
                            "type": "CREATE_SEQUENCE",
                            "table": table_name,
                            "description": f"Criar sequence '{s_name}' para a tabela '{table_name}'",
                            "sql": f'CREATE SEQUENCE IF NOT EXISTS "{s_name}";'
                        })

            # Monta DDL CREATE TABLE
            col_defs = []
            for col_name, col_meta in cols.items():
                col_type = format_sql_data_type(col_meta)
                nullable = "NOT NULL" if col_meta["is_nullable"] == "NO" else ""
                default = f"DEFAULT {col_meta['column_default']}" if col_meta["column_default"] else ""
                col_defs.append(f'    "{col_name}" {col_type} {default} {nullable}'.strip())

            # Primary Key
            if table_name in source_meta["primary_keys"]:
                pks = [f'"{p}"' for p in source_meta["primary_keys"][table_name]]
                col_defs.append(f'    CONSTRAINT "pk_{table_name}" PRIMARY KEY ({", ".join(pks)})')

            create_table_sql = f'CREATE TABLE "{table_name}" (\n' + ",\n".join(col_defs) + "\n);"
            diffs["sql_statements"].append({
                "type": "CREATE_TABLE",
                "table": table_name,
                "description": f"Criar tabela inexistente '{table_name}'",
                "sql": create_table_sql
            })
        else:
            # 2. Tabela existe em ambos: verificar Colunas Faltantes ou Modificadas
            target_cols = target_meta["tables"][table_name]

            for col_name, src_col in cols.items():
                if col_name not in target_cols:
                    col_type = format_sql_data_type(src_col)
                    default_str = f"DEFAULT {src_col['column_default']}" if src_col["column_default"] else ""
                    nullable_str = "NOT NULL" if src_col["is_nullable"] == "NO" and src_col["column_default"] else ""
                    
                    alter_sql = f'ALTER TABLE "{table_name}" ADD COLUMN "{col_name}" {col_type} {default_str} {nullable_str};'.strip()
                    
                    diffs["missing_columns"].append({
                        "table": table_name,
                        "column": col_name,
                        "type": col_type
                    })
                    diffs["sql_statements"].append({
                        "type": "ADD_COLUMN",
                        "table": table_name,
                        "column": col_name,
                        "description": f"Adicionar coluna '{col_name}' ({col_type}) na tabela '{table_name}'",
                        "sql": alter_sql
                    })
                else:
                    tgt_col = target_cols[col_name]
                    src_type = format_sql_data_type(src_col)
                    tgt_type = format_sql_data_type(tgt_col)

                    if src_type != tgt_type:
                        diffs["type_mismatches"].append({
                            "table": table_name,
                            "column": col_name,
                            "src_type": src_type,
                            "tgt_type": tgt_type
                        })
                        alter_type_sql = f'ALTER TABLE "{table_name}" ALTER COLUMN "{col_name}" TYPE {src_type} USING "{col_name}"::{src_type};'
                        diffs["sql_statements"].append({
                            "type": "ALTER_TYPE",
                            "table": table_name,
                            "column": col_name,
                            "description": f"Alterar tipo da coluna '{col_name}' na tabela '{table_name}' de {tgt_type} para {src_type}",
                            "sql": alter_type_sql
                        })

    # 3. Índices Faltantes
    for table_name, idx_dict in source_meta["indexes"].items():
        if table_name in target_meta["tables"]:
            tgt_idx_dict = target_meta["indexes"].get(table_name, {})
            for idx_name, idx_def in idx_dict.items():
                if "_pkey" not in idx_name and idx_name not in tgt_idx_dict:
                    diffs["missing_indexes"].append({
                        "table": table_name,
                        "index": idx_name
                    })
                    diffs["sql_statements"].append({
                        "type": "CREATE_INDEX",
                        "table": table_name,
                        "index": idx_name,
                        "description": f"Criar índice '{idx_name}' na tabela '{table_name}'",
                        "sql": f"{idx_def};"
                    })

    return diffs


def api_compare_schemas(source_cfg, target_cfg, table_names=None):
    """
    Função para uso da API Web: Compara os bancos e retorna o relatório estruturado em dict/JSON.
    """
    # Verifica se o banco de destino existe
    target_exists = database_exists(
        target_cfg["dbname"],
        host=target_cfg.get("host", "localhost"),
        port=target_cfg.get("port", 5432),
        user=target_cfg.get("user", "postgres"),
        password=target_cfg.get("password", "masterkey")
    )

    if not target_exists:
        return {
            "success": True,
            "target_exists": False,
            "source_info": f"{source_cfg['dbname']}@{source_cfg.get('host','localhost')}:{source_cfg.get('port', 5432)}",
            "target_info": f"{target_cfg['dbname']}@{target_cfg.get('host','localhost')}:{target_cfg.get('port', 5432)}",
            "message": f"O banco de dados de destino '{target_cfg['dbname']}' ainda não existe e será criado automaticamente durante a sincronização.",
            "diffs": {
                "missing_tables": ["(Todas as tabelas serão criadas)"],
                "missing_columns": [],
                "type_mismatches": [],
                "missing_indexes": [],
                "sql_statements": []
            }
        }

    try:
        src_conn = get_connection(**source_cfg)
    except Exception as e:
        return {"success": False, "message": f"Falha ao conectar no banco de Origem/Teste: {str(e)}"}

    try:
        tgt_conn = get_connection(**target_cfg)
    except Exception as e:
        src_conn.close()
        return {"success": False, "message": f"Falha ao conectar no banco de Destino/Produção: {str(e)}"}

    try:
        filter_set = set(table_names) if table_names else None
        src_meta = get_full_schema_metadata(src_conn, table_filter=filter_set)
        tgt_meta = get_full_schema_metadata(tgt_conn, table_filter=filter_set)

        diffs = compare_schemas(src_meta, tgt_meta, source_cfg['dbname'], target_cfg['dbname'])

        return {
            "success": True,
            "target_exists": True,
            "source_info": f"{source_cfg['dbname']}@{source_cfg.get('host','localhost')}:{source_cfg.get('port', 5432)}",
            "target_info": f"{target_cfg['dbname']}@{target_cfg.get('host','localhost')}:{target_cfg.get('port', 5432)}",
            "source_tables_count": len(src_meta["tables"]),
            "target_tables_count": len(tgt_meta["tables"]),
            "diffs": diffs,
            "total_changes": len(diffs["sql_statements"])
        }
    finally:
        src_conn.close()
        tgt_conn.close()


def api_apply_sync(source_cfg, target_cfg, table_names=None, sync_data=False, sql_statements=None, create_target_if_needed=True):
    """
    Aplica as alterações DDL (e opcionalmente DML/dados) do banco de Teste para o de Produção.
    """
    # 1. Verificar/Criar banco de destino se não existir
    target_exists = database_exists(
        target_cfg["dbname"],
        host=target_cfg.get("host", "localhost"),
        port=target_cfg.get("port", 5432),
        user=target_cfg.get("user", "postgres"),
        password=target_cfg.get("password", "masterkey")
    )

    if not target_exists:
        if create_target_if_needed:
            create_database(
                target_cfg["dbname"],
                host=target_cfg.get("host", "localhost"),
                port=target_cfg.get("port", 5432),
                user=target_cfg.get("user", "postgres"),
                password=target_cfg.get("password", "masterkey")
            )
        else:
            return {"success": False, "message": f"Banco de destino '{target_cfg['dbname']}' não existe."}

    # 2. Conectar
    try:
        src_conn = get_connection(**source_cfg)
        tgt_conn = get_connection(**target_cfg)
    except Exception as e:
        return {"success": False, "message": f"Erro ao conectar nos bancos: {str(e)}"}

    try:
        filter_set = set(table_names) if table_names else None
        
        # Se os comandos SQL não foram passados explicitamente, recalculamos
        if sql_statements is None:
            src_meta = get_full_schema_metadata(src_conn, table_filter=filter_set)
            tgt_meta = get_full_schema_metadata(tgt_conn, table_filter=filter_set)
            diffs = compare_schemas(src_meta, tgt_meta, source_cfg['dbname'], target_cfg['dbname'])
            statements = diffs["sql_statements"]
        else:
            statements = sql_statements

        executed_log = []
        success_count = 0
        error_count = 0

        # 3. Executar DDLs
        with tgt_conn.cursor() as cur:
            for item in statements:
                sql = item.get("sql")
                desc = item.get("description", sql)
                try:
                    cur.execute(sql)
                    tgt_conn.commit()
                    executed_log.append({
                        "status": "OK",
                        "description": desc,
                        "sql": sql
                    })
                    success_count += 1
                except Exception as err:
                    tgt_conn.rollback()
                    executed_log.append({
                        "status": "ERROR",
                        "description": desc,
                        "sql": sql,
                        "error": str(err)
                    })
                    error_count += 1

        # 4. Sincronizar Dados se solicitado
        data_sync_summary = {}
        if sync_data:
            # Obter tabelas a sincronizar dados
            src_meta_now = get_full_schema_metadata(src_conn, table_filter=filter_set)
            tgt_meta_now = get_full_schema_metadata(tgt_conn, table_filter=filter_set)

            for tbl_name in src_meta_now["tables"]:
                if tbl_name not in tgt_meta_now["tables"]:
                    continue

                pks = src_meta_now["primary_keys"].get(tbl_name, [])
                if not pks:
                    # Se não tem PK explícita, ignora upsert seguro para evitar duplicidade
                    data_sync_summary[tbl_name] = {"status": "SKIPPED_NO_PK", "rows": 0}
                    continue

                try:
                    with src_conn.cursor() as src_cur:
                        src_cur.execute(f'SELECT * FROM "{tbl_name}"')
                        src_rows = src_cur.fetchall()

                    if not src_rows:
                        data_sync_summary[tbl_name] = {"status": "EMPTY", "rows": 0}
                        continue

                    cols = list(src_rows[0].keys())
                    cols_escaped = [f'"{c}"' for c in cols]
                    pk_escaped = [f'"{p}"' for p in pks]

                    # Cláusula UPDATE para ON CONFLICT
                    update_cols = [c for c in cols if c not in pks]
                    if update_cols:
                        update_clause = ", ".join([f'"{c}" = EXCLUDED."{c}"' for c in update_cols])
                        conflict_action = f'DO UPDATE SET {update_clause}'
                    else:
                        conflict_action = 'DO NOTHING'

                    upsert_sql = f"""
                        INSERT INTO "{tbl_name}" ({', '.join(cols_escaped)})
                        VALUES %s
                        ON CONFLICT ({', '.join(pk_escaped)})
                        {conflict_action}
                    """

                    with tgt_conn.cursor() as tgt_cur:
                        values_tuples = [tuple(row[col] for col in cols) for row in src_rows]
                        execute_values(tgt_cur, upsert_sql, values_tuples)
                        tgt_conn.commit()

                    data_sync_summary[tbl_name] = {"status": "SYNCED", "rows": len(src_rows)}
                except Exception as ex:
                    tgt_conn.rollback()
                    data_sync_summary[tbl_name] = {"status": "ERROR", "error": str(ex)}

        return {
            "success": error_count == 0,
            "total_statements": len(statements),
            "success_count": success_count,
            "error_count": error_count,
            "executed_log": executed_log,
            "data_sync_summary": data_sync_summary,
            "message": f"Sincronização concluída! {success_count} alterações de estrutura aplicadas com sucesso no banco '{target_cfg['dbname']}'."
        }
    finally:
        src_conn.close()
        tgt_conn.close()


def sync_databases(source_cfg, target_cfg, dry_run=True, export_file=None, create_target=False):
    """
    Função CLI que conecta, compara e sincroniza a estrutura dos dois bancos.
    """
    print(f"\n{Colors.BOLD}{Colors.HEADER}================================================================={Colors.ENDC}")
    print(f"{Colors.BOLD}{Colors.HEADER}    ROTINA DE SINCRONIZAÇÃO DE BANCO DE DADOS (POSTGRESQL)       {Colors.ENDC}")
    print(f"{Colors.BOLD}{Colors.HEADER}================================================================={Colors.ENDC}\n")

    print(f"[*] {Colors.CYAN}Banco de Origem (Referencia/Teste):{Colors.ENDC} {source_cfg['dbname']}@{source_cfg['host']}:{source_cfg['port']}")
    print(f"[*] {Colors.BLUE}Banco de Destino (Producao/Alvo):{Colors.ENDC} {target_cfg['dbname']}@{target_cfg['host']}:{target_cfg['port']}")
    print(f"[*] {Colors.BOLD}Modo de Execucao:{Colors.ENDC} {'[DRY-RUN] (Apenas Analise / Sem alteracoes)' if dry_run else '[APLICAR MIGRACAO] (Modificar Destino)'}\n")

    if not database_exists(target_cfg["dbname"], host=target_cfg["host"], port=target_cfg["port"], user=target_cfg["user"], password=target_cfg["password"]):
        if create_target or not dry_run:
            print(f"{Colors.WARNING}[!] Banco de destino '{target_cfg['dbname']}' nao existe. Criando agora...{Colors.ENDC}")
            create_database(target_cfg["dbname"], host=target_cfg["host"], port=target_cfg["port"], user=target_cfg["user"], password=target_cfg["password"])
        else:
            print(f"{Colors.FAIL}[ERRO] Banco de destino '{target_cfg['dbname']}' nao existe no servidor.{Colors.ENDC}")
            print(f"Utilize o parametro --create-target para cria-lo automaticamente.")
            return False

    try:
        src_conn = get_connection(**source_cfg)
        tgt_conn = get_connection(**target_cfg)
    except Exception as e:
        print(f"{Colors.FAIL}[ERRO AO CONECTAR]: {e}{Colors.ENDC}")
        return False

    try:
        print(f"[*] Lendo metadados da base de origem [{source_cfg['dbname']}]...")
        src_meta = get_full_schema_metadata(src_conn)
        print(f"    -> {len(src_meta['tables'])} tabelas encontradas na origem.")

        print(f"[*] Lendo metadados da base de destino [{target_cfg['dbname']}]...")
        tgt_meta = get_full_schema_metadata(tgt_conn)
        print(f"    -> {len(tgt_meta['tables'])} tabelas encontradas no destino.")

        print(f"\n[*] Comparando estruturas e identificando diferencas...")
        diffs = compare_schemas(src_meta, tgt_meta, source_cfg['dbname'], target_cfg['dbname'])

        stmt_count = len(diffs["sql_statements"])

        print(f"\n{Colors.BOLD}--- RELATORIO DE DIFERENCAS ---{Colors.ENDC}")
        print(f" - Tabelas faltantes no destino: {Colors.WARNING}{len(diffs['missing_tables'])}{Colors.ENDC}")
        print(f" - Colunas faltantes no destino: {Colors.WARNING}{len(diffs['missing_columns'])}{Colors.ENDC}")
        print(f" - Colunas com tipos diferentes: {Colors.WARNING}{len(diffs['type_mismatches'])}{Colors.ENDC}")
        print(f" - Indices faltantes no destino: {Colors.WARNING}{len(diffs['missing_indexes'])}{Colors.ENDC}")
        print(f" - Total de instrucoes SQL a executar: {Colors.BOLD}{stmt_count}{Colors.ENDC}\n")

        if stmt_count == 0:
            print(f"{Colors.GREEN}{Colors.BOLD}[OK] PERFEITO! As estruturas dos dois bancos ja estao 100% identicas e sincronizadas!{Colors.ENDC}\n")
            return True

        print(f"{Colors.CYAN}{Colors.BOLD}Detalhamento das alteracoes necessarias:{Colors.ENDC}")
        for idx, item in enumerate(diffs["sql_statements"], 1):
            print(f" {idx}. [{item['type']}] {item['description']}")
            print(f"    {Colors.WARNING}{item['sql']}{Colors.ENDC}\n")

        if export_file:
            with open(export_file, "w", encoding="utf-8") as f:
                f.write(f"-- SCRIPT DE MIGRACAO GERADO EM: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"-- Origem: {source_cfg['dbname']} -> Destino: {target_cfg['dbname']}\n\n")
                for item in diffs["sql_statements"]:
                    f.write(f"-- {item['description']}\n")
                    f.write(f"{item['sql']}\n\n")
            print(f"[+] {Colors.GREEN}Script SQL salvo com sucesso em: '{export_file}'{Colors.ENDC}\n")

        if not dry_run:
            print(f"[*] {Colors.BOLD}Aplicando {stmt_count} alteracoes no banco '{target_cfg['dbname']}'...{Colors.ENDC}")
            success_count = 0
            with tgt_conn.cursor() as cur:
                for idx, item in enumerate(diffs["sql_statements"], 1):
                    try:
                        cur.execute(item["sql"])
                        tgt_conn.commit()
                        print(f"  [{idx}/{stmt_count}] {Colors.GREEN}[OK] {item['description']}{Colors.ENDC}")
                        success_count += 1
                    except Exception as err:
                        tgt_conn.rollback()
                        print(f"  [{idx}/{stmt_count}] {Colors.FAIL}[ERRO] Erro ao executar: {item['sql']}{Colors.ENDC}")
                        print(f"     Detalhe do erro: {err}")

            print(f"\n{Colors.BOLD}{Colors.GREEN}[SUCESSO] Sincronizacao concluida! {success_count}/{stmt_count} alteracoes aplicadas com sucesso no banco '{target_cfg['dbname']}'.{Colors.ENDC}\n")
        else:
            print(f"{Colors.WARNING}[INFO] Modo de simulacao (Dry-Run). Nenhuma alteracao foi realizada no banco de destino.{Colors.ENDC}")
            print(f"       Para aplicar as alteracoes, execute novamente com o parametro: {Colors.BOLD}--apply{Colors.ENDC}\n")

        return True
    finally:
        src_conn.close()
        tgt_conn.close()


def main():
    parser = argparse.ArgumentParser(
        description="Sincronizador de Estrutura de Banco de Dados PostgreSQL (Teste <-> Produção)"
    )
    parser.add_argument("--source-db", "--src-db", default="vendas_db", help="Nome do banco de dados de Origem/Teste (padrão: vendas_db)")
    parser.add_argument("--target-db", "--tgt-db", default="vendas_prod_db", help="Nome do banco de dados de Destino/Produção (padrão: vendas_prod_db)")
    parser.add_argument("--host", default=os.environ.get("PG_HOST", "localhost"), help="Host do PostgreSQL (padrão: localhost)")
    parser.add_argument("--port", type=int, default=int(os.environ.get("PG_PORT", 5432)), help="Porta do PostgreSQL (padrão: 5432)")
    parser.add_argument("--user", default=os.environ.get("PG_USER", "postgres"), help="Usuário do PostgreSQL (padrão: postgres)")
    parser.add_argument("--password", "--pass", default=os.environ.get("PG_PASS", "masterkey"), help="Senha do PostgreSQL (padrão: masterkey)")
    
    parser.add_argument("--apply", action="store_true", help="Aplica as alterações no banco de destino (sem este parâmetro roda em modo simulação Dry-Run)")
    parser.add_argument("--export-sql", help="Caminho do arquivo para exportar o script SQL gerado")
    parser.add_argument("--create-target", action="store_true", help="Cria o banco de destino caso ele não exista no servidor")

    args = parser.parse_args()

    source_cfg = {
        "host": args.host,
        "port": args.port,
        "user": args.user,
        "password": args.password,
        "dbname": args.source_db
    }

    target_cfg = {
        "host": args.host,
        "port": args.port,
        "user": args.user,
        "password": args.password,
        "dbname": args.target_db
    }

    sync_databases(
        source_cfg=source_cfg,
        target_cfg=target_cfg,
        dry_run=not args.apply,
        export_file=args.export_sql,
        create_target=args.create_target
    )


if __name__ == "__main__":
    main()
