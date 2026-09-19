import os
import psycopg2
from psycopg2.extras import RealDictCursor

PG_HOST = os.environ.get("PG_HOST", "localhost")
PG_PORT = int(os.environ.get("PG_PORT", 5432))
PG_USER = os.environ.get("PG_USER", "postgres")
PG_PASS = os.environ.get("PG_PASS", "masterkey")
PG_DB = os.environ.get("PG_DB", "vendas_db")

def get_connection():
    return psycopg2.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASS,
        dbname=PG_DB,
        cursor_factory=RealDictCursor
    )

# Helper function to convert numeric/decimal values to float or int for JSON serialization
def _convert_row(r):
    if not r:
        return None
    d = dict(r)
    for k, v in list(d.items()):
        if hasattr(v, '__float__') and not isinstance(v, (int, float, bool)):
            d[k] = float(v)
    if "nomecliente" in d and "NomeCliente" not in d:
        d["NomeCliente"] = d["nomecliente"]
    return d

# --- ENTITY COLUMNS & FUNCTIONS ---
def get_column_names(table_name="ENT"):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT column_name 
                FROM information_schema.columns 
                WHERE table_schema = 'public' AND LOWER(table_name) = LOWER(%s)
                ORDER BY ordinal_position
            """, (table_name,))
            return [row["column_name"] for row in cursor.fetchall()]
    finally:
        conn.close()

def get_table_column_types(table_name="ENT"):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT column_name, data_type 
                FROM information_schema.columns 
                WHERE table_schema = 'public' AND LOWER(table_name) = LOWER(%s)
                ORDER BY ordinal_position
            """, (table_name,))
            return {row["column_name"]: row["data_type"] for row in cursor.fetchall()}
    finally:
        conn.close()

def ensure_prd_foto_column():
    try:
        conn = get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    DO $$
                    BEGIN
                        IF EXISTS (
                            SELECT 1 FROM information_schema.columns 
                            WHERE table_schema = 'public' AND LOWER(table_name) = 'prd' AND LOWER(column_name) = 'foto'
                        ) THEN
                            ALTER TABLE "PRD" ALTER COLUMN "Foto" TYPE TEXT USING "Foto"::TEXT;
                        ELSE
                            ALTER TABLE "PRD" ADD COLUMN "Foto" TEXT;
                        END IF;
                    END $$;
                """)
                conn.commit()
        finally:
            conn.close()
    except Exception as e:
        print(f"Aviso ao verificar coluna Foto em PRD: {e}")

ensure_prd_foto_column()

ALL_COLUMNS = get_column_names("ENT")
GRU_COLUMNS = get_column_names("GRU")
PRD_COLUMNS = get_column_names("PRD")

ENT_COL_TYPES = get_table_column_types("ENT")
PRD_COL_TYPES = get_table_column_types("PRD")
GRU_COL_TYPES = get_table_column_types("GRU")

def sanitize_value_for_type(val, data_type):
    if val is None:
        return None
    
    if data_type in ('integer', 'bigint', 'smallint'):
        if isinstance(val, (int, bool)):
            return int(val)
        s = str(val).strip()
        if not s or s.lower() == 'null' or s.lower() == 'nan':
            return None
        try:
            return int(float(s))
        except (ValueError, TypeError):
            return None

    elif data_type in ('numeric', 'decimal', 'double precision', 'real'):
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace('R$', '').replace(' ', '')
        if not s or s.lower() == 'null' or s.lower() == 'nan':
            return None
        if ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        elif ',' in s:
            s = s.replace(',', '.')
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    elif data_type in ('date', 'timestamp', 'timestamp without time zone', 'time'):
        s = str(val).strip()
        if not s or s.lower() == 'null':
            return None
        return s

    else:
        # text / varchar
        s = str(val) if val is not None else None
        return s

def sanitize_data_for_table(data, col_types):
    sanitized = {}
    for k, v in data.items():
        matched_col = None
        if k in col_types:
            matched_col = k
        else:
            k_norm = k.replace('ç', 'c').replace('Ç', 'C').replace('ã', 'a').replace('õ', 'o').replace('é', 'e').replace('ê', 'e')
            for col_name in col_types:
                c_norm = col_name.replace('ç', 'c').replace('Ç', 'C').replace('ã', 'a').replace('õ', 'o').replace('é', 'e').replace('ê', 'e').replace('', 'c')
                if k_norm.lower() == c_norm.lower():
                    matched_col = col_name
                    break
        if matched_col:
            sanitized[matched_col] = sanitize_value_for_type(v, col_types[matched_col])
    return sanitized

DESC_GRU_COL = "Descricao_Grupo" if "Descricao_Grupo" in GRU_COLUMNS else (
    next((c for c in GRU_COLUMNS if "Descri" in c or "Grupo" in c), "Descricao_Grupo")
)
DESC_PRD_COL = "Descricao_Produto" if "Descricao_Produto" in PRD_COLUMNS else (
    next((c for c in PRD_COLUMNS if "Descri" in c or "Produto" in c), "Descricao_Produto")
)
MIN_PRD_COL = "Minimo" if "Minimo" in PRD_COLUMNS else (
    next((c for c in PRD_COLUMNS if "nimo" in c or "Min" in c), "Minimo")
)

def normalize_gru_row(row):
    if not row:
        return None
    d = _convert_row(row)
    if "totalprodutos" in d and "TotalProdutos" not in d:
        d["TotalProdutos"] = d.pop("totalprodutos")
    if DESC_GRU_COL in d and DESC_GRU_COL != "Descricao_Grupo":
        d["Descricao_Grupo"] = d.pop(DESC_GRU_COL)
    if "Desconto" not in d:
        d["Desconto"] = 0.0
    else:
        d["Desconto"] = float(d["Desconto"] or 0.0)
    return d

def normalize_prd_row(row):
    if not row:
        return None
    d = _convert_row(row)
    if "nome_grupo" in d and "Nome_Grupo" not in d:
        d["Nome_Grupo"] = d.pop("nome_grupo")
    elif DESC_GRU_COL in d and DESC_GRU_COL != "Nome_Grupo":
        d["Nome_Grupo"] = d[DESC_GRU_COL]

    if "descricao_produto" in d and "Descricao_Produto" not in d:
        d["Descricao_Produto"] = d.pop("descricao_produto")
    elif DESC_PRD_COL in d and DESC_PRD_COL != "Descricao_Produto":
        d["Descricao_Produto"] = d.pop(DESC_PRD_COL)

    if MIN_PRD_COL in d and MIN_PRD_COL != "Minimo":
        d["Minimo"] = d.pop(MIN_PRD_COL)

    if "CodPrd" not in d and "Codigo" in d:
        d["CodPrd"] = d["Codigo"]

    # Calculate automatic group discount if available
    desconto_grupo = float(d.get("DescontoGrupo", d.get("descontogrupo", d.get("Desconto_Grupo", 0.0))) or 0.0)
    d["DescontoGrupo"] = desconto_grupo
    venda = float(d.get("Venda", 0.0) or 0.0)
    if desconto_grupo > 0 and venda > 0:
        d["PrecoComDesconto"] = round(venda * (1.0 - (desconto_grupo / 100.0)), 2)
        d["ValorDescontoGrupo"] = round(venda * (desconto_grupo / 100.0), 2)
    else:
        d["PrecoComDesconto"] = venda
        d["ValorDescontoGrupo"] = 0.0

    return d

def get_entities(page=1, limit=20, tipo=None, ativo=None, q=None, sort_by="Nome", sort_order="ASC"):
    offset = (page - 1) * limit
    params = []
    where_clauses = []

    if tipo is not None and tipo != "" and tipo != "all":
        try:
            t = int(tipo)
            if t == 1: # Cliente
                where_clauses.append('("Tipo" = 1 OR "Tipo" = 7)')
            elif t == 2: # Fornecedor
                where_clauses.append('("Tipo" = 2 OR "Tipo" = 5)')
            elif t == 3: # Vendedor
                where_clauses.append('("Tipo" = 3 OR "Tipo" = 7)')
            elif t in (4, 5): # Transportadora (Tipo 5)
                where_clauses.append('"Tipo" = 5')
            else:
                where_clauses.append('"Tipo" = %s')
                params.append(t)
        except ValueError:
            pass

    if ativo is not None and ativo != "" and ativo != "all":
        if str(ativo) in ("1", "-1", "true", "True"):
            where_clauses.append('"Ativo" != 0')
        elif str(ativo) in ("0", "false", "False"):
            where_clauses.append('("Ativo" IS NULL OR "Ativo" = 0)')

    order_by_sql = ""
    order_params = []
    if q and q.strip():
        search_term = q.strip()
        search = f"%{search_term}%"
        prefix = f"{search_term}%"
        where_clauses.append("""(
            "Nome" ILIKE %s OR 
            "Fantasia" ILIKE %s OR 
            "CPF" ILIKE %s OR 
            "CGC" ILIKE %s OR 
            "Cidade" ILIKE %s OR 
            "Fone" ILIKE %s OR 
            "Celular" ILIKE %s OR
            CAST("CodEntidade" AS TEXT) ILIKE %s
        )""")
        params.extend([search] * 8)

        order_params = [search_term, search_term, prefix, prefix]
        order_by_sql = """CASE 
            WHEN CAST("CodEntidade" AS TEXT) = %s THEN 1 
            WHEN LOWER("Nome") = LOWER(%s) THEN 2 
            WHEN "Nome" ILIKE %s THEN 3 
            WHEN "Fantasia" ILIKE %s THEN 4 
            ELSE 5 
        END, """

    where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    valid_sorts = {col.lower(): col for col in ALL_COLUMNS}
    sort_field = valid_sorts.get(sort_by.lower(), "Nome")
    sort_dir = "DESC" if sort_order.upper() == "DESC" else "ASC"

    count_sql = f'SELECT COUNT(*) as total FROM "ENT"{where_sql}'
    query_sql = f'SELECT * FROM "ENT"{where_sql} ORDER BY {order_by_sql}"{sort_field}" {sort_dir} LIMIT %s OFFSET %s'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(count_sql, params)
            total = cursor.fetchone()["total"]

            cursor.execute(query_sql, params + order_params + [limit, offset])
            rows = [_convert_row(row) for row in cursor.fetchall()]
    finally:
        conn.close()

    return {
        "items": rows,
        "total": total,
        "page": page,
        "limit": limit,
        "pages": (total + limit - 1) // limit if limit > 0 else 1
    }

def get_entity_by_id(cod_entidade):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "ENT" WHERE "CodEntidade" = %s', (cod_entidade,))
            row = cursor.fetchone()
            return _convert_row(row)
    finally:
        conn.close()

def get_next_id():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT MAX("CodEntidade") as max_id FROM "ENT"')
            row = cursor.fetchone()
            max_id = row['max_id'] if row and row['max_id'] else 0
            return max_id + 1
    finally:
        conn.close()

def create_entity(data):
    clean_data = sanitize_data_for_table(data, ENT_COL_TYPES)
    if 'CodEntidade' not in clean_data or not clean_data['CodEntidade']:
        clean_data['CodEntidade'] = get_next_id()

    fields = []
    placeholders = []
    values = []

    for col in ALL_COLUMNS:
        if col in clean_data:
            fields.append(f'"{col}"')
            placeholders.append('%s')
            values.append(clean_data[col])

    sql = f'INSERT INTO "ENT" ({", ".join(fields)}) VALUES ({", ".join(placeholders)})'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, values)
            conn.commit()
            return clean_data['CodEntidade']
    finally:
        conn.close()

def update_entity(cod_entidade, data):
    clean_data = sanitize_data_for_table(data, ENT_COL_TYPES)
    assignments = []
    values = []

    for col in ALL_COLUMNS:
        if col != 'CodEntidade' and col in clean_data:
            assignments.append(f'"{col}" = %s')
            values.append(clean_data[col])

    if not assignments:
        return False

    values.append(cod_entidade)
    sql = f'UPDATE "ENT" SET {", ".join(assignments)} WHERE "CodEntidade" = %s'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, values)
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()

def delete_entity(cod_entidade):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "ENT" WHERE "CodEntidade" = %s', (cod_entidade,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()


# --- GROUPS (GRU) FUNCTIONS ---

def get_groups(q=None, page=1, limit=100):
    offset = (page - 1) * limit
    params = []
    where_sql = ""

    if q and q.strip():
        search = f"%{q.strip()}%"
        where_sql = f' WHERE (GRU."{DESC_GRU_COL}" ILIKE %s OR CAST(GRU."CodGru" AS TEXT) ILIKE %s)'
        params.extend([search, search])

    sql = f'''
        SELECT GRU.*, COUNT(PRD."CodPrd") as TotalProdutos
        FROM "GRU" GRU
        LEFT JOIN "PRD" PRD ON PRD."Grupo" = GRU."CodGru"
        {where_sql}
        GROUP BY GRU."CodGru", GRU."Descricao_Grupo", GRU."Familia", GRU."BaixarEstoque", GRU."id_empresa", GRU."Desconto"
        ORDER BY GRU."{DESC_GRU_COL}" ASC
        LIMIT %s OFFSET %s
    '''
    
    count_sql = f'SELECT COUNT(*) as total FROM "GRU"{where_sql}'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(count_sql, params)
            total = cursor.fetchone()["total"]

            cursor.execute(sql, params + [limit, offset])
            rows = [normalize_gru_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

    return {
        "items": rows,
        "total": total,
        "page": page,
        "limit": limit,
        "pages": (total + limit - 1) // limit if limit > 0 else 1
    }

def get_group_by_id(cod_gru):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "GRU" WHERE "CodGru" = %s', (cod_gru,))
            row = cursor.fetchone()
            return normalize_gru_row(row)
    finally:
        conn.close()

def get_next_group_id():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT MAX("CodGru") as max_id FROM "GRU"')
            row = cursor.fetchone()
            max_id = row['max_id'] if row and row['max_id'] else 0
            return max_id + 1
    finally:
        conn.close()

def create_group(data):
    if 'Descricao_Grupo' in data and DESC_GRU_COL != 'Descricao_Grupo':
        data[DESC_GRU_COL] = data.pop('Descricao_Grupo')

    clean_data = sanitize_data_for_table(data, GRU_COL_TYPES)
    if 'CodGru' not in clean_data or not clean_data['CodGru']:
        clean_data['CodGru'] = get_next_group_id()

    fields = []
    placeholders = []
    values = []

    for col in GRU_COLUMNS:
        if col in clean_data:
            fields.append(f'"{col}"')
            placeholders.append('%s')
            values.append(clean_data[col])

    sql = f'INSERT INTO "GRU" ({", ".join(fields)}) VALUES ({", ".join(placeholders)})'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, values)
            conn.commit()
            return clean_data['CodGru']
    finally:
        conn.close()

def update_group(cod_gru, data):
    if 'Descricao_Grupo' in data and DESC_GRU_COL != 'Descricao_Grupo':
        data[DESC_GRU_COL] = data.pop('Descricao_Grupo')

    clean_data = sanitize_data_for_table(data, GRU_COL_TYPES)
    assignments = []
    values = []

    for col in GRU_COLUMNS:
        if col != 'CodGru' and col in clean_data:
            assignments.append(f'"{col}" = %s')
            values.append(clean_data[col])

    if not assignments:
        return False

    values.append(cod_gru)
    sql = f'UPDATE "GRU" SET {", ".join(assignments)} WHERE "CodGru" = %s'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, values)
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()

def delete_group(cod_gru):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) as cnt FROM "PRD" WHERE "Grupo" = %s', (cod_gru,))
            count = cursor.fetchone()['cnt']
            if count > 0:
                raise ValueError(f"Não é possível excluir o grupo {cod_gru}: existem {count} produtos vinculados a ele.")

            cursor.execute('DELETE FROM "GRU" WHERE "CodGru" = %s', (cod_gru,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()


# --- PRODUCTS (PRD) FUNCTIONS ---

def get_products(page=1, limit=20, grupo=None, ativo=None, estoque_baixo=None, q=None, sort_by="CodPrd", sort_order="ASC"):
    offset = (page - 1) * limit
    params = []
    where_clauses = []

    if grupo is not None and grupo != "" and grupo != "all":
        try:
            g = int(grupo)
            where_clauses.append('PRD."Grupo" = %s')
            params.append(g)
        except ValueError:
            pass

    if ativo is not None and ativo != "" and ativo != "all":
        if str(ativo) in ("1", "-1", "true", "True"):
            where_clauses.append('PRD."Ativo" != 0')
        elif str(ativo) in ("0", "false", "False"):
            where_clauses.append('(PRD."Ativo" IS NULL OR PRD."Ativo" = 0)')

    if estoque_baixo is not None and str(estoque_baixo).lower() in ("1", "true"):
        where_clauses.append(f'PRD."Estoque" <= PRD."{MIN_PRD_COL}"')

    if q and q.strip():
        search = f"%{q.strip()}%"
        where_clauses.append(f"""(
            PRD."{DESC_PRD_COL}" ILIKE %s OR 
            PRD."CodBar" ILIKE %s OR 
            CAST(PRD."CodPrd" AS TEXT) ILIKE %s OR
            PRD."Marca" ILIKE %s
        )""")
        params.extend([search] * 4)

    where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    valid_sorts = {
        "codprd": 'PRD."CodPrd"',
        "descricao_produto": f'PRD."{DESC_PRD_COL}"',
        "venda": 'PRD."Venda"',
        "estoque": 'PRD."Estoque"',
        "grupo": 'PRD."Grupo"'
    }
    sort_field = valid_sorts.get(sort_by.lower(), 'PRD."CodPrd"')
    sort_dir = "DESC" if sort_order.upper() == "DESC" else "ASC"

    count_sql = f'SELECT COUNT(*) as total FROM "PRD" PRD{where_sql}'
    query_sql = f'''
        SELECT PRD.*, GRU."{DESC_GRU_COL}" as Nome_Grupo, COALESCE(GRU."Desconto", 0.0) as DescontoGrupo
        FROM "PRD" PRD
        LEFT JOIN "GRU" GRU ON PRD."Grupo" = GRU."CodGru"
        {where_sql}
        ORDER BY {sort_field} {sort_dir}
        LIMIT %s OFFSET %s
    '''

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(count_sql, params)
            total = cursor.fetchone()["total"]

            cursor.execute(query_sql, params + [limit, offset])
            rows = []
            for r in cursor.fetchall():
                item = normalize_prd_row(r)
                if DESC_GRU_COL in item and DESC_GRU_COL != "Nome_Grupo":
                    item["Nome_Grupo"] = item.pop(DESC_GRU_COL)
                rows.append(item)
    finally:
        conn.close()

    return {
        "items": rows,
        "total": total,
        "page": page,
        "limit": limit,
        "pages": (total + limit - 1) // limit if limit > 0 else 1
    }

def get_product_by_id(cod_prd):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f'''
                SELECT PRD.*, GRU."{DESC_GRU_COL}" as Nome_Grupo, COALESCE(GRU."Desconto", 0.0) as DescontoGrupo
                FROM "PRD" PRD
                LEFT JOIN "GRU" GRU ON PRD."Grupo" = GRU."CodGru"
                WHERE PRD."CodPrd" = %s
            ''', (cod_prd,))
            row = cursor.fetchone()
            item = normalize_prd_row(row)
            if item and DESC_GRU_COL in item and DESC_GRU_COL != "Nome_Grupo":
                item["Nome_Grupo"] = item.pop(DESC_GRU_COL)
            return item
    finally:
        conn.close()

def get_next_product_id():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT MAX("CodPrd") as max_id FROM "PRD"')
            row = cursor.fetchone()
            max_id = row['max_id'] if row and row['max_id'] else 0
            return max_id + 1
    finally:
        conn.close()

def create_product(data):
    if 'Descricao_Produto' in data and DESC_PRD_COL != 'Descricao_Produto':
        data[DESC_PRD_COL] = data.pop('Descricao_Produto')
    if 'Minimo' in data and MIN_PRD_COL != 'Minimo':
        data[MIN_PRD_COL] = data.pop('Minimo')

    clean_data = sanitize_data_for_table(data, PRD_COL_TYPES)
    if 'CodPrd' not in clean_data or not clean_data['CodPrd']:
        clean_data['CodPrd'] = get_next_product_id()
    if 'Codigo' not in clean_data or not clean_data['Codigo']:
        clean_data['Codigo'] = clean_data['CodPrd']

    fields = []
    placeholders = []
    values = []

    for col in PRD_COLUMNS:
        if col in clean_data:
            fields.append(f'"{col}"')
            placeholders.append('%s')
            values.append(clean_data[col])

    sql = f'INSERT INTO "PRD" ({", ".join(fields)}) VALUES ({", ".join(placeholders)})'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, values)
            conn.commit()
            return clean_data['CodPrd']
    finally:
        conn.close()

def update_product(cod_prd, data):
    if 'Descricao_Produto' in data and DESC_PRD_COL != 'Descricao_Produto':
        data[DESC_PRD_COL] = data.pop('Descricao_Produto')
    if 'Minimo' in data and MIN_PRD_COL != 'Minimo':
        data[MIN_PRD_COL] = data.pop('Minimo')

    clean_data = sanitize_data_for_table(data, PRD_COL_TYPES)
    assignments = []
    values = []

    for col in PRD_COLUMNS:
        if col != 'CodPrd' and col != 'Codigo' and col in clean_data:
            assignments.append(f'"{col}" = %s')
            values.append(clean_data[col])

    if not assignments:
        return False

    values.append(cod_prd)
    sql = f'UPDATE "PRD" SET {", ".join(assignments)} WHERE "CodPrd" = %s'

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, values)
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()

def delete_product(cod_prd):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "PRD" WHERE "CodPrd" = %s', (cod_prd,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()

def get_stats():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) as total FROM "ENT"')
            ent_total = cursor.fetchone()['total']

            cursor.execute('SELECT COUNT(*) as ativos FROM "ENT" WHERE "Ativo" != 0')
            ent_ativos = cursor.fetchone()['ativos']

            cursor.execute('SELECT COUNT(*) as clientes FROM "ENT" WHERE "Tipo" = 1 OR "Tipo" = 5 OR "Tipo" = 7')
            ent_clientes = cursor.fetchone()['clientes']

            cursor.execute('SELECT COUNT(*) as fornecedores FROM "ENT" WHERE "Tipo" = 2 OR "Tipo" = 5')
            ent_fornecedores = cursor.fetchone()['fornecedores']

            cursor.execute('SELECT COUNT(*) as total FROM "GRU"')
            gru_total = cursor.fetchone()['total']

            cursor.execute('SELECT COUNT(*) as total FROM "PRD"')
            prd_total = cursor.fetchone()['total']

            cursor.execute('SELECT COUNT(*) as ativos FROM "PRD" WHERE "Ativo" != 0')
            prd_ativos = cursor.fetchone()['ativos']

            cursor.execute(f'SELECT COUNT(*) as baixo FROM "PRD" WHERE "Estoque" <= "{MIN_PRD_COL}" AND "Ativo" != 0')
            prd_baixo = cursor.fetchone()['baixo']

            cursor.execute('SELECT SUM("Estoque" * "Venda") as valor_total FROM "PRD" WHERE "Estoque" > 0')
            row_val = cursor.fetchone()
            valor_total_estoque = float(row_val['valor_total']) if row_val and row_val['valor_total'] else 0.0

            cursor.execute(f'''
                SELECT GRU."{DESC_GRU_COL}" as "Nome_Grupo", COUNT(PRD."CodPrd") as "TotalProdutos"
                FROM "GRU" GRU
                JOIN "PRD" PRD ON PRD."Grupo" = GRU."CodGru"
                GROUP BY GRU."CodGru", GRU."{DESC_GRU_COL}"
                ORDER BY "TotalProdutos" DESC
                LIMIT 5
            ''')
            top_grupos = [_convert_row(r) for r in cursor.fetchall()]

            cursor.execute(f'''
                SELECT "CodPrd", "{DESC_PRD_COL}" as Descricao_Produto, "Venda", "Estoque"
                FROM "PRD"
                ORDER BY "CodPrd" DESC
                LIMIT 5
            ''')
            ultimos_produtos = [normalize_prd_row(r) for r in cursor.fetchall()]

            cursor.execute('''
                SELECT "CodEntidade", "Nome", "Cidade", "Tipo"
                FROM "ENT"
                ORDER BY "CodEntidade" DESC
                LIMIT 5
            ''')
            ultimas_entidades = [_convert_row(r) for r in cursor.fetchall()]

            return {
                "entidades": {
                    "total": ent_total,
                    "ativos": ent_ativos,
                    "clientes": ent_clientes,
                    "fornecedores": ent_fornecedores
                },
                "produtos": {
                    "total": prd_total,
                    "ativos": prd_ativos,
                    "estoque_baixo": prd_baixo,
                    "valor_total_estoque": valor_total_estoque
                },
                "grupos": {
                    "total": gru_total
                },
                "dashboard": {
                    "top_grupos": top_grupos,
                    "ultimos_produtos": ultimos_produtos,
                    "ultimas_entidades": ultimas_entidades
                }
            }
    finally:
        conn.close()

# --- PDV & NFC-E FUNCTIONS ---

def get_company_info(id_empresa=1):
    try:
        from app.empresa_manager import get_empresa_by_id
        emp = get_empresa_by_id(id_empresa or 1)
        if emp:
            return emp
    except Exception:
        pass
    
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "EMP" ORDER BY "id_empresa" ASC LIMIT 1')
            row = cursor.fetchone()
            if row:
                return _convert_row(row)
            return {
                "RazaoSocial": "SIDIVAL CARLOS CIOCA",
                "Fantasia": "SIDIVAL CARLOS CIOCA",
                "CNPJ": "23.103.347/0001-65",
                "InscEst": "707021792115",
                "Logradouro": "RUA GUSTAVO MARTINS CERQUEIRA",
                "Nro": "255",
                "Bairro": "CENTRO",
                "Cidade": "Urupês",
                "UF": "SP",
                "CEP": "15850-029",
                "Fone": "(17) 35521528",
                "CodigoIBGE": "3556008",
                "RegimeTrib": "3"
            }
    finally:
        conn.close()

def get_pdv_clients(q=None, limit=20):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if q and q.strip():
                search_term = q.strip()
                search = f"%{search_term}%"
                exact = search_term
                prefix = f"{search_term}%"
                sql = """
                    SELECT "CodEntidade", "Nome", "Fantasia", "CPF", "CGC", "Cidade" FROM "ENT" 
                    WHERE ("Tipo" = 1 OR "Tipo" = 7) 
                      AND (CAST("CodEntidade" AS TEXT) = %s OR "Nome" ILIKE %s OR "CPF" ILIKE %s OR "CGC" ILIKE %s OR "Fantasia" ILIKE %s)
                    ORDER BY CASE 
                        WHEN CAST("CodEntidade" AS TEXT) = %s THEN 1 
                        WHEN LOWER("Nome") = LOWER(%s) THEN 2 
                        WHEN "Nome" ILIKE %s THEN 3 
                        WHEN "Fantasia" ILIKE %s THEN 4 
                        ELSE 5 
                    END, "CodEntidade" ASC
                    LIMIT %s
                """
                cursor.execute(sql, (exact, search, search, search, search, exact, exact, prefix, prefix, limit))
            else:
                cursor.execute('SELECT "CodEntidade", "Nome", "Fantasia", "CPF", "CGC", "Cidade" FROM "ENT" WHERE "Tipo" = 1 OR "Tipo" = 7 ORDER BY "CodEntidade" ASC LIMIT %s', (limit,))
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_pdv_sellers(query=None, limit=15):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if query and query.strip():
                search = f"%{query.strip()}%"
                exact = query.strip()
                sql = """
                    SELECT "CodEntidade", "Nome", "CPF", "CGC"
                    FROM "ENT"
                    WHERE ("Tipo" = 3 OR "Tipo" = 7)
                      AND (CAST("CodEntidade" AS TEXT) = %s OR "Nome" ILIKE %s OR "CPF" ILIKE %s OR "CGC" ILIKE %s OR "Fantasia" ILIKE %s)
                    ORDER BY CASE WHEN CAST("CodEntidade" AS TEXT) = %s THEN 1 WHEN "Nome" ILIKE %s THEN 2 ELSE 3 END, "CodEntidade" ASC
                    LIMIT %s
                """
                cursor.execute(sql, (exact, search, search, search, search, exact, f"{exact}%", limit))
            else:
                cursor.execute("""
                    SELECT "CodEntidade", "Nome", "CPF", "CGC"
                    FROM "ENT"
                    WHERE ("Tipo" = 3 OR "Tipo" = 7)
                    ORDER BY "CodEntidade" ASC LIMIT %s
                """, (limit,))
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_pdv_carriers(query=None, limit=50):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if query and query.strip():
                search = f"%{query.strip()}%"
                exact = query.strip()
                sql = """
                    SELECT "CodEntidade", "Nome", "CPF", "CGC"
                    FROM "ENT"
                    WHERE "Tipo" = 5
                      AND (CAST("CodEntidade" AS TEXT) = %s OR "Nome" ILIKE %s OR "CPF" ILIKE %s OR "CGC" ILIKE %s OR "Fantasia" ILIKE %s)
                    ORDER BY CASE WHEN CAST("CodEntidade" AS TEXT) = %s THEN 1 WHEN "Nome" ILIKE %s THEN 2 ELSE 3 END, "CodEntidade" ASC
                    LIMIT %s
                """
                cursor.execute(sql, (exact, search, search, search, search, exact, f"{exact}%", limit))
            else:
                cursor.execute("""
                    SELECT "CodEntidade", "Nome", "CPF", "CGC"
                    FROM "ENT"
                    WHERE "Tipo" = 5
                    ORDER BY "CodEntidade" ASC LIMIT %s
                """, (limit,))
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_next_ped_id():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT MAX("CodPed") as max_id FROM "PED"')
            row = cursor.fetchone()
            max_id = row['max_id'] if row and row['max_id'] else 0
            return max_id + 1
    finally:
        conn.close()

def get_next_itp_id():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT MAX("CodItp") as max_id FROM "ITP"')
            row = cursor.fetchone()
            max_id = row['max_id'] if row and row['max_id'] else 0
            return max_id + 1
    finally:
        conn.close()

def search_pdv_products(query="", grupo=None, limit=60):
    where_clauses = ['(PRD."Ativo" IS NULL OR PRD."Ativo" != 0)']
    params = []

    if grupo is not None and str(grupo) != "" and str(grupo).lower() != "all":
        try:
            where_clauses.append('PRD."Grupo" = %s')
            params.append(int(grupo))
        except (ValueError, TypeError):
            pass

    if query and query.strip():
        search = f"%{query.strip()}%"
        exact = query.strip()
        where_clauses.append(f"""(
            PRD."CodBar" ILIKE %s OR 
            CAST(PRD."CodPrd" AS TEXT) = %s OR 
            PRD."{DESC_PRD_COL}" ILIKE %s OR 
            PRD."Marca" ILIKE %s
        )""")
        params.extend([search, exact, search, search])
        
        # Ordem inteligente
        order_sql = f"""
            ORDER BY 
                CASE 
                    WHEN PRD."CodBar" = %s THEN 1 
                    WHEN CAST(PRD."CodPrd" AS TEXT) = %s THEN 2 
                    ELSE 3 
                END, 
                PRD."{DESC_PRD_COL}" ASC
        """
        order_params = [exact, exact]
    else:
        # Quando busca vazia: prioriza itens com fotos e ordem alfabética
        order_sql = f"""
            ORDER BY 
                CASE 
                    WHEN PRD."Foto" IS NOT NULL AND PRD."Foto" != '' AND PRD."Foto" != '0' AND PRD."Foto" != '-1' THEN 1 
                    ELSE 2 
                END, 
                PRD."{DESC_PRD_COL}" ASC
        """
        order_params = []

    where_sql = " WHERE " + " AND ".join(where_clauses)
    sql = f'''
        SELECT PRD.*, GRU."{DESC_GRU_COL}" as Nome_Grupo, COALESCE(GRU."Desconto", 0.0) as DescontoGrupo
        FROM "PRD" PRD
        LEFT JOIN "GRU" GRU ON PRD."Grupo" = GRU."CodGru"
        {where_sql}
        {order_sql}
        LIMIT %s
    '''

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, params + order_params + [limit])
            rows = []
            for r in cursor.fetchall():
                item = normalize_prd_row(r)
                if DESC_GRU_COL in item and DESC_GRU_COL != "Nome_Grupo":
                    item["Nome_Grupo"] = item.pop(DESC_GRU_COL)
                rows.append(item)
            return rows
    finally:
        conn.close()

def create_pdv_sale(sale_data):
    from app.nfce import emit_nfce
    import datetime

    items = sale_data.get("items", [])
    if not items:
        raise ValueError("O carrinho de compras está vazio.")

    cod_entidade = sale_data.get("CodEntidade", 1)
    vendedor_id = sale_data.get("Vendedor", 1)
    cond_pgto = sale_data.get("CondPgto", "DINHEIRO")
    desconto_total = float(sale_data.get("Desconto", 0.0))
    
    # Calculate item totals and apply automatic group discounts if item discount not explicitly given
    processed_items = []
    calculated_subtotal = 0.0
    
    for it in items:
        cod_prd = it["CodPrd"]
        qtd = float(it.get("Qtd", 1.0))
        v_un = float(it.get("ValorUnit", 0.0))
        
        # Check product group discount
        prd_obj = get_product_by_id(cod_prd)
        desconto_grupo = prd_obj.get("DescontoGrupo", 0.0) if prd_obj else 0.0
        
        v_item_tot = qtd * v_un
        v_item_desc = float(it.get("Desconto", 0.0))
        
        if v_item_desc == 0.0 and desconto_grupo > 0:
            v_item_desc = round(v_item_tot * (desconto_grupo / 100.0), 2)
            
        calculated_subtotal += v_item_tot
        processed_items.append({
            "CodPrd": cod_prd,
            "Qtd": qtd,
            "ValorUnit": v_un,
            "Valor": v_item_tot,
            "Desconto": v_item_desc
        })

    subtotal = float(sale_data.get("SubTotal", calculated_subtotal))
    total_final = float(sale_data.get("Total", subtotal - desconto_total))
    emitir_nfce_flag = sale_data.get("EmitirNFCe", True)

    now = datetime.datetime.now()
    data_emiss = now.strftime("%d/%m/%Y")
    hora = now.strftime("%H:%M:%S")

    cod_ped = get_next_ped_id()
    company = get_company_info()

    entidade = get_entity_by_id(cod_entidade) if cod_entidade else None
    nome_cliente = entidade.get("Nome", "CONSUMIDOR FINAL") if entidade else "CONSUMIDOR FINAL"
    cpf_cliente = entidade.get("CPF", "") or entidade.get("CGC", "") if entidade else ""

    sale_dict_for_nfce = {
        "CodPed": cod_ped,
        "NomeCliente": nome_cliente,
        "CPF": cpf_cliente,
        "CondPgto": cond_pgto,
        "Desconto": desconto_total,
        "Total": total_final
    }

    nfce_res = None
    chave_nfe = ""
    if emitir_nfce_flag:
        nfce_res = emit_nfce(sale_dict_for_nfce, company, processed_items)
        chave_nfe = nfce_res["chave_nfe"]

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            # 1. Insert into PED
            cursor.execute('''
                INSERT INTO "PED" (
                    "CodPed", "Operacao", "Entidade", "DataEmiss", "Total", "SubTotal", "Desconto", 
                    "Vendedor", "CondPgto", "Cfo", "DtSaida", "Hora", "Sat"
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ''', (
                cod_ped, 2, cod_entidade, data_emiss, total_final, subtotal, desconto_total,
                vendedor_id, cond_pgto, "Nfc-e" if emitir_nfce_flag else "VENDA", data_emiss, hora, chave_nfe
            ))

            # 2. Insert into ITP & Deduct Inventory
            next_itp = get_next_itp_id()
            for idx, it in enumerate(processed_items):
                cod_prd = it["CodPrd"]
                qtd = it["Qtd"]
                v_un = it["ValorUnit"]
                v_item_tot = it["Valor"]
                v_item_desc = it["Desconto"]

                cursor.execute('''
                    INSERT INTO "ITP" (
                        "CodItp", "Pedido", "Produto", "Qtd", "ValorUnit", "Valor", "Desconto", "Loja", "VlrVdaDia"
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                ''', (
                    next_itp + idx, cod_ped, cod_prd, qtd, v_un, v_item_tot, v_item_desc, "LOJA", v_un
                ))

                cursor.execute('''
                    UPDATE "PRD" SET "Estoque" = "Estoque" - %s WHERE "CodPrd" = %s
                ''', (qtd, cod_prd))

            # 3. Insert into NFE if NFC-e enabled
            if nfce_res:
                cursor.execute('''
                    INSERT INTO "NFE" (
                        "CodPed", "NroChave", "Status", "Protocolo", "Mensagem", "Recibo"
                    ) VALUES (%s, %s, %s, %s, %s, %s)
                ''', (
                    cod_ped, nfce_res["chave_nfe"], nfce_res["status"],
                    nfce_res["protocolo"], nfce_res["mensagem"], "0"
                ))

            # 4. Insert Cash Movement into CXA if Cash payment
            if cond_pgto == "DINHEIRO":
                cursor.execute('SELECT MAX("CodCxa") as max_cxa FROM "CXA"')
                cxa_row = cursor.fetchone()
                next_cxa = (cxa_row["max_cxa"] or 0) + 1 if cxa_row else 1
                cursor.execute('''
                    INSERT INTO "CXA" ("CodCxa", "NroCaixa", "Historico", "Data", "Hora", "Valor", "D_C")
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                ''', (
                    next_cxa, 1, f"VENDA PDV #{cod_ped}", data_emiss, hora, total_final, "C"
                ))

            conn.commit()
    finally:
        conn.close()

    return get_pdv_sale_by_id(cod_ped)

def get_pdv_sale_by_id(cod_ped):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                SELECT PED.*, ENT."Nome" as "NomeCliente", ENT."CPF", ENT."CGC",
                       ENT."Endereco", ENT."Nro", ENT."Bairro", ENT."Cidade", ENT."Uf", ENT."Cep", ENT."Fone", ENT."InscrEst" as "InscEst"
                FROM "PED" PED
                LEFT JOIN "ENT" ENT ON PED."Entidade" = ENT."CodEntidade"
                WHERE PED."CodPed" = %s
            ''', (cod_ped,))
            ped_row = cursor.fetchone()
            if not ped_row:
                return None
            sale = _convert_row(ped_row)

            cursor.execute(f'''
                SELECT ITP.*, PRD."{DESC_PRD_COL}" as Descricao_Produto, PRD."CodBar", PRD."Embalagem"
                FROM "ITP" ITP
                LEFT JOIN "PRD" PRD ON ITP."Produto" = PRD."CodPrd"
                WHERE ITP."Pedido" = %s
                ORDER BY ITP."CodItp" ASC
            ''', (cod_ped,))
            items = [normalize_prd_row(r) for r in cursor.fetchall()]

            cursor.execute('SELECT * FROM "NFE" WHERE "CodPed" = %s', (cod_ped,))
            nfe_row = cursor.fetchone()
            nfe_data = _convert_row(nfe_row) if nfe_row else None

            company = get_company_info()

            return {
                "venda": sale,
                "itens": items,
                "nfe": nfe_data,
                "empresa": company
            }
    finally:
        conn.close()

# --- SALES ORDERS (PEDIDOS DE VENDA) & NF-E MODELO 55 ---

def get_orders(page=1, limit=20, q=None, data_inicio=None, data_fim=None, status=None):
    offset = (page - 1) * limit
    params = []
    where_clauses = []

    if q and q.strip():
        search = f"%{q.strip()}%"
        where_clauses.append("""(
            CAST(PED."CodPed" AS TEXT) ILIKE %s OR
            ENT."Nome" ILIKE %s OR
            ENT."CPF" ILIKE %s OR
            ENT."CGC" ILIKE %s OR
            PED."NroPedido" ILIKE %s
        )""")
        params.extend([search] * 5)

    if data_inicio:
        where_clauses.append('PED."DataEmiss" >= %s')
        params.append(data_inicio)

    if data_fim:
        where_clauses.append('PED."DataEmiss" <= %s')
        params.append(data_fim)

    if status == "nfe":
        where_clauses.append('NFE."NroChave" IS NOT NULL')
    elif status == "pendente":
        where_clauses.append('NFE."NroChave" IS NULL')

    where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    count_sql = f'''
        SELECT COUNT(DISTINCT PED."CodPed") as total
        FROM "PED" PED
        LEFT JOIN "ENT" ENT ON PED."Entidade" = ENT."CodEntidade"
        LEFT JOIN "NFE" NFE ON PED."CodPed" = NFE."CodPed"
        {where_sql}
    '''

    sql = f'''
        SELECT PED.*, ENT."Nome" as "NomeCliente", ENT."CPF", ENT."CGC", ENT."Cidade", ENT."Fone",
               NFE."NroChave", NFE."Status" as StatusNFe, NFE."Protocolo", NFE."Mensagem"
        FROM "PED" PED
        LEFT JOIN "ENT" ENT ON PED."Entidade" = ENT."CodEntidade"
        LEFT JOIN "NFE" NFE ON PED."CodPed" = NFE."CodPed"
        {where_sql}
        ORDER BY PED."CodPed" DESC
        LIMIT %s OFFSET %s
    '''

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(count_sql, params)
            total = cursor.fetchone()["total"]

            cursor.execute(sql, params + [limit, offset])
            rows = [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

    return {
        "items": rows,
        "total": total,
        "page": page,
        "limit": limit,
        "pages": (total + limit - 1) // limit if limit > 0 else 1
    }

def get_order_by_id(cod_ped):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                SELECT PED.*, ENT."Nome" as "NomeCliente", ENT."CPF", ENT."CGC", ENT."Endereco", ENT."Nro", ENT."Bairro", ENT."Cidade", ENT."Uf", ENT."Cep", ENT."Fone", ENT."Fax", ENT."InscrEst" as "InscEst"
                FROM "PED" PED
                LEFT JOIN "ENT" ENT ON PED."Entidade" = ENT."CodEntidade"
                WHERE PED."CodPed" = %s
            ''', (cod_ped,))
            ped_row = cursor.fetchone()
            if not ped_row:
                return None
            order = _convert_row(ped_row)

            cursor.execute(f'''
                SELECT ITP.*, PRD."{DESC_PRD_COL}" as Descricao_Produto, PRD."CodBar", PRD."Embalagem", PRD."Venda" as PrecoTabela,
                       PRD."ClasseFiscal" as NCM, PRD."SitTrib" as CST, PRD."CfOpPrd" as CFOP,
                       PRD."Icm" as AliqIcms, PRD."Ipi" as AliqIpi
                FROM "ITP" ITP
                LEFT JOIN "PRD" PRD ON ITP."Produto" = PRD."CodPrd"
                WHERE ITP."Pedido" = %s
                ORDER BY ITP."CodItp" ASC
            ''', (cod_ped,))
            items = [normalize_prd_row(r) for r in cursor.fetchall()]

            cursor.execute('SELECT * FROM "NFE" WHERE "CodPed" = %s', (cod_ped,))
            nfe_row = cursor.fetchone()
            nfe_data = _convert_row(nfe_row) if nfe_row else None

            company = get_company_info()

            return {
                "pedido": order,
                "itens": items,
                "nfe": nfe_data,
                "empresa": company
            }
    finally:
        conn.close()

def create_order(order_data):
    import datetime

    items = order_data.get("items", [])
    if not items:
        raise ValueError("O pedido deve conter pelo menos 1 item.")

    cod_entidade = order_data.get("CodEntidade", 1)
    vendedor_id = order_data.get("Vendedor", 1)
    transp_id = order_data.get("Transportadora", 1)
    cond_pgto = order_data.get("CondPgto", "A VISTA")
    cfo = order_data.get("Cfo", "5102 - VENDA DE MERCADORIAS")
    obs = order_data.get("Obs", "")
    desconto_total = float(order_data.get("Desconto", 0.0))
    valor_frete = float(order_data.get("ValorFrete", 0.0))

    calculated_subtotal = sum(float(it.get("Qtd", 1.0)) * float(it.get("ValorUnit", 0.0)) for it in items)
    subtotal = float(order_data.get("SubTotal", calculated_subtotal))
    total_final = float(order_data.get("Total", subtotal - desconto_total + valor_frete))

    now = datetime.datetime.now()
    data_emiss = order_data.get("DataEmiss", now.strftime("%d/%m/%Y"))
    hora = now.strftime("%H:%M:%S")

    cod_ped = get_next_ped_id()

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                INSERT INTO "PED" (
                    "CodPed", "Operacao", "Entidade", "DataEmiss", "Total", "SubTotal", "Desconto", 
                    "Vendedor", "Transportadora", "ValorFrete", "CondPgto", "Cfo", "DtSaida", "Hora", "Obs"
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ''', (
                cod_ped, 1, cod_entidade, data_emiss, total_final, subtotal, desconto_total,
                vendedor_id, transp_id, valor_frete, cond_pgto, cfo, data_emiss, hora, obs
            ))

            next_itp = get_next_itp_id()
            for idx, it in enumerate(items):
                cod_prd = it["CodPrd"]
                qtd = float(it.get("Qtd", 1.0))
                v_un = float(it.get("ValorUnit", 0.0))
                v_item_desc = float(it.get("Desconto", 0.0))
                v_item_tot = float(it.get("Valor", (qtd * v_un) - v_item_desc))

                cursor.execute('''
                    INSERT INTO "ITP" (
                        "CodItp", "Pedido", "Produto", "Qtd", "ValorUnit", "Valor", "Desconto", "Loja", "VlrVdaDia"
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                ''', (
                    next_itp + idx, cod_ped, cod_prd, qtd, v_un, v_item_tot, v_item_desc, "LOJA", v_un
                ))

            conn.commit()
    finally:
        conn.close()

    return get_order_by_id(cod_ped)

def update_order(cod_ped, order_data):
    items = order_data.get("items", [])
    if not items:
        raise ValueError("O pedido deve conter pelo menos 1 item.")

    cod_entidade = order_data.get("CodEntidade", 1)
    vendedor_id = order_data.get("Vendedor", 1)
    transp_id = order_data.get("Transportadora", 1)
    cond_pgto = order_data.get("CondPgto", "A VISTA")
    cfo = order_data.get("Cfo", "5102 - VENDA DE MERCADORIAS")
    obs = order_data.get("Obs", "")
    desconto_total = float(order_data.get("Desconto", 0.0))
    valor_frete = float(order_data.get("ValorFrete", 0.0))

    calculated_subtotal = sum(float(it.get("Qtd", 1.0)) * float(it.get("ValorUnit", 0.0)) for it in items)
    subtotal = float(order_data.get("SubTotal", calculated_subtotal))
    total_final = float(order_data.get("Total", subtotal - desconto_total + valor_frete))

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                UPDATE "PED" SET
                    "Entidade" = %s, "Total" = %s, "SubTotal" = %s, "Desconto" = %s,
                    "Vendedor" = %s, "Transportadora" = %s, "ValorFrete" = %s,
                    "CondPgto" = %s, "Cfo" = %s, "Obs" = %s
                WHERE "CodPed" = %s
            ''', (
                cod_entidade, total_final, subtotal, desconto_total,
                vendedor_id, transp_id, valor_frete, cond_pgto, cfo, obs, cod_ped
            ))

            cursor.execute('DELETE FROM "ITP" WHERE "Pedido" = %s', (cod_ped,))

            next_itp = get_next_itp_id()
            for idx, it in enumerate(items):
                cod_prd = it["CodPrd"]
                qtd = float(it.get("Qtd", 1.0))
                v_un = float(it.get("ValorUnit", 0.0))
                v_item_desc = float(it.get("Desconto", 0.0))
                v_item_tot = float(it.get("Valor", (qtd * v_un) - v_item_desc))

                cursor.execute('''
                    INSERT INTO "ITP" (
                        "CodItp", "Pedido", "Produto", "Qtd", "ValorUnit", "Valor", "Desconto", "Loja", "VlrVdaDia"
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                ''', (
                    next_itp + idx, cod_ped, cod_prd, qtd, v_un, v_item_tot, v_item_desc, "LOJA", v_un
                ))

            conn.commit()
    finally:
        conn.close()

    return get_order_by_id(cod_ped)

def delete_order(cod_ped):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "ITP" WHERE "Pedido" = %s', (cod_ped,))
            cursor.execute('DELETE FROM "NFE" WHERE "CodPed" = %s', (cod_ped,))
            cursor.execute('DELETE FROM "NOTA" WHERE "CodPed" = %s', (cod_ped,))
            cursor.execute('DELETE FROM "PED" WHERE "CodPed" = %s', (cod_ped,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()

def update_nfe_order_details(cod_ped, data):
    """
    Atualiza os parâmetros específicos de emissão de NF-e (Modelo 55) baseados no FrmNota:
    - Transportadora, ModFrete, Placa, UF, QtdVol, Especie, Marca, PesoBruto, PesoLiquido
    - CFO / Natureza de Operação, ObsRodape, ObsCorpo
    """
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                UPDATE "PED" SET
                    "Cfo" = COALESCE(%s, "Cfo"),
                    "Transportadora" = COALESCE(%s, "Transportadora"),
                    "ValorFrete" = COALESCE(%s, "ValorFrete"),
                    "Desconto" = COALESCE(%s, "Desconto"),
                    "Obs" = COALESCE(%s, "Obs")
                WHERE "CodPed" = %s
            ''', (
                data.get("Cfo"),
                data.get("Transportadora"),
                data.get("ValorFrete"),
                data.get("Desconto"),
                data.get("ObsRodape") or data.get("Obs"),
                cod_ped
            ))
            conn.commit()
    finally:
        conn.close()
    return get_order_by_id(cod_ped)

def emit_nfe_from_order(cod_ped, custom_data=None):
    from app.nfe_engine import emit_nfe_55

    order_details = get_order_by_id(cod_ped)
    if not order_details or not order_details.get("pedido"):
        raise ValueError(f"Pedido #{cod_ped} não foi encontrado.")

    order = order_details["pedido"]
    if custom_data and isinstance(custom_data, dict):
        order.update(custom_data)

    items = order_details["itens"]
    company = order_details["empresa"]

    nfe_res = emit_nfe_55(order, company, items)

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "NFE" WHERE "CodPed" = %s', (cod_ped,))
            cursor.execute('''
                INSERT INTO "NFE" (
                    "CodPed", "NroChave", "Status", "Protocolo", "Mensagem", "Recibo"
                ) VALUES (%s, %s, %s, %s, %s, %s)
            ''', (
                cod_ped, nfe_res["chave_nfe"], nfe_res["status"],
                nfe_res["protocolo"], nfe_res["mensagem"], "0"
            ))

            cursor.execute('''
                UPDATE "PED" SET "Sat" = %s, "Cfo" = %s WHERE "CodPed" = %s
            ''', (nfe_res["chave_nfe"], "NFe-55", cod_ped))

            conn.commit()
    finally:
        conn.close()

    res = get_order_by_id(cod_ped)
    if res:
        res["nfe_result"] = nfe_res
    return res


# --- CFOP (TABELA CFO) FUNCTIONS ---

def get_cfops(q=None, limit=50):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if q and q.strip():
                search = f"%{q.strip()}%"
                cursor.execute('''
                    SELECT * FROM "CFO"
                    WHERE "Codigo" ILIKE %s OR "Descricao" ILIKE %s
                    ORDER BY "Codigo" ASC
                    LIMIT %s
                ''', (search, search, limit))
            else:
                cursor.execute('SELECT * FROM "CFO" ORDER BY "Codigo" ASC LIMIT %s', (limit,))
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_cfop_by_code(codigo):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "CFO" WHERE "Codigo" = %s', (str(codigo),))
            r = cursor.fetchone()
            return _convert_row(r) if r else None
    finally:
        conn.close()

def save_cfop(data):
    codigo = str(data.get("Codigo", "")).strip()
    descricao = str(data.get("Descricao", "")).strip()
    if not codigo or not descricao:
        raise ValueError("Código CFOP e Descrição são obrigatórios.")

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "CFO" WHERE "Codigo" = %s', (codigo,))
            exists = cursor.fetchone()
            if exists:
                cursor.execute('''
                    UPDATE "CFO" SET "Descricao" = %s WHERE "Codigo" = %s
                ''', (descricao, codigo))
            else:
                cursor.execute('''
                    INSERT INTO "CFO" ("Codigo", "Descricao", "id_empresa")
                    VALUES (%s, %s, %s)
                ''', (codigo, descricao, 1))
            conn.commit()
    finally:
        conn.close()
    return get_cfop_by_code(codigo)

def delete_cfop(codigo):
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "CFO" WHERE "Codigo" = %s', (str(codigo),))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()


# --- FORMAS E CONDIÇÕES DE PAGAMENTO (TABELA FormaPgto) FUNCTIONS ---

def init_forma_pgto_table():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS "FormaPgto" (
                    "id_forma" SERIAL PRIMARY KEY,
                    "Codigo" VARCHAR(20) NOT NULL,
                    "Nome" VARCHAR(100) NOT NULL,
                    "Tipo" VARCHAR(50) DEFAULT 'A_VISTA',
                    "Parcelas" INTEGER DEFAULT 1,
                    "DiasEntreParcelas" INTEGER DEFAULT 0,
                    "Ativo" INTEGER DEFAULT 1,
                    "id_empresa" INTEGER DEFAULT 1
                );
            ''')
            
            cursor.execute('SELECT COUNT(*) as cnt FROM "FormaPgto"')
            cnt = cursor.fetchone()["cnt"]
            if cnt == 0:
                default_formas = [
                    ("DINHEIRO", "DINHEIRO", "DINHEIRO", 1, 0, 1),
                    ("PIX", "PIX / TRANSFERÊNCIA BANCÁRIA", "PIX", 1, 0, 1),
                    ("CARTAO_CRED", "CARTÃO DE CRÉDITO", "CARTAO", 1, 0, 1),
                    ("CARTAO_DEB", "CARTÃO DE DÉBITO", "CARTAO", 1, 0, 1),
                    ("A_VISTA", "A VISTA", "DINHEIRO", 1, 0, 1),
                    ("30_DIAS", "30 DIAS", "PRAZO", 1, 30, 1),
                    ("30_60_DIAS", "30/60 DIAS", "PRAZO", 2, 30, 1),
                    ("30_60_90_DIAS", "30/60/90 DIAS", "PRAZO", 3, 30, 1),
                    ("BOLETO", "BOLETO BANCÁRIO", "BOLETO", 1, 30, 1),
                ]
                cursor.executemany('''
                    INSERT INTO "FormaPgto" ("Codigo", "Nome", "Tipo", "Parcelas", "DiasEntreParcelas", "Ativo")
                    VALUES (%s, %s, %s, %s, %s, %s)
                ''', default_formas)

            conn.commit()
    finally:
        conn.close()

def get_formas_pgto(include_inactive=False, q=None):
    init_forma_pgto_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            where_clauses = []
            params = []
            if not include_inactive:
                where_clauses.append('"Ativo" = 1')
            if q and q.strip():
                search = f"%{q.strip()}%"
                where_clauses.append('("Codigo" ILIKE %s OR "Nome" ILIKE %s)')
                params.extend([search, search])

            where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""
            cursor.execute(f'SELECT * FROM "FormaPgto" {where_sql} ORDER BY "id_forma" ASC', params)
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_forma_pgto_by_id(id_forma):
    init_forma_pgto_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "FormaPgto" WHERE "id_forma" = %s', (id_forma,))
            r = cursor.fetchone()
            return _convert_row(r) if r else None
    finally:
        conn.close()

def save_forma_pgto(data):
    init_forma_pgto_table()
    id_forma = data.get("id_forma")
    if id_forma:
        try:
            id_forma = int(id_forma)
        except (ValueError, TypeError):
            id_forma = None

    codigo = str(data.get("Codigo", "")).strip() or "GERAL"
    nome = str(data.get("Nome", "")).strip()
    tipo = str(data.get("Tipo", "A_VISTA")).strip()
    parcelas = int(data.get("Parcelas", 1))
    dias = int(data.get("DiasEntreParcelas", 0))
    ativo = 1 if str(data.get("Ativo", "1")) in ["1", "true", "True", "on"] else 0

    if not nome:
        raise ValueError("O Nome da forma de pagamento é obrigatório.")

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if id_forma:
                cursor.execute('''
                    UPDATE "FormaPgto" SET
                        "Codigo" = %s, "Nome" = %s, "Tipo" = %s,
                        "Parcelas" = %s, "DiasEntreParcelas" = %s, "Ativo" = %s
                    WHERE "id_forma" = %s
                ''', (codigo, nome, tipo, parcelas, dias, ativo, id_forma))
                saved_id = id_forma
            else:
                cursor.execute('''
                    INSERT INTO "FormaPgto" ("Codigo", "Nome", "Tipo", "Parcelas", "DiasEntreParcelas", "Ativo", "id_empresa")
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    RETURNING "id_forma"
                ''', (codigo, nome, tipo, parcelas, dias, ativo, 1))
                saved_id = cursor.fetchone()["id_forma"]

            conn.commit()
    finally:
        conn.close()
    return get_forma_pgto_by_id(saved_id)

def delete_forma_pgto(id_forma):
    init_forma_pgto_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "FormaPgto" WHERE "id_forma" = %s', (id_forma,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()


# --- NATUREZA DA OPERAÇÃO (TABELA NaturezaOperacao) FUNCTIONS ---

def init_natureza_operacao_table():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS "NaturezaOperacao" (
                    "id_natureza" SERIAL PRIMARY KEY,
                    "Codigo" VARCHAR(20) NOT NULL,
                    "Descricao" VARCHAR(150) NOT NULL,
                    "CFOP" VARCHAR(10) DEFAULT '5102',
                    "Ativo" INTEGER DEFAULT 1,
                    "id_empresa" INTEGER DEFAULT 1
                );
            ''')
            
            cursor.execute('SELECT COUNT(*) as cnt FROM "NaturezaOperacao"')
            cnt = cursor.fetchone()["cnt"]
            if cnt == 0:
                default_naturezas = [
                    ("VENDA_MERC", "VENDA DE MERCADORIAS ADQUIRIDAS DE TERCEIROS", "5102", 1),
                    ("VENDA_PROD", "VENDA DE PRODUÇÃO DO ESTABELECIMENTO", "5101", 1),
                    ("BONIFICACAO", "REMESSA EM BONIFICAÇÃO, DOAÇÃO OU BRINDE", "5910", 1),
                    ("DEVOLUCAO", "DEVOLUÇÃO DE COMPRA PARA INDUSTRIALIZAÇÃO OU COMERCIALIZAÇÃO", "5202", 1),
                    ("VENDA_FORA_ESTADO", "VENDA DE MERCADORIAS ADQUIRIDAS DE TERCEIROS FORA DO ESTADO", "6102", 1),
                ]
                cursor.executemany('''
                    INSERT INTO "NaturezaOperacao" ("Codigo", "Descricao", "CFOP", "Ativo")
                    VALUES (%s, %s, %s, %s)
                ''', default_naturezas)

            conn.commit()
    finally:
        conn.close()

def get_naturezas_operacao(include_inactive=False, q=None):
    init_natureza_operacao_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            where_clauses = []
            params = []
            if not include_inactive:
                where_clauses.append('"Ativo" = 1')
            if q and q.strip():
                search = f"%{q.strip()}%"
                where_clauses.append('("Codigo" ILIKE %s OR "Descricao" ILIKE %s OR "CFOP" ILIKE %s)')
                params.extend([search, search, search])

            where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""
            cursor.execute(f'SELECT * FROM "NaturezaOperacao" {where_sql} ORDER BY "id_natureza" ASC', params)
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_natureza_operacao_by_id(id_natureza):
    init_natureza_operacao_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM "NaturezaOperacao" WHERE "id_natureza" = %s', (id_natureza,))
            r = cursor.fetchone()
            return _convert_row(r) if r else None
    finally:
        conn.close()

def save_natureza_operacao(data):
    init_natureza_operacao_table()
    id_natureza = data.get("id_natureza")
    if id_natureza:
        try:
            id_natureza = int(id_natureza)
        except (ValueError, TypeError):
            id_natureza = None

    codigo = str(data.get("Codigo", "")).strip() or "GERAL"
    descricao = str(data.get("Descricao", "")).strip()
    cfop = str(data.get("CFOP", "5102")).strip()
    ativo = 1 if str(data.get("Ativo", "1")) in ["1", "true", "True", "on"] else 0

    if not descricao:
        raise ValueError("A Descrição da Natureza de Operação é obrigatória.")

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if id_natureza:
                cursor.execute('''
                    UPDATE "NaturezaOperacao" SET
                        "Codigo" = %s, "Descricao" = %s, "CFOP" = %s, "Ativo" = %s
                    WHERE "id_natureza" = %s
                ''', (codigo, descricao, cfop, ativo, id_natureza))
                saved_id = id_natureza
            else:
                cursor.execute('''
                    INSERT INTO "NaturezaOperacao" ("Codigo", "Descricao", "CFOP", "Ativo", "id_empresa")
                    VALUES (%s, %s, %s, %s, %s)
                    RETURNING "id_natureza"
                ''', (codigo, descricao, cfop, ativo, 1))
                saved_id = cursor.fetchone()["id_natureza"]

            conn.commit()
    finally:
        conn.close()
    return get_natureza_operacao_by_id(saved_id)

def delete_natureza_operacao(id_natureza):
    init_natureza_operacao_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "NaturezaOperacao" WHERE "id_natureza" = %s', (id_natureza,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()


# --- USUÁRIOS, LOGIN & PERFIS FUNCTIONS ---

import hashlib

def _hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def init_usuario_table():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS "Usuario" (
                    "id_usuario" SERIAL PRIMARY KEY,
                    "Login" VARCHAR(50) UNIQUE NOT NULL,
                    "Nome" VARCHAR(100) NOT NULL,
                    "Senha" VARCHAR(255) NOT NULL,
                    "Perfil" VARCHAR(50) DEFAULT 'OPERADOR',
                    "Ativo" INTEGER DEFAULT 1,
                    "DtCadastro" VARCHAR(20),
                    "id_empresa" INTEGER DEFAULT 1
                );
            ''')
            
            cursor.execute('SELECT COUNT(*) as cnt FROM "Usuario"')
            cnt = cursor.fetchone()["cnt"]
            if cnt == 0:
                admin_pass = _hash_password("admin123")
                cursor.execute('''
                    INSERT INTO "Usuario" ("Login", "Nome", "Senha", "Perfil", "Ativo", "DtCadastro")
                    VALUES (%s, %s, %s, %s, %s, CURRENT_DATE::text)
                ''', ("admin", "ADMINISTRADOR DO SISTEMA", admin_pass, "ADMIN", 1))

            conn.commit()
    finally:
        conn.close()

def authenticate_user(login, senha):
    init_usuario_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            hashed = _hash_password(senha)
            cursor.execute('SELECT * FROM "Usuario" WHERE LOWER("Login") = LOWER(%s) AND "Ativo" = 1', (login.strip(),))
            user = cursor.fetchone()
            if not user:
                return None
            
            user_dict = _convert_row(user)
            db_pass = user_dict.get("Senha", "")
            if db_pass == hashed or db_pass == senha:
                user_dict.pop("Senha", None)
                return user_dict
            return None
    finally:
        conn.close()

def get_usuarios(q=None):
    init_usuario_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            where_sql = ""
            params = []
            if q and q.strip():
                search = f"%{q.strip()}%"
                where_sql = ' WHERE ("Login" ILIKE %s OR "Nome" ILIKE %s OR "Perfil" ILIKE %s)'
                params = [search, search, search]

            cursor.execute(f'SELECT "id_usuario", "Login", "Nome", "Perfil", "Ativo", "DtCadastro", "id_empresa" FROM "Usuario" {where_sql} ORDER BY "id_usuario" ASC', params)
            return [_convert_row(r) for r in cursor.fetchall()]
    finally:
        conn.close()

def get_usuario_by_id(id_usuario):
    init_usuario_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT "id_usuario", "Login", "Nome", "Perfil", "Ativo", "DtCadastro", "id_empresa" FROM "Usuario" WHERE "id_usuario" = %s', (id_usuario,))
            r = cursor.fetchone()
            return _convert_row(r) if r else None
    finally:
        conn.close()

def save_usuario(data):
    init_usuario_table()
    id_usuario = data.get("id_usuario")
    if id_usuario:
        try:
            id_usuario = int(id_usuario)
        except (ValueError, TypeError):
            id_usuario = None

    login = str(data.get("Login", "")).strip().lower()
    nome = str(data.get("Nome", "")).strip()
    senha = str(data.get("Senha", "")).strip()
    perfil = str(data.get("Perfil", "OPERADOR")).strip().upper()
    ativo = 1 if str(data.get("Ativo", "1")) in ["1", "true", "True", "on"] else 0

    if not login or not nome:
        raise ValueError("Login e Nome são obrigatórios.")

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            if id_usuario:
                cursor.execute('SELECT "id_usuario" FROM "Usuario" WHERE LOWER("Login") = LOWER(%s) AND "id_usuario" != %s', (login, id_usuario))
            else:
                cursor.execute('SELECT "id_usuario" FROM "Usuario" WHERE LOWER("Login") = LOWER(%s)', (login,))
            if cursor.fetchone():
                raise ValueError(f"O login '{login}' já está cadastrado no sistema.")

            if id_usuario:
                if senha:
                    hashed = _hash_password(senha)
                    cursor.execute('''
                        UPDATE "Usuario" SET
                            "Login" = %s, "Nome" = %s, "Senha" = %s, "Perfil" = %s, "Ativo" = %s
                        WHERE "id_usuario" = %s
                    ''', (login, nome, hashed, perfil, ativo, id_usuario))
                else:
                    cursor.execute('''
                        UPDATE "Usuario" SET
                            "Login" = %s, "Nome" = %s, "Perfil" = %s, "Ativo" = %s
                        WHERE "id_usuario" = %s
                    ''', (login, nome, perfil, ativo, id_usuario))
                saved_id = id_usuario
            else:
                if not senha:
                    senha = "123"
                hashed = _hash_password(senha)
                cursor.execute('''
                    INSERT INTO "Usuario" ("Login", "Nome", "Senha", "Perfil", "Ativo", "DtCadastro", "id_empresa")
                    VALUES (%s, %s, %s, %s, %s, CURRENT_DATE::text, %s)
                    RETURNING "id_usuario"
                ''', (login, nome, hashed, perfil, ativo, 1))
                saved_id = cursor.fetchone()["id_usuario"]

            conn.commit()
    finally:
        conn.close()
    return get_usuario_by_id(saved_id)

def delete_usuario(id_usuario):
    init_usuario_table()
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM "Usuario" WHERE "id_usuario" = %s', (id_usuario,))
            conn.commit()
            return cursor.rowcount > 0
    finally:
        conn.close()


# --- DATABASE DIAGNOSTICS & CONFIGURATION FUNCTIONS ---

def test_db_connection_params(host, port, user, password, dbname):
    try:
        test_conn = psycopg2.connect(
            host=host,
            port=int(port),
            user=user,
            password=password,
            dbname=dbname,
            connect_timeout=3,
            cursor_factory=RealDictCursor
        )
        with test_conn.cursor() as cursor:
            cursor.execute('SELECT version() as version')
            ver = cursor.fetchone()["version"]
            
            cursor.execute("SELECT COUNT(*) as tbl_count FROM information_schema.tables WHERE table_schema = 'public'")
            tbl_cnt = cursor.fetchone()["tbl_count"]
            
        test_conn.close()
        return {
            "success": True,
            "message": "Conexão com PostgreSQL estabelecida com sucesso!",
            "version": ver,
            "tables_count": tbl_cnt,
            "host": host,
            "port": port,
            "dbname": dbname,
            "user": user
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"Erro na conexão com o banco de dados: {str(e)}"
        }

def get_db_info():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute('SELECT version() as version')
            ver = cursor.fetchone()["version"]
            
            cursor.execute("SELECT COUNT(*) as tbl_count FROM information_schema.tables WHERE table_schema = 'public'")
            tbl_cnt = cursor.fetchone()["tbl_count"]

            cursor.execute('SELECT COUNT(*) as cnt FROM "ENT"')
            ent_cnt = cursor.fetchone()["cnt"]

            cursor.execute('SELECT COUNT(*) as cnt FROM "PRD"')
            prd_cnt = cursor.fetchone()["cnt"]

            cursor.execute('SELECT COUNT(*) as cnt FROM "PED"')
            ped_cnt = cursor.fetchone()["cnt"]

            return {
                "host": PG_HOST,
                "port": PG_PORT,
                "user": PG_USER,
                "dbname": PG_DB,
                "version": ver,
                "tables_count": tbl_cnt,
                "records_summary": {
                    "entidades": ent_cnt,
                    "produtos": prd_cnt,
                    "pedidos": ped_cnt
                }
            }
    finally:
        conn.close()


# --- RELATÓRIO DE VENDAS AGRUPADO POR PRODUTO ---

def _parse_to_iso_date(d_str):
    import re
    if not d_str:
        return None
    d_str = str(d_str).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}$', d_str):
        return d_str
    if re.match(r'^\d{2}/\d{2}/\d{4}$', d_str):
        parts = d_str.split('/')
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return None

def get_relatorio_vendas_produto(data_inicio=None, data_fim=None, grupo=None, q=None, id_empresa=None, status="ativos", sort_by="total_valor", sort_order="DESC", page=1, limit=50):
    params = []
    where_clauses = [
        """PED."DataEmiss" IS NOT NULL AND PED."DataEmiss" != ''"""
    ]
    
    # Status do Pedido e Emissão Fiscal
    st = str(status).lower() if status else "ativos"
    if st in ("ativos", "todos_ativos"):
        where_clauses.append("""(
            (PED."PrevEntrega" IS NULL OR UPPER(PED."PrevEntrega") != 'CANCELADO') AND 
            (PED."Cfo" IS NULL OR UPPER(PED."Cfo") NOT LIKE %s) AND
            (PED."Operacao" != 4 OR PED."Operacao" IS NULL)
        )""")
        params.append('%CANCEL%')
    elif st in ("com_nota", "faturados"):
        where_clauses.append("""(
            (PED."PrevEntrega" IS NULL OR UPPER(PED."PrevEntrega") != 'CANCELADO') AND 
            (PED."Cfo" IS NULL OR UPPER(PED."Cfo") NOT LIKE %s) AND
            (PED."Operacao" != 4 OR PED."Operacao" IS NULL) AND
            (COALESCE(PED."NroNt", 0) > 0 OR (NFE."NroChave" IS NOT NULL AND NFE."NroChave" != ''))
        )""")
        params.append('%CANCEL%')
    elif st in ("sem_nota", "pendentes"):
        where_clauses.append("""(
            (PED."PrevEntrega" IS NULL OR UPPER(PED."PrevEntrega") != 'CANCELADO') AND 
            (PED."Cfo" IS NULL OR UPPER(PED."Cfo") NOT LIKE %s) AND
            (PED."Operacao" != 4 OR PED."Operacao" IS NULL) AND
            COALESCE(PED."NroNt", 0) = 0 AND 
            (NFE."NroChave" IS NULL OR NFE."NroChave" = '')
        )""")
        params.append('%CANCEL%')
    elif st in ("cancelados", "cancel"):
        where_clauses.append("""(
            UPPER(PED."PrevEntrega") = 'CANCELADO' OR 
            UPPER(PED."Cfo") LIKE %s OR
            PED."Operacao" = 4
        )""")
        params.append('%CANCEL%')
    # "todos" não adiciona cláusula de status

    # Filtro por Período de Emissão
    iso_inicio = _parse_to_iso_date(data_inicio)
    if iso_inicio:
        where_clauses.append("TO_DATE(PED.\"DataEmiss\", 'DD/MM/YYYY') >= %s::date")
        params.append(iso_inicio)

    iso_fim = _parse_to_iso_date(data_fim)
    if iso_fim:
        where_clauses.append("TO_DATE(PED.\"DataEmiss\", 'DD/MM/YYYY') <= %s::date")
        params.append(iso_fim)

    # Filtro de Empresa / Filial
    if id_empresa is not None and str(id_empresa).lower() not in ('all', ''):
        try:
            where_clauses.append('PED."id_empresa" = %s')
            params.append(int(id_empresa))
        except ValueError:
            pass

    # Filtro de Grupo de Produtos
    if grupo is not None and str(grupo).lower() not in ('all', ''):
        try:
            where_clauses.append('PRD."Grupo" = %s')
            params.append(int(grupo))
        except ValueError:
            pass

    # Busca por Nome, Código, Código de Barras ou Marca
    if q and q.strip():
        search = f"%{q.strip()}%"
        where_clauses.append(f"""(
            CAST(PRD."CodPrd" AS TEXT) ILIKE %s OR
            PRD."{DESC_PRD_COL}" ILIKE %s OR
            PRD."CodBar" ILIKE %s OR
            PRD."Marca" ILIKE %s
        )""")
        params.extend([search] * 4)

    where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    # Mapeamento de Ordenação
    sort_map = {
        "total_valor": "total_valor",
        "total_qtd": "total_qtd",
        "total_pedidos": "total_pedidos",
        "preco_medio": "preco_medio",
        "lucro_bruto": "lucro_bruto",
        "margem_pct": "margem_pct",
        "descricao": f'PRD."{DESC_PRD_COL}"',
        "codprd": 'PRD."CodPrd"'
    }
    sort_col = sort_map.get(sort_by.lower(), "total_valor")
    sort_dir = "DESC" if sort_order.upper() == "DESC" else "ASC"

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            # 1. Totalizadores Globais (KPIs) para o escopo filtrado
            summary_sql = f"""
                SELECT 
                    COUNT(DISTINCT ITP."Produto") as total_produtos_distintos,
                    COUNT(DISTINCT PED."CodPed") as total_pedidos_geral,
                    COUNT(DISTINCT CASE WHEN (COALESCE(PED."NroNt", 0) > 0 OR (NFE."NroChave" IS NOT NULL AND NFE."NroChave" != '')) THEN PED."CodPed" END) as total_pedidos_com_nota,
                    COUNT(DISTINCT CASE WHEN (COALESCE(PED."NroNt", 0) = 0 AND (NFE."NroChave" IS NULL OR NFE."NroChave" = '')) THEN PED."CodPed" END) as total_pedidos_sem_nota,
                    COALESCE(SUM(ITP."Qtd"), 0) as total_qtd_geral,
                    COALESCE(SUM(ITP."Qtd" * ITP."ValorUnit"), 0) as total_bruto_geral,
                    COALESCE(SUM(ITP."Desconto"), 0) as total_desconto_geral,
                    COALESCE(SUM(ITP."Valor"), 0) as total_faturamento_geral,
                    COALESCE(SUM(ITP."Qtd" * COALESCE(PRD."Custo", 0)), 0) as total_custo_geral
                FROM "ITP" ITP
                JOIN "PED" PED ON ITP."Pedido" = PED."CodPed"
                LEFT JOIN "NFE" NFE ON PED."CodPed" = NFE."CodPed"
                LEFT JOIN "PRD" PRD ON ITP."Produto" = PRD."CodPrd"
                {where_sql}
            """
            cursor.execute(summary_sql, params)
            sum_row = _convert_row(cursor.fetchone())
            
            faturamento_geral = float(sum_row.get("total_faturamento_geral", 0.0))
            custo_geral = float(sum_row.get("total_custo_geral", 0.0))
            qtd_geral = float(sum_row.get("total_qtd_geral", 0.0))
            lucro_geral = faturamento_geral - custo_geral
            margem_geral = (lucro_geral / faturamento_geral * 100.0) if faturamento_geral > 0 else 0.0
            preco_medio_geral = (faturamento_geral / qtd_geral) if qtd_geral > 0 else 0.0

            summary = {
                "total_produtos_distintos": int(sum_row.get("total_produtos_distintos", 0)),
                "total_pedidos_geral": int(sum_row.get("total_pedidos_geral", 0)),
                "total_pedidos_com_nota": int(sum_row.get("total_pedidos_com_nota", 0)),
                "total_pedidos_sem_nota": int(sum_row.get("total_pedidos_sem_nota", 0)),
                "total_qtd_geral": qtd_geral,
                "total_bruto_geral": float(sum_row.get("total_bruto_geral", 0.0)),
                "total_desconto_geral": float(sum_row.get("total_desconto_geral", 0.0)),
                "total_faturamento_geral": faturamento_geral,
                "total_custo_geral": custo_geral,
                "lucro_bruto_geral": lucro_geral,
                "margem_lucro_pct_geral": round(margem_geral, 2),
                "preco_medio_geral": round(preco_medio_geral, 2)
            }

            # 2. Agrupamento por Produto
            offset = (page - 1) * limit if page > 0 and limit > 0 else 0
            limit_clause = f"LIMIT {limit} OFFSET {offset}" if limit > 0 else ""
            
            agg_query = f"""
                SELECT 
                    PRD."CodPrd",
                    COALESCE(PRD."{DESC_PRD_COL}", 'PRODUTO NÃO IDENTIFICADO') as "Descricao_Produto",
                    PRD."CodBar",
                    PRD."Foto",
                    COALESCE(PRD."Embalagem", 'UN') as "Embalagem",
                    COALESCE(PRD."Custo", 0.0) as "Custo",
                    COALESCE(PRD."Venda", 0.0) as "PrecoTabela",
                    COALESCE(PRD."Estoque", 0.0) as "EstoqueAtual",
                    GRU."CodGru",
                    COALESCE(GRU."{DESC_GRU_COL}", 'SEM GRUPO') as "Descricao_Grupo",
                    COUNT(DISTINCT PED."CodPed") as total_pedidos,
                    SUM(ITP."Qtd") as total_qtd,
                    SUM(ITP."Qtd" * ITP."ValorUnit") as total_bruto,
                    SUM(ITP."Desconto") as total_desconto,
                    SUM(ITP."Valor") as total_valor,
                    (SUM(ITP."Valor") / NULLIF(SUM(ITP."Qtd"), 0)) as preco_medio,
                    SUM(ITP."Qtd" * COALESCE(PRD."Custo", 0.0)) as custo_total,
                    (SUM(ITP."Valor") - SUM(ITP."Qtd" * COALESCE(PRD."Custo", 0.0))) as lucro_bruto,
                    CASE 
                        WHEN SUM(ITP."Valor") > 0 THEN 
                            ((SUM(ITP."Valor") - SUM(ITP."Qtd" * COALESCE(PRD."Custo", 0.0))) / SUM(ITP."Valor")) * 100.0
                        ELSE 0.0 
                    END as margem_pct
                FROM "ITP" ITP
                JOIN "PED" PED ON ITP."Pedido" = PED."CodPed"
                LEFT JOIN "NFE" NFE ON PED."CodPed" = NFE."CodPed"
                LEFT JOIN "PRD" PRD ON ITP."Produto" = PRD."CodPrd"
                LEFT JOIN "GRU" GRU ON PRD."Grupo" = GRU."CodGru"
                {where_sql}
                GROUP BY 
                    PRD."CodPrd", PRD."{DESC_PRD_COL}", PRD."CodBar", PRD."Foto",
                    PRD."Embalagem", PRD."Custo", PRD."Venda", PRD."Estoque",
                    GRU."CodGru", GRU."{DESC_GRU_COL}"
                ORDER BY {sort_col} {sort_dir}
                {limit_clause}
            """
            cursor.execute(agg_query, params)
            raw_items = cursor.fetchall()

            items = []
            for r in raw_items:
                item = _convert_row(r)
                val_tot = float(item.get("total_valor", 0.0))
                item["participacao_pct"] = round((val_tot / faturamento_geral * 100.0), 2) if faturamento_geral > 0 else 0.0
                item["margem_pct"] = round(float(item.get("margem_pct", 0.0)), 2)
                item["lucro_bruto"] = round(float(item.get("lucro_bruto", 0.0)), 2)
                item["preco_medio"] = round(float(item.get("preco_medio", 0.0)), 2)
                item["total_bruto"] = round(float(item.get("total_bruto", 0.0)), 2)
                item["total_desconto"] = round(float(item.get("total_desconto", 0.0)), 2)
                item["total_valor"] = round(val_tot, 2)
                item["custo_total"] = round(float(item.get("custo_total", 0.0)), 2)
                items.append(item)

            total_distintos = summary["total_produtos_distintos"]
            pages = (total_distintos + limit - 1) // limit if limit > 0 else 1

            return {
                "summary": summary,
                "items": items,
                "page": page,
                "limit": limit,
                "total": total_distintos,
                "pages": pages
            }
    finally:
        conn.close()

def get_relatorio_vendas_produto_detalhes(cod_prd, data_inicio=None, data_fim=None, id_empresa=None, status="ativos"):
    params = [cod_prd]
    where_clauses = [
        'ITP."Produto" = %s',
        """PED."DataEmiss" IS NOT NULL AND PED."DataEmiss" != ''"""
    ]

    st = str(status).lower() if status else "ativos"
    if st in ("ativos", "todos_ativos"):
        where_clauses.append("""(
            (PED."PrevEntrega" IS NULL OR UPPER(PED."PrevEntrega") != 'CANCELADO') AND 
            (PED."Cfo" IS NULL OR UPPER(PED."Cfo") NOT LIKE %s) AND
            (PED."Operacao" != 4 OR PED."Operacao" IS NULL)
        )""")
        params.append('%CANCEL%')
    elif st in ("com_nota", "faturados"):
        where_clauses.append("""(
            (PED."PrevEntrega" IS NULL OR UPPER(PED."PrevEntrega") != 'CANCELADO') AND 
            (PED."Cfo" IS NULL OR UPPER(PED."Cfo") NOT LIKE %s) AND
            (PED."Operacao" != 4 OR PED."Operacao" IS NULL) AND
            (COALESCE(PED."NroNt", 0) > 0 OR (NFE."NroChave" IS NOT NULL AND NFE."NroChave" != ''))
        )""")
        params.append('%CANCEL%')
    elif st in ("sem_nota", "pendentes"):
        where_clauses.append("""(
            (PED."PrevEntrega" IS NULL OR UPPER(PED."PrevEntrega") != 'CANCELADO') AND 
            (PED."Cfo" IS NULL OR UPPER(PED."Cfo") NOT LIKE %s) AND
            (PED."Operacao" != 4 OR PED."Operacao" IS NULL) AND
            COALESCE(PED."NroNt", 0) = 0 AND 
            (NFE."NroChave" IS NULL OR NFE."NroChave" = '')
        )""")
        params.append('%CANCEL%')
    elif st in ("cancelados", "cancel"):
        where_clauses.append("""(
            UPPER(PED."PrevEntrega") = 'CANCELADO' OR 
            UPPER(PED."Cfo") LIKE %s OR
            PED."Operacao" = 4
        )""")
        params.append('%CANCEL%')

    iso_inicio = _parse_to_iso_date(data_inicio)
    if iso_inicio:
        where_clauses.append("TO_DATE(PED.\"DataEmiss\", 'DD/MM/YYYY') >= %s::date")
        params.append(iso_inicio)

    iso_fim = _parse_to_iso_date(data_fim)
    if iso_fim:
        where_clauses.append("TO_DATE(PED.\"DataEmiss\", 'DD/MM/YYYY') <= %s::date")
        params.append(iso_fim)

    if id_empresa is not None and str(id_empresa).lower() not in ('all', ''):
        try:
            where_clauses.append('PED."id_empresa" = %s')
            params.append(int(id_empresa))
        except ValueError:
            pass

    where_sql = " WHERE " + " AND ".join(where_clauses)

    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            # 1. Dados cadastrais do Produto
            cursor.execute(f"""
                SELECT PRD.*, GRU."{DESC_GRU_COL}" as "Nome_Grupo"
                FROM "PRD" PRD
                LEFT JOIN "GRU" GRU ON PRD."Grupo" = GRU."CodGru"
                WHERE PRD."CodPrd" = %s
            """, (cod_prd,))
            prd_row = cursor.fetchone()
            prd_info = normalize_prd_row(prd_row) if prd_row else None

            # 2. Transações e pedidos individuais
            sql = f"""
                SELECT 
                    PED."CodPed",
                    PED."DataEmiss",
                    PED."Hora",
                    PED."CondPgto",
                    PED."Cfo",
                    PED."PrevEntrega",
                    PED."NroNt",
                    PED."DtFat",
                    NFE."NroChave",
                    NFE."Status" as "StatusNFe",
                    ENT."Nome" as "NomeCliente",
                    ENT."Cidade" as "CidadeCliente",
                    ITP."CodItp",
                    ITP."Qtd",
                    ITP."ValorUnit",
                    ITP."Desconto",
                    ITP."Valor" as "ValorTotal"
                FROM "ITP" ITP
                JOIN "PED" PED ON ITP."Pedido" = PED."CodPed"
                LEFT JOIN "NFE" NFE ON PED."CodPed" = NFE."CodPed"
                LEFT JOIN "ENT" ENT ON PED."Entidade" = ENT."CodEntidade"
                {where_sql}
                ORDER BY TO_DATE(PED."DataEmiss", 'DD/MM/YYYY') DESC, PED."CodPed" DESC
                LIMIT 200
            """
            cursor.execute(sql, params)
            vendas_raw = cursor.fetchall()
            
            vendas = []
            for r in vendas_raw:
                v = _convert_row(r)
                nro_nt = int(v.get("NroNt") or 0)
                nro_chave = v.get("NroChave") or ""
                prev_entrega = str(v.get("PrevEntrega") or "").strip().upper()
                cfo = str(v.get("Cfo") or "").strip().upper()
                
                is_cancelado = (prev_entrega == 'CANCELADO' or 'CANCEL' in cfo)
                has_nota = (nro_nt > 0 or bool(nro_chave))
                
                v["is_cancelado"] = is_cancelado
                v["has_nota"] = has_nota
                if is_cancelado:
                    v["status_label"] = "Cancelado"
                    v["status_badge_class"] = "badge-danger"
                elif has_nota:
                    v["status_label"] = f"Nota #{nro_nt}" if nro_nt > 0 else "NF-e Emitida"
                    v["status_badge_class"] = "badge-success"
                else:
                    v["status_label"] = "Pendente de Nota"
                    v["status_badge_class"] = "badge-warning"

                vendas.append(v)

            total_qtd = sum(float(v.get("Qtd", 0)) for v in vendas)
            total_valor = sum(float(v.get("ValorTotal", 0)) for v in vendas)
            total_desconto = sum(float(v.get("Desconto", 0)) for v in vendas)
            preco_medio = (total_valor / total_qtd) if total_qtd > 0 else 0.0

            return {
                "produto": prd_info,
                "vendas": vendas,
                "resumo_produto": {
                    "total_pedidos": len(vendas),
                    "total_qtd": total_qtd,
                    "total_desconto": round(total_desconto, 2),
                    "total_valor": round(total_valor, 2),
                    "preco_medio": round(preco_medio, 2)
                }
            }
    finally:
        conn.close()





