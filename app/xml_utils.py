import os
import datetime
import shutil

def save_xml_to_disk(xml_content, chave_nfe, modelo="55", cnpj=None):
    """
    Salva os XMLs emitidos em pastas estruturadas por CNPJ da Empresa (Multiempresa),
    Tipo de Documento (NFE/NFCE) e Mês de Emissão (YYYY-MM).
    
    Exemplo de estrutura gerada:
      xmls/23103347000165/NFE/2026-09/NFe35260923103347000165550010000000101520565896.xml
      xmls/23103347000165/NFCE/2026-09/NFCe35260923103347000165650010000000541429701887.xml
    """
    try:
        if not xml_content or not chave_nfe:
            return None

        is_nfce = str(modelo) == "65" or (len(str(chave_nfe)) >= 22 and str(chave_nfe)[20:22] == "65")
        doc_type = "NFCE" if is_nfce else "NFE"
        prefix = "NFCe" if is_nfce else "NFe"

        clean_chave = ''.join(filter(str.isdigit, str(chave_nfe)))
        
        # 1. Extração do CNPJ da Empresa (Parâmetro explícito ou extraído da chave nacional de 44 dígitos)
        clean_cnpj = ''.join(filter(str.isdigit, str(cnpj or '')))
        if not clean_cnpj and len(clean_chave) >= 20:
            clean_cnpj = clean_chave[6:20]
        if not clean_cnpj:
            clean_cnpj = "GERAL"

        # 2. Extrai ano e mês da chave de acesso (AAMM) ou utiliza a data atual
        aamm = clean_chave[2:6] if len(clean_chave) >= 6 else datetime.datetime.now().strftime("%y%m")
        ano_mes = f"20{aamm[:2]}-{aamm[2:4]}"

        # 3. Montagem do caminho do diretório com nível de CNPJ
        project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        target_dir = os.path.join(project_dir, 'xmls', clean_cnpj, doc_type, ano_mes)
        os.makedirs(target_dir, exist_ok=True)

        filename = f"{prefix}{clean_chave}.xml"
        filepath = os.path.join(target_dir, filename)

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(xml_content)

        print(f"[XML SAVED] {doc_type} gravado em '{filepath}'")
        return filepath
    except Exception as e:
        print(f"[XML SAVE ERROR] Erro ao gravar arquivo XML no disco: {e}")
        return None

def migrate_existing_xml_folders():
    """
    Organiza retroativamente XMLs que estejam no caminho antigo (xmls/NFE/... ou xmls/NFCE/...)
    para a nova estrutura com nível de CNPJ (xmls/<CNPJ>/<TIPO>/<ANO-MES>/...).
    """
    project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    xmls_root = os.path.join(project_dir, 'xmls')
    if not os.path.exists(xmls_root):
        return

    for doc_type in ['NFE', 'NFCE']:
        old_type_dir = os.path.join(xmls_root, doc_type)
        if os.path.isdir(old_type_dir):
            for root, _, files in os.walk(old_type_dir):
                for file in files:
                    if file.endswith('.xml'):
                        src_path = os.path.join(root, file)
                        # Extrai a chave a partir do nome do arquivo
                        digits = ''.join(filter(str.isdigit, file))
                        if len(digits) >= 20:
                            cnpj = digits[6:20]
                            aamm = digits[2:6]
                            ano_mes = f"20{aamm[:2]}-{aamm[2:4]}"
                            dest_dir = os.path.join(xmls_root, cnpj, doc_type, ano_mes)
                            os.makedirs(dest_dir, exist_ok=True)
                            dest_path = os.path.join(dest_dir, file)
                            try:
                                shutil.move(src_path, dest_path)
                                print(f"[XML MIGRATED] Movido: {src_path} -> {dest_path}")
                            except Exception as ex:
                                print(f"[XML MIGRATE NOTICE] Não foi possível mover {file}: {ex}")
            # Limpa pastas antigas vazias
            try:
                shutil.rmtree(old_type_dir)
            except Exception:
                pass
