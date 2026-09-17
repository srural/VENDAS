import os
import datetime

def save_xml_to_disk(xml_content, chave_nfe, modelo="55"):
    """
    Salva os XMLs emitidos em pastas estruturadas por Tipo de Documento (NFE/NFCE) e Mês (YYYY-MM).
    Exemplo de estrutura gerada:
      xmls/NFE/2026-08/NFe35260823103347000165550010000000011000000014.xml
      xmls/NFCE/2026-08/NFCe35260823103347000165650010000000011000000018.xml
    """
    try:
        if not xml_content or not chave_nfe:
            return None

        is_nfce = str(modelo) == "65" or (len(str(chave_nfe)) >= 22 and str(chave_nfe)[20:22] == "65")
        doc_type = "NFCE" if is_nfce else "NFE"
        prefix = "NFCe" if is_nfce else "NFe"

        # Extrai ano e mês da chave de acesso (AAMM) ou utiliza a data atual
        clean_chave = str(chave_nfe).strip()
        aamm = clean_chave[2:6] if len(clean_chave) >= 6 and clean_chave[2:6].isdigit() else datetime.datetime.now().strftime("%y%m")
        ano_mes = f"20{aamm[:2]}-{aamm[2:4]}"

        project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        target_dir = os.path.join(project_dir, 'xmls', doc_type, ano_mes)
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
