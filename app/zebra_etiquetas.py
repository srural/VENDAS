"""
Módulo Emissor de Etiquetas de Gôndola de Supermercado 9x3 (90mm x 30mm) para Impressoras Zebra
Compatível com ZPL II (Zebra Programming Language)
Suporta resoluções 203 DPI (8 dots/mm) e 300 DPI (12 dots/mm)
Modelos:
- Padrão Gôndola Varejo
- Oferta / Promocional (De / Por)
- Preço por Unidade de Medida (Kg / Litro / Unidade)
- Compacto
"""

import re
import socket
import unicodedata
import subprocess
import os
import datetime
from app.empresa_manager import get_empresa_by_id, get_empresas

# Dimensões 90x30 mm
# 203 DPI: 90mm * 8 = 720 dots | 30mm * 8 = 240 dots (Margem de segurança: 712 x 236)
# 300 DPI: 90mm * 12 = 1080 dots | 30mm * 12 = 360 dots (Margem de segurança: 1060 x 350)
DPI_CONFIG = {
    203: {
        "width": 712,
        "height": 236,
        "dpmm": 8,
        "font_header": ("0", 20, 18),
        "font_date": ("0", 18, 16),
        "font_desc": ("0", 28, 24),
        "font_cod": ("0", 18, 16),
        "font_currency": ("0", 22, 18),
        "font_price": ("0", 72, 62),
        "font_price_cents": ("0", 42, 36),
        "font_promo_old": ("0", 19, 17),
        "font_promo_new": ("0", 62, 54),
        "barcode_height": 52
    },
    300: {
        "width": 1060,
        "height": 350,
        "dpmm": 12,
        "font_header": ("0", 30, 26),
        "font_date": ("0", 26, 22),
        "font_desc": ("0", 40, 36),
        "font_cod": ("0", 26, 22),
        "font_currency": ("0", 32, 26),
        "font_price": ("0", 105, 90),
        "font_price_cents": ("0", 60, 50),
        "font_promo_old": ("0", 28, 24),
        "font_promo_new": ("0", 90, 78),
        "barcode_height": 78
    }
}

def remove_accents(text):
    """Remove acentuação para máxima compatibilidade com modelos antigos de Zebra"""
    if not text:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(text))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper()

def format_currency_br(val):
    """Formata valor float para moeda brasileira: 12.5 -> '12,50'"""
    try:
        f = float(val or 0.0)
        return f"{f:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return "0,00"

def calculate_price_per_unit(preco, peso_liquido=0, embalagem='UN'):
    """Calcula preço proporcional (ex: Preço por Kg ou Preço por Litro) para conformidade legal"""
    try:
        p = float(preco or 0.0)
        peso = float(peso_liquido or 0.0)
        emb = str(embalagem or 'UN').strip().upper()
        
        if peso > 0 and emb in ('KG', 'G', 'GR'):
            # se peso está em gramas (> 10 e <= 1000)
            if emb in ('G', 'GR') or peso > 50:
                kg = peso / 1000.0
            else:
                kg = peso
            if kg > 0:
                preco_kg = p / kg
                return f"PRECO P/ KG: R$ {format_currency_br(preco_kg)}"
        elif peso > 0 and emb in ('L', 'LT', 'ML'):
            if emb == 'ML' or peso > 50:
                litros = peso / 1000.0
            else:
                litros = peso
            if litros > 0:
                preco_lt = p / litros
                return f"PRECO P/ LT: R$ {format_currency_br(preco_lt)}"
        
        return f"PRECO P/ {emb}: R$ {format_currency_br(p)}"
    except Exception:
        return ""

def generate_barcode_zpl(codbar, cod_prd, dpi=203, x=20, y=125):
    """Gera o bloco ZPL adequado para o código de barras (EAN-13, EAN-8 ou Code 128)"""
    cfg = DPI_CONFIG.get(dpi, DPI_CONFIG[203])
    bc_height = cfg["barcode_height"]
    
    clean_bar = re.sub(r'[^0-9A-Za-z]', '', str(codbar or '')).strip()
    
    # Se não tem código de barras ou é inválido, usa o código do produto como Code 128
    if not clean_bar:
        clean_bar = str(cod_prd)
        # Code 128
        return f"^FO{x},{y}^BCN,{bc_height},Y,N,N^FD{clean_bar}^FS"
    
    # Se é puramente numérico e tem 13 dígitos -> EAN-13
    if clean_bar.isdigit() and len(clean_bar) == 13:
        return f"^FO{x},{y}^BEN,{bc_height},Y,N^FD{clean_bar}^FS"
    
    # Se é numérico e tem 8 dígitos -> EAN-8
    if clean_bar.isdigit() and len(clean_bar) == 8:
        return f"^FO{x},{y}^B8N,{bc_height},Y,N^FD{clean_bar}^FS"
    
    # Caso geral -> Code 128
    return f"^FO{x},{y}^BCN,{bc_height},Y,N,N^FD{clean_bar}^FS"

def get_empresa_header_name():
    """Obtém o nome fantasia ou razão social da empresa para o cabeçalho da gôndola"""
    try:
        emp = get_empresa_by_id(1)
        if not emp:
            empresas = get_empresas()
            if empresas:
                emp = empresas[0]
        if emp:
            return (emp.get('Fantasia') or emp.get('RazaoSocial') or 'SUPERMERCADO').strip()
    except Exception:
        pass
    return "SUPERMERCADO"

def build_single_label_zpl(item, options=None):
    """
    Gera o código ZPL para uma etiqueta individual de 90x30 mm.
    item: dict contendo CodPrd, Descricao_Produto, Venda, CodBar, etc.
    options: dict com configurações opcionais (modelo, dpi, header_text, etc.)
    """
    opts = options or {}
    dpi = int(opts.get('dpi', 203))
    if dpi not in DPI_CONFIG:
        dpi = 203
        
    cfg = DPI_CONFIG[dpi]
    width = cfg["width"]
    height = cfg["height"]
    
    modelo = opts.get('modelo', 'padrao')  # padrao, promocional, por_unidade, compacto
    header_custom = opts.get('header_text') or get_empresa_header_name()
    show_header = opts.get('show_header', True)
    show_date = opts.get('show_date', True)
    show_cod = opts.get('show_cod', True)
    show_unit = opts.get('show_unit', True)
    show_unit_price = opts.get('show_unit_price', True)
    copies = int(item.get('copies', 1))
    if copies < 1:
        copies = 1

    # Dados do item
    cod_prd = str(item.get('CodPrd', '') or item.get('Codigo', ''))
    descricao = remove_accents(item.get('Descricao_Produto', 'PRODUTO SEM DESCRICAO')).strip()
    embalagem = str(item.get('Embalagem', 'UN') or 'UN').strip().upper()
    marca = remove_accents(item.get('Marca', '')).strip()
    codbar = str(item.get('CodBar', '') or '').strip()
    local = remove_accents(item.get('Local', '')).strip()
    
    # Preços
    preco_venda = float(item.get('Venda', 0.0) or 0.0)
    preco_promo = float(item.get('Venda_Promocao', 0.0) or 0.0)
    is_promo = bool(item.get('Promocao') or (preco_promo > 0 and preco_promo < preco_venda) or modelo == 'promocional')
    
    # Se o usuário digitou um preço customizado para a etiqueta
    if 'PrecoEtiqueta' in item and item['PrecoEtiqueta'] is not None:
        try:
            preco_venda = float(item['PrecoEtiqueta'])
        except Exception:
            pass

    preco_formatado = format_currency_br(preco_venda)
    
    # Data de emissão
    dt_emissao = opts.get('custom_date') or datetime.datetime.now().strftime("%d/%m/%Y")
    
    # Preço por unidade
    peso_liq = float(item.get('PesoLiquido', 0.0) or 0.0)
    texto_preco_unidade = calculate_price_per_unit(preco_venda, peso_liq, embalagem) if show_unit_price else ""

    zpl_commands = []
    
    # Inicialização da Etiqueta
    zpl_commands.append("^XA")
    zpl_commands.append("^MMT")            # Tear-off mode
    zpl_commands.append(f"^PW{width}")     # Print Width
    zpl_commands.append(f"^LL{height}")    # Label Length
    zpl_commands.append("^LS0")            # Label Shift
    zpl_commands.append("^PON")            # Normal Orientation
    zpl_commands.append("^CI28")           # UTF-8 Encoding

    scale = 1.0 if dpi == 203 else (300.0 / 203.0)
    
    def sc(val):
        return int(round(val * scale))

    # --- CABEÇALHO (SUPERIOR) ---
    if show_header:
        header_text = remove_accents(header_custom)[:30]
        f_h = cfg["font_header"]
        zpl_commands.append(f"^FO{sc(20)},{sc(10)}^A{f_h[0]}N,{f_h[1]},{f_h[2]}^FD{header_text}^FS")
    
    if show_date:
        f_d = cfg["font_date"]
        date_str = f"EMIS: {dt_emissao}"
        zpl_commands.append(f"^FO{sc(510)},{sc(12)}^A{f_d[0]}N,{f_d[1]},{f_d[2]}^FD{date_str}^FS")

    # Linha divisória fina abaixo do cabeçalho
    zpl_commands.append(f"^FO{sc(20)},{sc(32)}^GB{sc(672)},{sc(2)},{sc(2)}^FS")

    # --- DESCRIÇÃO DO PRODUTO (DESTAQUE) ---
    f_desc = cfg["font_desc"]
    # Limita descrição para caber perfeitamente no bloco (2 linhas max, 38 chars por linha)
    desc_curta = descricao[:76]
    zpl_commands.append(f"^FO{sc(20)},{sc(38)}^FB{sc(672)},2,0,L,0^A{f_desc[0]}N,{f_desc[1]},{f_desc[2]}^FD{desc_curta}^FS")

    # --- INFORMAÇÕES SECUNDÁRIAS (CÓDIGO, UNIDADE, MARCA, LOCAL) ---
    info_parts = []
    if show_cod and cod_prd:
        info_parts.append(f"COD: {cod_prd}")
    if show_unit and embalagem:
        info_parts.append(f"UN: {embalagem}")
    if marca:
        info_parts.append(f"{marca[:12]}")
    if local:
        info_parts.append(f"LOC: {local[:8]}")
        
    info_str = " | ".join(info_parts)
    f_cod = cfg["font_cod"]
    zpl_commands.append(f"^FO{sc(20)},{sc(98)}^A{f_cod[0]}N,{f_cod[1]},{f_cod[2]}^FD{info_str}^FS")

    # --- CÓDIGO DE BARRAS (ÁREA INFERIOR ESQUERDA) ---
    bc_zpl = generate_barcode_zpl(codbar, cod_prd, dpi=dpi, x=sc(20), y=sc(122))
    zpl_commands.append(bc_zpl)

    # --- PREÇO DESTAQUE (ÁREA INFERIOR DIREITA) ---
    if is_promo and (preco_promo > 0 and preco_promo < preco_venda):
        # MODO OFERTA / PROMOCIONAL (DE / POR)
        # Caixa "OFERTA" com tarja preta e texto reverso branco
        zpl_commands.append(f"^FO{sc(420)},{sc(95)}^GB{sc(272)},{sc(26)},{sc(26)}^FS")
        zpl_commands.append(f"^FO{sc(435)},{sc(99)}^FR^A0N,{sc(18)},{sc(16)}^FD*** OFERTA ESPECIAL ***^FS")
        
        # Preço DE
        f_pold = cfg["font_promo_old"]
        preco_de_txt = f"DE R$ {preco_formatado}"
        zpl_commands.append(f"^FO{sc(425)},{sc(126)}^A{f_pold[0]}N,{f_pold[1]},{f_pold[2]}^FD{preco_de_txt}^FS")
        # Linha riscando o "DE"
        zpl_commands.append(f"^FO{sc(425)},{sc(134)}^GB{sc(180)},{sc(2)},{sc(2)}^FS")

        # Preço POR em destaque
        f_pnew = cfg["font_promo_new"]
        preco_por_fmt = format_currency_br(preco_promo)
        zpl_commands.append(f"^FO{sc(420)},{sc(146)}^A0N,{sc(20)},{sc(18)}^FDPOR R$^FS")
        zpl_commands.append(f"^FO{sc(465)},{sc(142)}^A{f_pnew[0]}N,{f_pnew[1]},{f_pnew[2]}^FD{preco_por_fmt}^FS")
        
        if show_unit_price and texto_preco_unidade:
            zpl_commands.append(f"^FO{sc(420)},{sc(208)}^A0N,{sc(15)},{sc(13)}^FD{texto_preco_unidade[:30]}^FS")

    else:
        # MODO PADRÃO DE GÔNDOLA (PREÇO GIGANTE E LIMPO)
        # Moldura suave ao redor do preço
        zpl_commands.append(f"^FO{sc(425)},{sc(96)}^GB{sc(267)},{sc(130)},{sc(2)}^FS")
        
        # "R$" no canto superior da caixa de preço
        f_curr = cfg["font_currency"]
        zpl_commands.append(f"^FO{sc(435)},{sc(105)}^A{f_curr[0]}N,{f_curr[1]},{f_curr[2]}^FDR$^FS")

        # Preço de Venda em tamanho gigante
        f_p = cfg["font_price"]
        zpl_commands.append(f"^FO{sc(468)},{sc(106)}^A{f_p[0]}N,{f_p[1]},{f_p[2]}^FD{preco_formatado}^FS")

        # Preço por Unidade de Medida / Preço à Vista no rodapé da caixa
        if show_unit_price and texto_preco_unidade:
            zpl_commands.append(f"^FO{sc(432)},{sc(196)}^A0N,{sc(15)},{sc(14)}^FD{texto_preco_unidade[:32]}^FS")
        else:
            zpl_commands.append(f"^FO{sc(450)},{sc(196)}^A0N,{sc(16)},{sc(14)}^FDAVISTA / UNIDADE^FS")

    # Quantidade de cópias
    if copies > 1:
        zpl_commands.append(f"^PQ{copies}")

    zpl_commands.append("^XZ\n")
    return "\n".join(zpl_commands)

def generate_batch_labels_zpl(items, options=None):
    """
    Gera o pacote ZPL completo para uma lista de produtos.
    """
    if not items:
        return ""
    
    labels_zpl = []
    for item in items:
        labels_zpl.append(build_single_label_zpl(item, options))
        
    return "\n".join(labels_zpl)

def send_zpl_to_network_printer(ip_address, zpl_content, port=9100, timeout=5):
    """
    Envia o código ZPL diretamente para a porta TCP/IP (padrão 9100) da impressora Zebra.
    Retorna (sucesso: bool, mensagem: str)
    """
    if not ip_address:
        return False, "Endereço IP da impressora não informado."
    
    clean_ip = ip_address.strip()
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        sock.connect((clean_ip, int(port)))
        sock.sendall(zpl_content.encode('utf-8', errors='replace'))
        sock.close()
        return True, f"Etiquetas enviadas com sucesso para a impressora Zebra em {clean_ip}:{port}!"
    except socket.timeout:
        return False, f"Tempo de conexão esgotado ao tentar comunicar com a impressora em {clean_ip}:{port}."
    except ConnectionRefusedError:
        return False, f"Conexão recusada pela impressora em {clean_ip}:{port}. Verifique se a porta 9100 está aberta."
    except Exception as e:
        return False, f"Falha ao comunicar com a impressora ({clean_ip}:{port}): {str(e)}"

def send_zpl_to_windows_printer(printer_name, zpl_content):
    """
    Envia dados brutos ZPL diretamente para uma impressora configurada no Windows.
    Utiliza PowerShell para enviar os bytes brutos para o spooler.
    """
    if not printer_name:
        return False, "Nome da impressora do Windows não informado."

    temp_dir = os.path.join(os.path.dirname(__file__), '..', 'scratch')
    os.makedirs(temp_dir, exist_ok=True)
    temp_file = os.path.join(temp_dir, f"etiquetas_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.zpl")
    
    try:
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(zpl_content)

        # Envia para a impressora via PowerShell
        # Se a impressora é compartilhada localmente: copy /b temp_file "\\localhost\printer_name"
        # Ou usando Out-Printer ou RawPrint
        cmd = f'Get-Content -Path "{temp_file}" -Raw | Out-Printer -Name "{printer_name}"'
        res = subprocess.run(['powershell', '-Command', cmd], capture_output=True, text=True, timeout=8)
        
        if res.returncode == 0:
            return True, f"Etiquetas enviadas para a impressora '{printer_name}' com sucesso!"
        else:
            # Fallback usando copy /b para UNC
            unc_path = f"\\\\127.0.0.1\\{printer_name}"
            copy_res = subprocess.run(f'copy /b "{temp_file}" "{unc_path}"', shell=True, capture_output=True, text=True)
            if copy_res.returncode == 0:
                return True, f"Etiquetas enviadas para a impressora '{printer_name}' via spooler!"
            return False, f"Erro ao enviar para impressora Windows: {res.stderr or copy_res.stderr or 'Falha no spooler'}"
    except Exception as e:
        return False, f"Exceção ao imprimir no Windows: {str(e)}"
    finally:
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except Exception:
                pass

def get_installed_printers():
    """Retorna lista de impressoras instaladas no Windows"""
    try:
        cmd = 'Get-CimInstance Win32_Printer | Select-Object -ExpandProperty Name'
        res = subprocess.run(['powershell', '-Command', cmd], capture_output=True, text=True, timeout=5)
        if res.returncode == 0:
            printers = [p.strip() for p in res.stdout.strip().split('\n') if p.strip()]
            return printers
    except Exception:
        pass
    return []
