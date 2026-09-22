import datetime
import html as pyhtml
import os
import base64
import mimetypes

def format_cnpj_cpf(val):
    if not val:
        return ""
    digits = ''.join(filter(str.isdigit, str(val)))
    if len(digits) == 14:
        return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"
    elif len(digits) == 11:
        return f"{digits[0:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:11]}"
    return str(val)

def format_cep(val):
    if not val:
        return ""
    digits = ''.join(filter(str.isdigit, str(val)))
    if len(digits) == 8:
        return f"{digits[0:5]}-{digits[5:8]}"
    return str(val)

def format_money(val):
    try:
        f = float(val or 0.0)
        return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"

def format_qty(val):
    try:
        f = float(val or 0.0)
        return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"

def format_nfe_number(val):
    try:
        n = int(''.join(filter(str.isdigit, str(val or 1))))
        s = f"{n:09d}"
        return f"{s[0:3]}.{s[3:6]}.{s[6:9]}"
    except:
        return str(val).zfill(9)

def format_serie(val):
    try:
        n = int(''.join(filter(str.isdigit, str(val or 1))))
        return f"{n:03d}"
    except:
        return str(val).zfill(3)

def format_chave_44(chave):
    digits = ''.join(filter(str.isdigit, str(chave or '')))
    if len(digits) == 44:
        return ' '.join([digits[i:i+4] for i in range(0, 44, 4)])
    return str(chave or '')

def generate_code128c_svg(digits):
    """
    Gera código de barras Code 128C para os 44 dígitos da Chave de Acesso da NF-e
    em SVG puro com alta fidelidade de impressão.
    """
    code128_patterns = [
        '212222', '222122', '222221', '121223', '121322', '131222', '122213', '122312', '132212', '221213',
        '221312', '231212', '112232', '122132', '122231', '113222', '123122', '123221', '223211', '221132',
        '221231', '213212', '223112', '312131', '311222', '321122', '321221', '312212', '322112', '322211',
        '212123', '212321', '232121', '111323', '131123', '131321', '112313', '132113', '132311', '211313',
        '231113', '231311', '112133', '112331', '132131', '113123', '113321', '133121', '313121', '211331',
        '231131', '213113', '213311', '213131', '311123', '311321', '331121', '312113', '312311', '332111',
        '314111', '221411', '431111', '111224', '111422', '121124', '121421', '141122', '141221', '112214',
        '112412', '122114', '122411', '142112', '142211', '241211', '221114', '413111', '241112', '134111',
        '111242', '121142', '121241', '114212', '124112', '124211', '411212', '421112', '421211', '212141',
        '214121', '412121', '111143', '111341', '131141', '114113', '114311', '411113', '411311', '113141',
        '114131', '311141', '411131', '211412', '211214', '211232', '2331112'
    ]
    digits = ''.join(filter(str.isdigit, str(digits)))
    if not digits:
        return ''
    if len(digits) % 2 != 0:
        digits = '0' + digits
    values = [105] # Start C
    for i in range(0, len(digits), 2):
        values.append(int(digits[i:i+2]))
    checksum = values[0]
    for idx, v in enumerate(values[1:], 1):
        checksum += idx * v
    values.append(checksum % 103)
    values.append(106) # Stop
    
    modules = ''
    for v in values:
        pat = code128_patterns[v]
        is_bar = True
        for count in pat:
            modules += ('1' if is_bar else '0') * int(count)
            is_bar = not is_bar
            
    svg_bars = []
    x = 0
    w = 1
    h = 42
    for b in modules:
        if b == '1':
            svg_bars.append(f'<rect x="{x}" y="0" width="{w}" height="{h}" fill="black" />')
        x += w
    return f'<svg viewBox="0 0 {x} {h}" preserveAspectRatio="none" style="width: 95%; height: 38px; display: block; margin: 2px auto 3px auto;" xmlns="http://www.w3.org/2000/svg">' + ''.join(svg_bars) + '</svg>'

def render_danfe_html(order, company, items, nfe_res):
    """
    Renderiza o layout completo e oficial do DANFE (Documento Auxiliar da Nota Fiscal Eletrônica - Modelo 55)
    em formato retrato (A4), fiel ao padrão nacional SEFAZ.
    """
    nro_nota_fmt = format_nfe_number(order.get('NroNfe', order.get('CodPed', 1)))
    serie_fmt = format_serie(order.get('SerieNfe', '1'))
    chave_nfe = nfe_res.get('chave_nfe', '')
    chave_formatada = format_chave_44(chave_nfe)
    barcode_svg = generate_code128c_svg(chave_nfe) if chave_nfe else ''
    
    dt_emiss = order.get('DataEmiss') or order.get('Emissao') or datetime.date.today().strftime('%d/%m/%Y')
    if isinstance(dt_emiss, datetime.datetime) or isinstance(dt_emiss, datetime.date):
        dt_emiss = dt_emiss.strftime('%d/%m/%Y')
    elif 'T' in str(dt_emiss):
        dt_emiss = str(dt_emiss).split('T')[0]
        dt_parts = dt_emiss.split('-')
        if len(dt_parts) == 3:
            dt_emiss = f"{dt_parts[2]}/{dt_parts[1]}/{dt_parts[0]}"
            
    dt_saida = order.get('DtSaida') or order.get('Saida') or dt_emiss
    if isinstance(dt_saida, datetime.datetime) or isinstance(dt_saida, datetime.date):
        dt_saida = dt_saida.strftime('%d/%m/%Y')
    elif 'T' in str(dt_saida):
        dt_saida = str(dt_saida).split('T')[0]
        dt_parts = dt_saida.split('-')
        if len(dt_parts) == 3:
            dt_saida = f"{dt_parts[2]}/{dt_parts[1]}/{dt_parts[0]}"

    hora_saida = order.get('HoraSaida') or order.get('Hora') or ''
    
    # Totals
    total_prod = sum(float(it.get('Qtd') or 1.0) * float(it.get('ValorUnit') or 0.0) for it in items)
    desconto = float(order.get('Desconto') or 0.0)
    frete = float(order.get('ValorFrete') or 0.0)
    seguro = float(order.get('ValorSeguro') or 0.0)
    outras_desp = float(order.get('OutrasDesp') or 0.0)
    base_icms = float(order.get('BaseIcms') or 0.0)
    valor_icms = float(order.get('ValorIcms') or 0.0)
    base_st = float(order.get('BaseSt') or order.get('PedBaseSub') or 0.0)
    valor_st = float(order.get('ValorSt') or order.get('PedSub') or 0.0)
    valor_ipi = float(order.get('ValorIpi') or order.get('PedIpi') or 0.0)
    total_nota = float(order.get('Total') or (total_prod - desconto + frete + seguro + outras_desp + valor_st + valor_ipi))

    # Natureza de operação
    nat_op = str(order.get('Cfo') or 'VENDA DE MERCADORIAS').upper()
    if ' - ' in nat_op:
        nat_op_desc = nat_op.split(' - ', 1)[1]
    else:
        nat_op_desc = nat_op

    # Protocolo e Ambiente Fiscal
    from app.config_manager import get_pdv_config
    cfg = get_pdv_config()
    ambiente = str(order.get('Ambiente') or nfe_res.get('ambiente') or cfg.get('Nfe', {}).get('Ambiente', '2'))
    status_nfe = str(nfe_res.get('status') or order.get('StatusNFe') or order.get('Status') or '').upper()
    protocolo = nfe_res.get('protocolo') or ''

    is_not_authorized = ('AUTORIZ' not in status_nfe and not protocolo)
    is_homologacao = (ambiente != '1')
    has_no_fiscal_validity = is_homologacao or is_not_authorized

    if is_homologacao:
        proto_data = f"HOMOLOGAÇÃO - SEM VALIDADE FISCAL ({protocolo or 'SIMULAÇÃO'}) {dt_emiss}"
        strip_text = "EMITIDA EM AMBIENTE DE HOMOLOGAÇÃO - SEM VALOR FISCAL"
        sub_watermark = "EMITIDA EM AMBIENTE DE HOMOLOGAÇÃO"
    elif is_not_authorized:
        proto_data = f"NÃO AUTORIZADA NA SEFAZ ({protocolo or 'PENDENTE'}) {dt_emiss}"
        strip_text = "DOCUMENTO NÃO VALIDADO NA SEFAZ - SEM VALOR FISCAL"
        sub_watermark = "NÃO VALIDADA NA SEFAZ"
    else:
        protocolo_val = protocolo or '135180587979473'
        proto_data = f"{protocolo_val} {dt_emiss} {hora_saida or '18:22:40'}"
        strip_text = ""
        sub_watermark = ""

    # Frete por conta
    mod_frete = str(order.get('ModFrete') or order.get('TipoFrete') or '0')
    frete_map = {
        '0': '0 - REMETENTE',
        '1': '1 - DESTINATÁRIO',
        '2': '2 - TERCEIROS',
        '3': '3 - PRÓPRIO REMET',
        '4': '4 - PRÓPRIO DEST',
        '9': '9 - SEM TRANSPORTE'
    }
    frete_txt = frete_map.get(mod_frete, f'{mod_frete} - REMETENTE')

    # Tipo E/S (1 = Saída, 0 = Entrada)
    tp_nf = "1"
    
    # Destinatário
    dest_nome = order.get('NomeCliente') or order.get('Nome') or 'CONSUMIDOR FINAL'
    dest_doc = format_cnpj_cpf(order.get('CPF') or order.get('CGC') or '')
    dest_end = f"{order.get('Endereco') or ''}{', ' + str(order.get('Nro')) if order.get('Nro') else ''}"
    dest_bairro = order.get('Bairro') or ''
    dest_cep = format_cep(order.get('Cep') or order.get('CEP') or '')
    dest_mun = order.get('Cidade') or ''
    dest_uf = order.get('Uf') or order.get('UF') or ''
    dest_fone = order.get('Fone') or ''
    dest_ie = order.get('InscEst') or order.get('InscrEst') or order.get('IE') or ''

    # Emitente
    emit_nome = company.get('RazaoSocial') or company.get('NomeEmpresa') or 'SIDIVAL CARLOS CIOCA ME'
    emit_cnpj = format_cnpj_cpf(company.get('CNPJ') or '23.103.347/0001-65')
    emit_ie = company.get('InscEst') or company.get('IE') or '707021792115'
    emit_end = f"{company.get('Logradouro') or company.get('Endereco') or 'RUA PRINCIPAL'}, {company.get('Nro') or '100'}"
    emit_bairro = company.get('Bairro') or 'CENTRO'
    emit_cep = format_cep(company.get('CEP') or '15850-000')
    emit_mun = company.get('Cidade') or 'URUPES'
    emit_uf = company.get('UF') or 'SP'
    emit_fone = company.get('Telefone') or company.get('Fone') or ''
    emit_logo = company.get('Logo') or company.get('logo') or ''

    # Process logo (convert local static upload to base64 data URI if available)
    logo_src = emit_logo
    if emit_logo and str(emit_logo).startswith('/static/'):
        local_path = os.path.join(os.path.dirname(__file__), emit_logo.lstrip('/'))
        if os.path.exists(local_path):
            try:
                mime_type, _ = mimetypes.guess_type(local_path)
                mime_type = mime_type or 'image/png'
                with open(local_path, 'rb') as img_f:
                    b64_data = base64.b64encode(img_f.read()).decode('utf-8')
                    logo_src = f"data:{mime_type};base64,{b64_data}"
            except Exception:
                logo_src = emit_logo

    # Items HTML
    items_rows_html = []
    for idx, it in enumerate(items, 1):
        c_prod = it.get('CodPrd') or it.get('Produto') or idx
        x_prod = it.get('Descricao_Produto') or it.get('Descricao') or f'PRODUTO {c_prod}'
        ncm = str(it.get('NCM') or it.get('ClasseFiscal') or '85176241').replace('.', '')[:8]
        cst = str(it.get('CST') or it.get('CSOSN') or it.get('SitTrib') or '0500')
        cfop = str(it.get('CFOP') or it.get('CfOpPrd') or '5405').replace('.', '')[:4]
        unid = it.get('Embalagem') or 'UN'
        qtd = float(it.get('Qtd') or 1.0)
        v_un = float(it.get('ValorUnit') or 0.0)
        v_tot = qtd * v_un
        v_desc_it = float(it.get('Desconto') or 0.0)
        v_bc_it = float(it.get('BaseIcms') or 0.0)
        v_icms_it = float(it.get('ValorIcms') or 0.0)
        v_ipi_it = float(it.get('ValorIpi') or 0.0)
        aliq_icms = float(it.get('AliqIcms') or 0.0)
        aliq_ipi = float(it.get('AliqIpi') or 0.0)

        items_rows_html.append(f"""
        <tr>
            <td class="text-left">{pyhtml.escape(str(c_prod))}</td>
            <td class="text-left font-bold">{pyhtml.escape(str(x_prod))}</td>
            <td class="text-center">{pyhtml.escape(str(ncm))}</td>
            <td class="text-center">{pyhtml.escape(str(cst))}</td>
            <td class="text-center">{pyhtml.escape(str(cfop))}</td>
            <td class="text-center">{pyhtml.escape(str(unid))}</td>
            <td class="text-right">{format_qty(qtd)}</td>
            <td class="text-right">{format_money(v_un)}</td>
            <td class="text-right">{format_money(v_tot)}</td>
            <td class="text-right">{format_money(v_desc_it)}</td>
            <td class="text-right">{format_money(v_bc_it)}</td>
            <td class="text-right">{format_money(v_icms_it)}</td>
            <td class="text-right">{format_money(v_ipi_it)}</td>
            <td class="text-right">{format_money(aliq_icms)}</td>
            <td class="text-right">{format_money(aliq_ipi)}</td>
        </tr>
        """)

    # Observações e Tributos Aproximados (Lei da Transparência 12.741)
    obs_texto = order.get('ObsCorpo') or order.get('ObsRodape') or order.get('Obs') or ''
    trib_federal = total_nota * 0.0700
    trib_estadual = total_nota * 0.1242
    trib_total = trib_federal + trib_estadual
    inf_compl = f"Trib.RS {format_money(trib_total)}(19,42%) Federal RS {format_money(trib_federal)}(7,00%) Estadual.<br>{pyhtml.escape(str(obs_texto))}"

    homolog_strip_html = f"""
    <div class="danfe-homolog-strip">
        {strip_text}
    </div>
    """ if has_no_fiscal_validity else ""

    watermark_html = f"""
    <div class="danfe-watermark">
        <div class="watermark-main">SEM VALOR FISCAL</div>
        <div class="watermark-sub">{sub_watermark}</div>
    </div>
    """ if has_no_fiscal_validity else ""

    html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <title>DANFE - NF-e Nº {nro_nota_fmt} Série {serie_fmt}</title>
    <style>
        * {{
            box-sizing: border-box;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
            color-adjust: exact !important;
            margin: 0;
            padding: 0;
        }}
        body {{
            font-family: Arial, Helvetica, sans-serif;
            background-color: #f0f2f5;
            color: #000;
            font-size: 8px;
            line-height: 1.15;
            padding: 20px 10px;
        }}
        .danfe-page {{
            width: 204mm;
            min-height: 288mm;
            margin: 0 auto;
            background: #fff;
            padding: 4mm 5mm;
            border: 1px solid #ccc;
            box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            position: relative;
            overflow: hidden;
        }}
        @media print {{
            @page {{
                size: A4 portrait;
                margin: 4mm 5mm;
            }}
            body {{
                background: #fff;
                padding: 0;
            }}
            .danfe-page {{
                width: 100%;
                min-height: auto;
                border: none;
                box-shadow: none;
                padding: 0;
                margin: 0;
            }}
            .no-print {{
                display: none !important;
            }}
            .danfe-watermark {{
                display: block !important;
            }}
        }}

        /* Watermark & Homologation Banner */
        .danfe-homolog-strip {{
            background-color: #fee2e2 !important;
            border: 2px dashed #dc2626 !important;
            color: #991b1b !important;
            padding: 4px 6px;
            text-align: center;
            font-weight: 900;
            font-size: 10.5px;
            margin-bottom: 3px;
            letter-spacing: 0.8px;
            text-transform: uppercase;
        }}
        .danfe-watermark {{
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%) rotate(-32deg);
            width: 90%;
            border: 6px dashed rgba(220, 38, 38, 0.22);
            background: rgba(254, 226, 226, 0.12);
            padding: 25px 15px;
            text-align: center;
            pointer-events: none;
            z-index: 999;
        }}
        .watermark-main {{
            font-size: 38px;
            font-weight: 900;
            color: rgba(220, 38, 38, 0.22);
            letter-spacing: 3px;
            line-height: 1.1;
            text-transform: uppercase;
        }}
        .watermark-sub {{
            font-size: 14px;
            font-weight: 800;
            color: rgba(220, 38, 38, 0.25);
            letter-spacing: 1.5px;
            margin-top: 6px;
            text-transform: uppercase;
        }}

        /* Typography & Labels */
        .lbl {{
            font-size: 5.8px;
            font-weight: bold;
            color: #000;
            text-transform: uppercase;
            display: block;
            margin-bottom: 1px;
            line-height: 1;
        }}
        .val {{
            font-size: 8px;
            font-weight: normal;
            color: #000;
            word-break: break-word;
        }}
        .val-bold {{
            font-size: 8.5px;
            font-weight: bold;
            color: #000;
        }}
        .sec-title {{
            font-size: 7px;
            font-weight: bold;
            text-transform: uppercase;
            margin-top: 3px;
            margin-bottom: 1px;
            color: #000;
        }}
        .text-center {{ text-align: center; }}
        .text-right {{ text-align: right; }}
        .text-left {{ text-align: left; }}
        .font-bold {{ font-weight: bold; }}

        /* Canhoto de Recebimento */
        .canhoto-container {{
            border: 1px solid #000;
            margin-bottom: 2px;
        }}
        .canhoto-header {{
            font-size: 6px;
            font-weight: bold;
            padding: 2px 4px;
            border-bottom: 1px solid #000;
            line-height: 1.1;
        }}
        .canhoto-body {{
            display: flex;
            height: 24px;
        }}
        .canhoto-box-right {{
            width: 155px;
            border-left: 1px solid #000;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            padding: 2px;
        }}
        .cut-line {{
            border-bottom: 1px dashed #000;
            margin: 3px 0 4px 0;
            position: relative;
        }}

        /* Header 3 Blocks */
        .header-container {{
            border: 1px solid #000;
            display: flex;
            min-height: 108px;
        }}
        .header-emitente {{
            width: 44%;
            padding: 3px 5px;
            border-right: 1px solid #000;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .header-danfe {{
            width: 22%;
            padding: 3px;
            border-right: 1px solid #000;
            text-align: center;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            align-items: center;
        }}
        .header-chave {{
            width: 34%;
            padding: 3px 4px;
            text-align: center;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}

        .tp-nf-box {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            margin: 1px 0;
        }}
        .tp-nf-num {{
            border: 1px solid #000;
            width: 15px;
            height: 15px;
            font-size: 10px;
            font-weight: bold;
            line-height: 14px;
            text-align: center;
        }}

        /* Tables and Cells */
        .danfe-table {{
            width: 100%;
            border-collapse: collapse;
            border: 1px solid #000;
            margin-bottom: 1px;
        }}
        .danfe-table td, .danfe-table th {{
            border: 1px solid #000;
            padding: 1px 3px;
            vertical-align: top;
        }}
        .danfe-table th {{
            background: #fff;
            font-size: 5.8px;
            font-weight: bold;
            text-align: center;
            text-transform: uppercase;
        }}

        .row-flex {{
            display: flex;
            width: 100%;
        }}

        /* Products Table */
        .items-table {{
            width: 100%;
            border-collapse: collapse;
            border: 1px solid #000;
            margin-bottom: 2px;
        }}
        .items-table th {{
            border: 1px solid #000;
            padding: 2px 1px;
            font-size: 5.5px;
            font-weight: bold;
            text-align: center;
            background: #fff;
            line-height: 1.05;
        }}
        .items-table td {{
            border-left: 1px solid #000;
            border-right: 1px solid #000;
            border-bottom: 1px solid #000;
            border-top: none;
            padding: 1.5px 2px;
            font-size: 7.5px;
            height: 14px;
        }}

        /* Floating Print Bar */
        .print-btn-bar {{
            position: fixed;
            bottom: 24px;
            right: 24px;
            background: #1e40af;
            color: #fff;
            padding: 11px 20px;
            border-radius: 50px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 6px 16px rgba(0,0,0,0.3);
            font-size: 13px;
            display: flex;
            align-items: center;
            gap: 8px;
            z-index: 9999;
            transition: all 0.2s ease;
        }}
        .print-btn-bar:hover {{
            background: #1d4ed8;
            transform: translateY(-2px);
        }}
    </style>
</head>
<body>

    <div class="print-btn-bar no-print" onclick="window.print()">
        <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4H7v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
        Imprimir DANFE
    </div>

    <div class="danfe-page">
        {homolog_strip_html}
        {watermark_html}
        <!-- CANHOTO DE RECEBIMENTO -->
        <div class="canhoto-container">
            <div class="row-flex">
                <div style="flex: 1;">
                    <div class="canhoto-header">
                        RECEBEMOS DE {pyhtml.escape(emit_nome)} OS PRODUTOS / SERVIÇOS CONSTANTES DA NOTA FISCAL INDICADO AO LADO
                    </div>
                    <div class="canhoto-body">
                        <div style="width: 25%; padding: 2px 4px; border-right: 1px solid #000;">
                            <span class="lbl">DATA DE RECEBIMENTO</span>
                        </div>
                        <div style="width: 75%; padding: 2px 4px;">
                            <span class="lbl">IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR</span>
                        </div>
                    </div>
                </div>
                <div class="canhoto-box-right">
                    <div style="font-size: 11px; font-weight: bold;">NF-e</div>
                    <div style="font-size: 10px; font-weight: bold;">Nº {nro_nota_fmt}</div>
                    <div style="font-size: 8.5px; font-weight: bold;">SÉRIE {serie_fmt}</div>
                </div>
            </div>
        </div>

        <div class="cut-line"></div>

        <!-- QUADRO 1: IDENTIFICAÇÃO DO EMITENTE / DANFE / CHAVE -->
        <div class="header-container">
            <!-- EMITENTE -->
            <div class="header-emitente">
                {f'''
                <div style="display: flex; gap: 8px; align-items: center; height: 100%;">
                    <div style="width: 75px; min-width: 75px; text-align: center; display: flex; align-items: center; justify-content: center;">
                        <img src="{pyhtml.escape(logo_src)}" alt="Logo" style="max-height: 72px; max-width: 75px; object-fit: contain;">
                    </div>
                    <div style="flex: 1; min-width: 0;">
                        <span class="lbl" style="font-size: 5.5px;">IDENTIFICAÇÃO DO EMITENTE</span>
                        <div class="val-bold" style="font-size: 9.5px; line-height: 1.15; margin-top: 1px;">{pyhtml.escape(emit_nome)}</div>
                        <div class="val" style="font-size: 7.5px; margin-top: 2px; line-height: 1.15;">
                            {pyhtml.escape(emit_end)}<br>
                            {pyhtml.escape(emit_bairro)} - CEP: {emit_cep}<br>
                            {pyhtml.escape(emit_mun)} - {pyhtml.escape(emit_uf)} {f"Fone: {pyhtml.escape(emit_fone)}" if emit_fone else ""}
                        </div>
                    </div>
                </div>
                ''' if logo_src else f'''
                <div>
                    <span class="lbl" style="font-size: 5.5px;">IDENTIFICAÇÃO DO EMITENTE</span>
                    <div class="val-bold" style="font-size: 11px; line-height: 1.15; margin-top: 2px;">{pyhtml.escape(emit_nome)}</div>
                    <div class="val" style="font-size: 8px; margin-top: 3px;">
                        {pyhtml.escape(emit_end)}<br>
                        {pyhtml.escape(emit_bairro)} - CEP: {emit_cep}<br>
                        {pyhtml.escape(emit_mun)} - {pyhtml.escape(emit_uf)} {f"Fone: {pyhtml.escape(emit_fone)}" if emit_fone else ""}
                    </div>
                </div>
                '''}
            </div>

            <!-- DANFE -->
            <div class="header-danfe">
                <div style="font-size: 15px; font-weight: bold; letter-spacing: 0.5px;">DANFE</div>
                <div style="font-size: 6px; font-weight: bold; line-height: 1.1;">
                    DOCUMENTO AUXILIAR DA<br>NOTA FISCAL ELETRÔNICA
                </div>
                <div class="tp-nf-box">
                    <div style="font-size: 6.5px; text-align: left; line-height: 1.15;">
                        0 - ENTRADA<br>
                        1 - SAÍDA
                    </div>
                    <div class="tp-nf-num">{tp_nf}</div>
                </div>
                <div>
                    <div style="font-size: 8.5px; font-weight: bold;">Nº {nro_nota_fmt} fl. 1 /1</div>
                    <div style="font-size: 8px; font-weight: bold;">SÉRIE {serie_fmt}</div>
                </div>
            </div>

            <!-- CHAVE DE ACESSO & BARCODE -->
            <div class="header-chave">
                <div>
                    {barcode_svg}
                    <div style="border-top: 1px solid #000; margin-top: 1px; padding-top: 1px;">
                        <span class="lbl" style="text-align: left;">CHAVE DE ACESSO</span>
                        <div class="val-bold" style="font-size: 7.8px; letter-spacing: 0.2px; text-align: center;">{chave_formatada}</div>
                    </div>
                </div>
                <div style="font-size: 6.5px; line-height: 1.2; text-align: center; border-top: 1px solid #000; padding-top: 2px; margin-top: 1px;">
                    Consulta de autenticidade no portal nacional da NF-e<br>
                    <strong>www.nfe.fazenda.gov.br/portal</strong><br>
                    ou no site da Sefaz Autorizadora
                </div>
            </div>
        </div>

        <!-- NATUREZA DE OPERAÇÃO & PROTOCOLO -->
        <table class="danfe-table" style="border-top: none; margin-bottom: 0;">
            <tr>
                <td style="width: 65%;">
                    <span class="lbl">NATUREZA DE OPERAÇÃO</span>
                    <span class="val-bold">{pyhtml.escape(nat_op_desc)}</span>
                </td>
                <td style="width: 35%;">
                    <span class="lbl">PROTOCOLO DE AUTORIZAÇÃO DE USO</span>
                    <span class="val-bold">{proto_data}</span>
                </td>
            </tr>
        </table>

        <!-- INSCRIÇÃO ESTADUAL & CNPJ -->
        <table class="danfe-table" style="border-top: none; margin-bottom: 2px;">
            <tr>
                <td style="width: 33.3%;">
                    <span class="lbl">INSCRIÇÃO ESTADUAL</span>
                    <span class="val-bold">{emit_ie}</span>
                </td>
                <td style="width: 33.3%;">
                    <span class="lbl">INSCRIÇÃO ESTADUAL DO SUBST. TRIB.</span>
                    <span class="val-bold"></span>
                </td>
                <td style="width: 33.4%;">
                    <span class="lbl">CNPJ / CPF</span>
                    <span class="val-bold">{emit_cnpj}</span>
                </td>
            </tr>
        </table>

        <!-- DESTINATÁRIO / REMETENTE -->
        <div class="sec-title">DESTINATÁRIO / REMETENTE</div>
        <table class="danfe-table">
            <tr>
                <td style="width: 68%;">
                    <span class="lbl">NOME / RAZÃO SOCIAL</span>
                    <span class="val-bold">{pyhtml.escape(dest_nome)}</span>
                </td>
                <td style="width: 20%;">
                    <span class="lbl">CNPJ / CPF</span>
                    <span class="val-bold">{dest_doc}</span>
                </td>
                <td style="width: 12%;">
                    <span class="lbl">DATA DA EMISSÃO</span>
                    <span class="val-bold">{dt_emiss}</span>
                </td>
            </tr>
            <tr>
                <td>
                    <span class="lbl">ENDEREÇO</span>
                    <span class="val">{pyhtml.escape(dest_end)}</span>
                </td>
                <td>
                    <span class="lbl">BAIRRO / DISTRITO</span>
                    <span class="val">{pyhtml.escape(dest_bairro)}</span>
                </td>
                <td>
                    <div class="row-flex" style="margin: 0; padding: 0;">
                        <div style="flex: 1;">
                            <span class="lbl">CEP</span>
                            <span class="val">{dest_cep}</span>
                        </div>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" style="padding: 0;">
                    <div class="row-flex">
                        <div style="width: 45%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">MUNICÍPIO</span>
                            <span class="val">{pyhtml.escape(dest_mun)}</span>
                        </div>
                        <div style="width: 35%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">FONE / FAX</span>
                            <span class="val">{dest_fone}</span>
                        </div>
                        <div style="width: 20%; padding: 1px 3px;">
                            <span class="lbl">UF</span>
                            <span class="val">{dest_uf}</span>
                        </div>
                    </div>
                </td>
                <td style="padding: 0;">
                    <div style="padding: 1px 3px;">
                        <span class="lbl">DATA SAÍDA / ENTRADA</span>
                        <span class="val">{dt_saida}</span>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" style="padding: 0;">
                    <div class="row-flex">
                        <div style="width: 100%; padding: 1px 3px;">
                            <span class="lbl">INSCRIÇÃO ESTADUAL</span>
                            <span class="val">{dest_ie}</span>
                        </div>
                    </div>
                </td>
                <td style="padding: 0;">
                    <div style="padding: 1px 3px;">
                        <span class="lbl">HORA DA SAÍDA</span>
                        <span class="val">{hora_saida}</span>
                    </div>
                </td>
            </tr>
        </table>

        <!-- CÁLCULO DO IMPOSTO -->
        <div class="sec-title">CÁLCULO DO IMPOSTO</div>
        <table class="danfe-table">
            <tr>
                <td style="width: 20%;">
                    <span class="lbl">BASE DE CÁLCULO DO ICMS</span>
                    <div class="val text-right">{format_money(base_icms)}</div>
                </td>
                <td style="width: 20%;">
                    <span class="lbl">VALOR DO ICMS</span>
                    <div class="val text-right">{format_money(valor_icms)}</div>
                </td>
                <td style="width: 20%;">
                    <span class="lbl">BASE CÁLC. ICMS SUBST.</span>
                    <div class="val text-right">{format_money(base_st)}</div>
                </td>
                <td style="width: 20%;">
                    <span class="lbl">VALOR DO ICMS SUBST.</span>
                    <div class="val text-right">{format_money(valor_st)}</div>
                </td>
                <td style="width: 20%;">
                    <span class="lbl">VALOR TOTAL DOS PRODUTOS</span>
                    <div class="val text-right font-bold">{format_money(total_prod)}</div>
                </td>
            </tr>
            <tr>
                <td style="width: 16.6%;">
                    <span class="lbl">VALOR DO FRETE</span>
                    <div class="val text-right">{format_money(frete)}</div>
                </td>
                <td style="width: 16.6%;">
                    <span class="lbl">VALOR DO SEGURO</span>
                    <div class="val text-right">{format_money(seguro)}</div>
                </td>
                <td style="width: 16.6%;">
                    <span class="lbl">DESCONTO</span>
                    <div class="val text-right">{format_money(desconto)}</div>
                </td>
                <td style="width: 16.6%;">
                    <span class="lbl">OUTRAS DESP. ACESS.</span>
                    <div class="val text-right">{format_money(outras_desp)}</div>
                </td>
                <td style="width: 16.6%;">
                    <span class="lbl">VALOR DO IPI</span>
                    <div class="val text-right">{format_money(valor_ipi)}</div>
                </td>
                <td style="width: 17%;">
                    <span class="lbl">VALOR TOTAL DA NOTA</span>
                    <div class="val text-right font-bold">{format_money(total_nota)}</div>
                </td>
            </tr>
        </table>

        <!-- TRANSPORTADOR / VOLUMES TRANSPORTADOS -->
        <div class="sec-title">TRANSPORTADOR / VOLUMES TRANSPORTADOS</div>
        <table class="danfe-table">
            <tr>
                <td style="width: 38%;">
                    <span class="lbl">RAZÃO SOCIAL</span>
                    <span class="val">{pyhtml.escape(str(order.get('NomeTransp') or ''))}</span>
                </td>
                <td style="width: 16%;">
                    <span class="lbl">FRETE POR CONTA</span>
                    <span class="val-bold">{frete_txt}</span>
                </td>
                <td style="width: 11%;">
                    <span class="lbl">CÓDIGO ANTT</span>
                    <span class="val">{pyhtml.escape(str(order.get('CodigoAntt') or ''))}</span>
                </td>
                <td style="width: 13%;">
                    <span class="lbl">PLACA DO VEÍCULO</span>
                    <span class="val">{pyhtml.escape(str(order.get('PlacaVeic') or order.get('PlacaVeiculo') or ''))}</span>
                </td>
                <td style="width: 4%;">
                    <span class="lbl">UF</span>
                    <span class="val">{pyhtml.escape(str(order.get('UfVeic') or ''))}</span>
                </td>
                <td style="width: 18%;">
                    <span class="lbl">CNPJ / CPF</span>
                    <span class="val">{format_cnpj_cpf(order.get('CnpjsTransp') or '')}</span>
                </td>
            </tr>
            <tr>
                <td>
                    <span class="lbl">ENDEREÇO</span>
                    <span class="val">{pyhtml.escape(str(order.get('EnderecoTransp') or ''))}</span>
                </td>
                <td colspan="3">
                    <span class="lbl">MUNICÍPIO</span>
                    <span class="val">{pyhtml.escape(str(order.get('CidadeTransp') or ''))}</span>
                </td>
                <td>
                    <span class="lbl">UF</span>
                    <span class="val">{pyhtml.escape(str(order.get('UfTransp') or ''))}</span>
                </td>
                <td>
                    <span class="lbl">INSCRIÇÃO ESTADUAL</span>
                    <span class="val">{pyhtml.escape(str(order.get('IeTransp') or ''))}</span>
                </td>
            </tr>
            <tr>
                <td colspan="6" style="padding: 0;">
                    <div class="row-flex">
                        <div style="width: 12%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">QUANTIDADE</span>
                            <span class="val">{order.get('QtdVol') or order.get('QtdVolume') or ''}</span>
                        </div>
                        <div style="width: 15%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">ESPÉCIE</span>
                            <span class="val">{pyhtml.escape(str(order.get('EspecieVol') or order.get('Especie') or ''))}</span>
                        </div>
                        <div style="width: 15%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">MARCA</span>
                            <span class="val">{pyhtml.escape(str(order.get('MarcaVol') or order.get('Marca') or ''))}</span>
                        </div>
                        <div style="width: 18%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">NUMERAÇÃO</span>
                            <span class="val">{pyhtml.escape(str(order.get('NumeracaoVol') or ''))}</span>
                        </div>
                        <div style="width: 20%; padding: 1px 3px; border-right: 1px solid #000;">
                            <span class="lbl">PESO BRUTO (Kg)</span>
                            <span class="val">{format_qty(order.get('PesoBruto')) if order.get('PesoBruto') else ''}</span>
                        </div>
                        <div style="width: 20%; padding: 1px 3px;">
                            <span class="lbl">PESO LÍQUIDO (Kg)</span>
                            <span class="val">{format_qty(order.get('PesoLiquido')) if order.get('PesoLiquido') else ''}</span>
                        </div>
                    </div>
                </td>
            </tr>
        </table>

        <!-- DADOS DO PRODUTO / SERVIÇOS -->
        <div class="sec-title">DADOS DO PRODUTO / SERVIÇOS</div>
        <table class="items-table">
            <thead>
                <tr>
                    <th style="width: 6%;">CÓDIGO DO<br>PROD. /</th>
                    <th style="width: 28%;">DESCRIÇÃO DO PRODUTO / SERVIÇO</th>
                    <th style="width: 6%;">NCM/SH</th>
                    <th style="width: 4%;">O/CST</th>
                    <th style="width: 4%;">CFOP</th>
                    <th style="width: 3%;">UNID.</th>
                    <th style="width: 5%;">QUANT.</th>
                    <th style="width: 6%;">VALOR<br>UNITÁRIO</th>
                    <th style="width: 7%;">VALOR<br>TOTAL</th>
                    <th style="width: 5%;">DESCONTO</th>
                    <th style="width: 6%;">BASE<br>CÁLC. ICMS</th>
                    <th style="width: 6%;">VALOR<br>ICMS</th>
                    <th style="width: 5%;">VALOR<br>IPI</th>
                    <th style="width: 4%;">ALÍQ.<br>ICMS</th>
                    <th style="width: 4%;">ALÍQ.<br>IPI</th>
                </tr>
            </thead>
            <tbody>
                {''.join(items_rows_html)}
            </tbody>
        </table>

        <!-- DADOS ADICIONAIS -->
        <div class="sec-title" style="margin-top: 14px;">DADOS ADICIONAIS</div>
        <table class="danfe-table">
            <tr>
                <td style="width: 65%; height: 85px; vertical-align: top;">
                    <span class="lbl">INFORMAÇÕES COMPLEMENTARES</span>
                    <div class="val" style="font-size: 7.5px; line-height: 1.25; margin-top: 2px;">
                        {inf_compl}
                    </div>
                </td>
                <td style="width: 35%; height: 85px; vertical-align: top;">
                    <span class="lbl">RESERVADO AO FISCO</span>
                    <div class="val"></div>
                </td>
            </tr>
        </table>
    </div>

    <script>
        window.addEventListener('load', function() {{
            if (window.location.search.includes('print=1')) {{
                window.print();
            }}
        }});
    </script>
</body>
</html>"""
    return html_content
