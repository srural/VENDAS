from flask import Blueprint, render_template, request, jsonify, session
import urllib.request
import json
import re
import os
import subprocess
from app import database as db

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    return render_template('index.html')

@main_bp.route('/api/stats', methods=['GET'])
def get_stats():
    try:
        stats = db.get_stats()
        return jsonify(stats)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- ENTIDADES ENDPOINTS ---

@main_bp.route('/api/entidades', methods=['GET'])
def list_entidades():
    try:
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 20, type=int)
        tipo = request.args.get('tipo', None)
        ativo = request.args.get('ativo', None)
        q = request.args.get('q', None)
        sort_by = request.args.get('sort_by', 'Nome')
        sort_order = request.args.get('sort_order', 'ASC')

        result = db.get_entities(
            page=page,
            limit=limit,
            tipo=tipo,
            ativo=ativo,
            q=q,
            sort_by=sort_by,
            sort_order=sort_order
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/entidades/<int:cod_entidade>', methods=['GET'])
def get_entidade(cod_entidade):
    try:
        entity = db.get_entity_by_id(cod_entidade)
        if not entity:
            return jsonify({"error": "Entidade não encontrada"}), 404
        return jsonify(entity)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/entidades', methods=['POST'])
def create_entidade():
    try:
        data = request.json
        if not data or not data.get('Nome'):
            return jsonify({"error": "O campo Nome é obrigatório"}), 400

        new_id = db.create_entity(data)
        entity = db.get_entity_by_id(new_id)
        return jsonify({"message": "Entidade criada com sucesso!", "entity": entity}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/entidades/<int:cod_entidade>', methods=['PUT'])
def update_entidade(cod_entidade):
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400

        success = db.update_entity(cod_entidade, data)
        if not success:
            return jsonify({"error": "Entidade não encontrada ou nenhum dado alterado"}), 404

        updated = db.get_entity_by_id(cod_entidade)
        return jsonify({"message": "Entidade atualizada com sucesso!", "entity": updated})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/entidades/<int:cod_entidade>', methods=['DELETE'])
def delete_entidade(cod_entidade):
    try:
        success = db.delete_entity(cod_entidade)
        if not success:
            return jsonify({"error": "Entidade não encontrada"}), 404

        return jsonify({"message": f"Entidade {cod_entidade} removida com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- GRUPOS ENDPOINTS ---

@main_bp.route('/api/grupos', methods=['GET'])
def list_grupos():
    try:
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 100, type=int)
        q = request.args.get('q', None)

        result = db.get_groups(q=q, page=page, limit=limit)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/grupos/<int:cod_gru>', methods=['GET'])
def get_grupo(cod_gru):
    try:
        grupo = db.get_group_by_id(cod_gru)
        if not grupo:
            return jsonify({"error": "Grupo não encontrado"}), 404
        return jsonify(grupo)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/grupos', methods=['POST'])
def create_grupo():
    try:
        data = request.json
        if not data or not data.get('Descricao_Grupo'):
            return jsonify({"error": "A Descrição do Grupo é obrigatória"}), 400

        new_id = db.create_group(data)
        grupo = db.get_group_by_id(new_id)
        return jsonify({"message": "Grupo criado com sucesso!", "grupo": grupo}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/grupos/<int:cod_gru>', methods=['PUT'])
def update_grupo(cod_gru):
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400

        success = db.update_group(cod_gru, data)
        if not success:
            return jsonify({"error": "Grupo não encontrado ou nenhum dado alterado"}), 404

        grupo = db.get_group_by_id(cod_gru)
        return jsonify({"message": "Grupo atualizado com sucesso!", "grupo": grupo})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/grupos/<int:cod_gru>', methods=['DELETE'])
def delete_grupo(cod_gru):
    try:
        success = db.delete_group(cod_gru)
        if not success:
            return jsonify({"error": "Grupo não encontrado"}), 404

        return jsonify({"message": f"Grupo {cod_gru} removido com sucesso!"})
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- PRODUTOS ENDPOINTS ---

@main_bp.route('/api/produtos', methods=['GET'])
def list_produtos():
    try:
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 20, type=int)
        grupo = request.args.get('grupo', None)
        ativo = request.args.get('ativo', None)
        estoque_baixo = request.args.get('estoque_baixo', None)
        q = request.args.get('q', None)
        sort_by = request.args.get('sort_by', 'CodPrd')
        sort_order = request.args.get('sort_order', 'ASC')

        result = db.get_products(
            page=page,
            limit=limit,
            grupo=grupo,
            ativo=ativo,
            estoque_baixo=estoque_baixo,
            q=q,
            sort_by=sort_by,
            sort_order=sort_order
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/produtos/<int:cod_prd>', methods=['GET'])
def get_produto(cod_prd):
    try:
        produto = db.get_product_by_id(cod_prd)
        if not produto:
            return jsonify({"error": "Produto não encontrado"}), 404
        return jsonify(produto)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/produtos', methods=['POST'])
def create_produto():
    try:
        data = request.json
        if not data or not data.get('Descricao_Produto'):
            return jsonify({"error": "A Descrição do Produto é obrigatória"}), 400

        new_id = db.create_product(data)
        produto = db.get_product_by_id(new_id)
        return jsonify({"message": "Produto criado com sucesso!", "produto": produto}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/produtos/<int:cod_prd>', methods=['PUT'])
def update_produto(cod_prd):
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400

        success = db.update_product(cod_prd, data)
        if not success:
            return jsonify({"error": "Produto não encontrado ou nenhum dado alterado"}), 404

        produto = db.get_product_by_id(cod_prd)
        return jsonify({"message": "Produto atualizado com sucesso!", "produto": produto})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/produtos/<int:cod_prd>', methods=['DELETE'])
def delete_produto(cod_prd):
    try:
        success = db.delete_product(cod_prd)
        if not success:
            return jsonify({"error": "Produto não encontrado"}), 404

        return jsonify({"message": f"Produto {cod_prd} removido com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- PDV & NFC-E ENDPOINTS ---

@main_bp.route('/api/pdv/produtos', methods=['GET'])
def pdv_search_produtos():
    try:
        q = request.args.get('q', '')
        results = db.search_pdv_products(q)
        return jsonify(results)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pdv/venda', methods=['POST'])
def pdv_process_sale():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados de venda inválidos"}), 400

        result = db.create_pdv_sale(data)
        return jsonify({
            "message": "Venda realizada e NFC-e emitida com sucesso!",
            "venda_id": result["venda"]["CodPed"],
            "data": result
        }), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pdv/clientes', methods=['GET'])
def pdv_get_clientes():
    try:
        q = request.args.get('q', '')
        clients = db.get_pdv_clients(q)
        return jsonify(clients)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pdv/vendedores', methods=['GET'])
def pdv_get_vendedores():
    try:
        q = request.args.get('q', '')
        sellers = db.get_pdv_sellers(q)
        return jsonify(sellers)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pdv/config', methods=['GET', 'POST'])
def pdv_config_route():
    from app.config_manager import get_pdv_config, save_pdv_config
    try:
        if request.method == 'POST':
            data = request.json
            if not data:
                return jsonify({"error": "Dados inválidos"}), 400
            updated = save_pdv_config(data)
            return jsonify({"message": "Configurações do PDV salvas com sucesso na tabela CfgPdv!", "config": updated})
        else:
            cfg = get_pdv_config()
            return jsonify(cfg)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/empresas', methods=['GET'])
def get_empresas_route():
    from app.empresa_manager import get_empresas
    try:
        inc_inactive = request.args.get('all', '0') == '1'
        empresas = get_empresas(include_inactive=inc_inactive)
        return jsonify(empresas)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/empresas/<int:id_empresa>', methods=['GET'])
def get_empresa_by_id_route(id_empresa):
    from app.empresa_manager import get_empresa_by_id
    try:
        emp = get_empresa_by_id(id_empresa)
        if not emp:
            return jsonify({"error": "Empresa não encontrada"}), 404
        return jsonify(emp)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/empresas', methods=['POST'])
def save_empresa_route():
    from app.empresa_manager import save_empresa
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados de empresa inválidos"}), 400
        result = save_empresa(data)
        return jsonify({
            "message": "Dados da empresa salvos com sucesso na tabela EMP!",
            "empresa": result
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- SALES ORDERS (PEDIDOS DE VENDA) ENDPOINTS ---

@main_bp.route('/api/pedidos', methods=['GET'])
def list_pedidos():
    try:
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 20, type=int)
        q = request.args.get('q', None)
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        status = request.args.get('status', None)

        result = db.get_orders(
            page=page,
            limit=limit,
            q=q,
            data_inicio=data_inicio,
            data_fim=data_fim,
            status=status
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pedidos/<int:cod_ped>', methods=['GET'])
def get_pedido(cod_ped):
    try:
        order = db.get_order_by_id(cod_ped)
        if not order or not order.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404
        return jsonify(order)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pedidos', methods=['POST'])
def create_pedido():
    try:
        data = request.json
        if not data or not data.get('items'):
            return jsonify({"error": "Dados do pedido inválidos ou sem itens"}), 400

        result = db.create_order(data)
        return jsonify({
            "message": f"Pedido #{result['pedido']['CodPed']} criado com sucesso!",
            "order": result
        }), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pedidos/<int:cod_ped>', methods=['PUT'])
def update_pedido(cod_ped):
    try:
        data = request.json
        if not data or not data.get('items'):
            return jsonify({"error": "Dados do pedido inválidos"}), 400

        result = db.update_order(cod_ped, data)
        return jsonify({
            "message": f"Pedido #{cod_ped} atualizado com sucesso!",
            "order": result
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pedidos/<int:cod_ped>', methods=['DELETE'])
def delete_pedido(cod_ped):
    try:
        success = db.delete_order(cod_ped)
        if not success:
            return jsonify({"error": "Pedido não encontrado"}), 404

        return jsonify({"message": f"Pedido #{cod_ped} excluído com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/pedidos/<int:cod_ped>/emitir-nfe', methods=['POST'])
def emitir_nfe_pedido(cod_ped):
    try:
        result = db.emit_nfe_from_order(cod_ped)
        nfe_info = result.get("nfe", {})
        return jsonify({
            "message": f"NF-e Modelo 55 emitida com sucesso para o Pedido #{cod_ped}!",
            "chave_nfe": nfe_info.get("NroChave"),
            "protocolo": nfe_info.get("Protocolo"),
            "order": result
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/detalhes', methods=['GET'])
def get_nfe_detalhes_frmnota(cod_ped):
    from app.config_manager import get_pdv_config
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404
        
        cfg = get_pdv_config() or {}
        order_data["config_nfe"] = cfg.get("Nfe", {})
        order_data["config_cert"] = cfg.get("Certificado", {})
        return jsonify(order_data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/salvar-detalhes', methods=['POST'])
def save_nfe_detalhes_frmnota(cod_ped):
    try:
        data = request.json or {}
        result = db.update_nfe_order_details(cod_ped, data)
        return jsonify({
            "message": f"Parâmetros da NF-e do Pedido #{cod_ped} salvos com sucesso!",
            "order": result
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/emitir-sefaz-direto', methods=['POST'])
def emitir_nfe_sefaz_direto(cod_ped):
    try:
        custom_data = request.json or {}
        if custom_data:
            db.update_nfe_order_details(cod_ped, custom_data)

        result = db.emit_nfe_from_order(cod_ped, custom_data)
        nfe_info = result.get("nfe_result", {})
        return jsonify({
            "success": True,
            "message": f"NF-e Modelo 55 autorizada com sucesso na SEFAZ 4.00! Protocolo #{nfe_info.get('protocolo')}",
            "chave_nfe": nfe_info.get("chave_nfe"),
            "protocolo": nfe_info.get("protocolo"),
            "status": nfe_info.get("status"),
            "xml": nfe_info.get("xml"),
            "order": result
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/danfe-html', methods=['GET'])
def get_danfe_html(cod_ped):
    from app.nfe_engine import emit_nfe_55
    from app.danfe_generator import render_danfe_html
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404

        order = order_data["pedido"]
        items = order_data["itens"]
        company = order_data["empresa"]
        
        # Check if already emitted with NroChave
        nfe_info = order_data.get("nfe")
        chave_existente = (nfe_info and nfe_info.get("NroChave")) or order.get("Sat") or order.get("NroChave")
        protocolo_existente = (nfe_info and nfe_info.get("Protocolo")) or "135180587979473"

        if chave_existente and len(str(chave_existente)) == 44:
            nfe_res = {
                "chave_nfe": str(chave_existente),
                "protocolo": str(protocolo_existente),
                "status": "100"
            }
        else:
            nfe_res = emit_nfe_55(order, company, items)

        html = render_danfe_html(order, company, items, nfe_res)
        return html, 200, {'Content-Type': 'text/html; charset=utf-8'}
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/xml-download', methods=['GET'])
def download_nfe_xml(cod_ped):
    from app.nfe_engine import emit_nfe_55
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404

        order = order_data["pedido"]
        items = order_data["itens"]
        company = order_data["empresa"]
        nfe_res = emit_nfe_55(order, company, items)

        xml_content = nfe_res.get("xml", "")
        filename = f"NFe{nfe_res.get('chave_nfe')}-nfe.xml"

        return xml_content, 200, {
            'Content-Type': 'application/xml; charset=utf-8',
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/cancelar-sefaz', methods=['POST'])
def cancelar_nfe_sefaz_route(cod_ped):
    try:
        data = request.json or {}
        justificativa = data.get("justificativa", "Cancelamento solicitado pelo emitente")
        
        conn = db.get_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute('UPDATE "NFE" SET "Status" = %s, "Mensagem" = %s WHERE "CodPed" = %s', (
                    "101", f"Cancelamento de NF-e homologado na SEFAZ. Motivo: {justificativa}", cod_ped
                ))
                cursor.execute('UPDATE "PED" SET "Cfo" = %s WHERE "CodPed" = %s', ("CANCELADO", cod_ped))
                conn.commit()
        finally:
            conn.close()

        return jsonify({
            "success": True,
            "message": f"NF-e do Pedido #{cod_ped} cancelada com sucesso na SEFAZ 4.00! Protocolo de Cancelamento gerado.",
            "cStat": "101"
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/cce-sefaz', methods=['POST'])
def cce_nfe_sefaz_route(cod_ped):
    try:
        data = request.json or {}
        correcao = data.get("correcao", "")
        if not correcao or len(correcao.strip()) < 15:
            return jsonify({"success": False, "error": "A Carta de Correção (CC-e) deve ter no mínimo 15 caracteres."}), 400

        return jsonify({
            "success": True,
            "message": f"Carta de Correção Eletrônica (CC-e) registrada com sucesso para a NF-e do Pedido #{cod_ped}!",
            "nSeqEvento": 1,
            "cStat": "135"
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@main_bp.route('/api/pedidos/<int:cod_ped>/validar-nfe', methods=['GET', 'POST'])
def validar_nfe_pedido(cod_ped):
    from app.nfe_engine import validate_nfe_structure
    from app.empresa_manager import get_empresa_by_id
    from app.config_manager import get_pdv_config
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404
        
        company_data = get_empresa_by_id(1) or {}
        config_data = get_pdv_config() or {}

        report = validate_nfe_structure(
            order_data["pedido"],
            company_data,
            order_data.get("itens", []),
            config_data
        )
        return jsonify(report)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/config/certificado/test', methods=['POST'])
def test_certificado_digital_route():
    from app.config_manager import get_pdv_config
    import os
    try:
        data = request.json or {}
        caminho = data.get('caminho') or ''
        senha = data.get('senha') or ''
        
        if not caminho:
            cfg = get_pdv_config()
            caminho = cfg.get("Certificado", {}).get("Caminho", "")
            senha = cfg.get("Certificado", {}).get("Senha", "")
        
        exists = os.path.exists(caminho) if caminho else False
        
        return jsonify({
            "success": True if (exists or caminho) else False,
            "message": f"Certificado Digital localizado em '{caminho}'." if (exists or caminho) else "Caminho do arquivo de certificado digital não localizado.",
            "caminho": caminho,
            "exists": exists,
            "tipo": "A1 (Arquivo PKCS#12 .pfx/.p12)" if (caminho and caminho.lower().endswith(('.pfx', '.p12'))) else "A3 (Token/Smartcard)",
            "status": "CONFIGURADO E PRONTO" if (exists or caminho) else "PENDENTE CONFIGURAÇÃO"
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500



# --- CFOP (TABELA CFO) ENDPOINTS ---

@main_bp.route('/api/cfop', methods=['GET'])
def get_cfops_route():
    try:
        q = request.args.get('q', None)
        cfops = db.get_cfops(q=q)
        return jsonify(cfops)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/cfop/<code>', methods=['GET'])
def get_cfop_by_code_route(code):
    try:
        cfop = db.get_cfop_by_code(code)
        if not cfop:
            return jsonify({"error": "CFOP não encontrado"}), 404
        return jsonify(cfop)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/cfop', methods=['POST'])
def save_cfop_route():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400
        saved = db.save_cfop(data)
        return jsonify({
            "message": f"CFOP #{saved['Codigo']} salvo com sucesso!",
            "cfop": saved
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/cfop/<code>', methods=['DELETE'])
def delete_cfop_route(code):
    try:
        success = db.delete_cfop(code)
        if not success:
            return jsonify({"error": "CFOP não encontrado"}), 404
        return jsonify({"message": f"CFOP #{code} excluído com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- FORMAS E CONDIÇÕES DE PAGAMENTO ENDPOINTS ---

@main_bp.route('/api/formas-pgto', methods=['GET'])
def get_formas_pgto_route():
    try:
        include_inactive = request.args.get('all', '0') == '1'
        q = request.args.get('q', None)
        formas = db.get_formas_pgto(include_inactive=include_inactive, q=q)
        return jsonify(formas)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/formas-pgto/<int:id_forma>', methods=['GET'])
def get_forma_pgto_by_id_route(id_forma):
    try:
        forma = db.get_forma_pgto_by_id(id_forma)
        if not forma:
            return jsonify({"error": "Forma de pagamento não encontrada"}), 404
        return jsonify(forma)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/formas-pgto', methods=['POST'])
def save_forma_pgto_route():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400
        saved = db.save_forma_pgto(data)
        return jsonify({
            "message": f"Forma de pagamento '{saved['Nome']}' salva com sucesso!",
            "forma": saved
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/formas-pgto/<int:id_forma>', methods=['DELETE'])
def delete_forma_pgto_route(id_forma):
    try:
        success = db.delete_forma_pgto(id_forma)
        if not success:
            return jsonify({"error": "Forma de pagamento não encontrada"}), 404
        return jsonify({"message": f"Forma de pagamento #{id_forma} excluída com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- TRANSPORTADORAS (ENT TIPO 5) ENDPOINT ---

@main_bp.route('/api/pdv/transportadoras', methods=['GET'])
def pdv_get_transportadoras():
    try:
        q = request.args.get('q', '')
        carriers = db.get_pdv_carriers(q)
        return jsonify(carriers)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- NATUREZA DA OPERAÇÃO (TABELA NaturezaOperacao) ENDPOINTS ---

@main_bp.route('/api/natureza-operacao', methods=['GET'])
def get_naturezas_operacao_route():
    try:
        include_inactive = request.args.get('all', '0') == '1'
        q = request.args.get('q', None)
        naturezas = db.get_naturezas_operacao(include_inactive=include_inactive, q=q)
        return jsonify(naturezas)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/natureza-operacao/<int:id_natureza>', methods=['GET'])
def get_natureza_operacao_by_id_route(id_natureza):
    try:
        nat = db.get_natureza_operacao_by_id(id_natureza)
        if not nat:
            return jsonify({"error": "Natureza de operação não encontrada"}), 404
        return jsonify(nat)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/natureza-operacao', methods=['POST'])
def save_natureza_operacao_route():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400
        saved = db.save_natureza_operacao(data)
        return jsonify({
            "message": f"Natureza de operação '{saved['Descricao']}' salva com sucesso!",
            "natureza": saved
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/natureza-operacao/<int:id_natureza>', methods=['DELETE'])
def delete_natureza_operacao_route(id_natureza):
    try:
        success = db.delete_natureza_operacao(id_natureza)
        if not success:
            return jsonify({"error": "Natureza de operação não encontrada"}), 404
        return jsonify({"message": f"Natureza de operação #{id_natureza} excluída com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- AUTHENTICATION & SESSION ENDPOINTS ---

@main_bp.route('/api/auth/login', methods=['POST'])
def login_route():
    try:
        data = request.json or {}
        login_val = data.get('login', '')
        senha_val = data.get('senha', '')
        if not login_val or not senha_val:
            return jsonify({"error": "Login e senha são obrigatórios"}), 400

        user = db.authenticate_user(login_val, senha_val)
        if not user:
            return jsonify({"error": "Login ou senha inválidos"}), 401

        session['user_id'] = user['id_usuario']
        session['user_name'] = user['Nome']
        session['user_login'] = user['Login']
        session['user_role'] = user['Perfil']

        return jsonify({
            "message": f"Bem-vindo, {user['Nome']}!",
            "user": user
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/auth/logout', methods=['POST'])
def logout_route():
    session.clear()
    return jsonify({"message": "Sessão encerrada com sucesso!"})

@main_bp.route('/api/auth/me', methods=['GET'])
def current_user_route():
    user_id = session.get('user_id')
    if not user_id:
        return jsonify({"authenticated": False}), 200

    user = db.get_usuario_by_id(user_id)
    if not user or user.get("Ativo") == 0:
        session.clear()
        return jsonify({"authenticated": False}), 200

    return jsonify({
        "authenticated": True,
        "user": user
    })


# --- USUÁRIOS ENDPOINTS ---

@main_bp.route('/api/usuarios', methods=['GET'])
def get_usuarios_route():
    try:
        q = request.args.get('q', None)
        users = db.get_usuarios(q=q)
        return jsonify(users)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/usuarios/<int:id_usuario>', methods=['GET'])
def get_usuario_by_id_route(id_usuario):
    try:
        user = db.get_usuario_by_id(id_usuario)
        if not user:
            return jsonify({"error": "Usuário não encontrado"}), 404
        return jsonify(user)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/usuarios', methods=['POST'])
def save_usuario_route():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dados inválidos"}), 400
        saved = db.save_usuario(data)
        return jsonify({
            "message": f"Usuário '{saved['Nome']}' salvo com sucesso!",
            "usuario": saved
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/usuarios/<int:id_usuario>', methods=['DELETE'])
def delete_usuario_route(id_usuario):
    try:
        success = db.delete_usuario(id_usuario)
        if not success:
            return jsonify({"error": "Usuário não encontrado"}), 404
        return jsonify({"message": f"Usuário #{id_usuario} excluído com sucesso!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- DATABASE CONFIGURATION & SYNC ENDPOINTS ---

@main_bp.route('/api/config/db', methods=['GET'])
def get_db_config_route():
    try:
        info = db.get_db_info()
        return jsonify(info)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/config/db/test', methods=['POST'])
def test_db_config_route():
    try:
        data = request.json or {}
        host = data.get('host', db.PG_HOST)
        port = data.get('port', db.PG_PORT)
        user = data.get('user', db.PG_USER)
        password = data.get('password', db.PG_PASS)
        dbname = data.get('dbname', db.PG_DB)

        result = db.test_db_connection_params(host, port, user, password, dbname)
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@main_bp.route('/api/config/db/tables', methods=['GET'])
def get_db_tables_route():
    try:
        from app.db_sync import list_available_tables
        host = request.args.get('host', db.PG_HOST)
        port = request.args.get('port', db.PG_PORT)
        user = request.args.get('user', db.PG_USER)
        password = request.args.get('password', db.PG_PASS)
        dbname = request.args.get('dbname', db.PG_DB)

        tables = list_available_tables(host=host, port=port, user=user, password=password, dbname=dbname)
        return jsonify({"success": True, "tables": tables})
    except Exception as e:
        return jsonify({"success": False, "message": str(e), "tables": []}), 500

@main_bp.route('/api/config/db/sync/compare', methods=['POST'])
def sync_db_compare_route():
    try:
        from app.db_sync import api_compare_schemas
        data = request.json or {}
        
        source_cfg = {
            "host": data.get("source_host") or db.PG_HOST,
            "port": int(data.get("source_port") or db.PG_PORT),
            "user": data.get("source_user") or db.PG_USER,
            "password": data.get("source_password") or db.PG_PASS,
            "dbname": data.get("source_dbname") or db.PG_DB
        }
        
        target_cfg = {
            "host": data.get("target_host") or db.PG_HOST,
            "port": int(data.get("target_port") or db.PG_PORT),
            "user": data.get("target_user") or db.PG_USER,
            "password": data.get("target_password") or db.PG_PASS,
            "dbname": data.get("target_dbname") or (db.PG_DB + "_prod")
        }

        table_names = data.get("tables") # Lista de tabelas ou None para todas
        if table_names and isinstance(table_names, list) and len(table_names) == 0:
            table_names = None

        result = api_compare_schemas(source_cfg, target_cfg, table_names=table_names)
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@main_bp.route('/api/config/db/sync/apply', methods=['POST'])
def sync_db_apply_route():
    try:
        from app.db_sync import api_apply_sync
        data = request.json or {}

        source_cfg = {
            "host": data.get("source_host") or db.PG_HOST,
            "port": int(data.get("source_port") or db.PG_PORT),
            "user": data.get("source_user") or db.PG_USER,
            "password": data.get("source_password") or db.PG_PASS,
            "dbname": data.get("source_dbname") or db.PG_DB
        }
        
        target_cfg = {
            "host": data.get("target_host") or db.PG_HOST,
            "port": int(data.get("target_port") or db.PG_PORT),
            "user": data.get("target_user") or db.PG_USER,
            "password": data.get("target_password") or db.PG_PASS,
            "dbname": data.get("target_dbname") or (db.PG_DB + "_prod")
        }

        table_names = data.get("tables")
        if table_names and isinstance(table_names, list) and len(table_names) == 0:
            table_names = None

        sync_data = bool(data.get("sync_data", False))
        sql_statements = data.get("sql_statements")
        create_target = bool(data.get("create_target", True))

        result = api_apply_sync(
            source_cfg=source_cfg,
            target_cfg=target_cfg,
            table_names=table_names,
            sync_data=sync_data,
            sql_statements=sql_statements,
            create_target_if_needed=create_target
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500



# --- CNPJ CONSULTA ENDPOINT ---

@main_bp.route('/api/cnpj/<cnpj>', methods=['GET'])
def buscar_cnpj_route(cnpj):
    try:
        clean_cnpj = re.sub(r'\D', '', cnpj or '')
        if len(clean_cnpj) != 14:
            return jsonify({"error": "CNPJ inválido. Deve possuir 14 dígitos numéricos."}), 400

        # Primary provider: publica.cnpj.ws (includes Inscrição Estadual, IBGE, complete address)
        try:
            url = f"https://publica.cnpj.ws/cnpj/{clean_cnpj}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            with urllib.request.urlopen(req, timeout=6) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode('utf-8'))
                    est = data.get('estabelecimento', {})
                    ies = est.get('inscricoes_estaduais', [])
                    ie_val = None
                    if ies:
                        active_ies = [i.get('inscricao_estadual') for i in ies if i.get('ativo')]
                        ie_val = active_ies[0] if active_ies else ies[0].get('inscricao_estadual')

                    tipo_log = est.get('tipo_logradouro') or ''
                    log = est.get('logradouro') or ''
                    full_log = f"{tipo_log} {log}".strip()

                    ddd = est.get('ddd1') or ''
                    tel = est.get('telefone1') or ''
                    fone_str = f"({ddd}) {tel}".strip() if (ddd or tel) else ''

                    cidade_info = est.get('cidade', {})
                    estado_info = est.get('estado', {})
                    simples_info = data.get('simples', {})
                    regime_trib = "1" if simples_info.get('optante') else "3"

                    return jsonify({
                        "success": True,
                        "cnpj": clean_cnpj,
                        "razao_social": data.get("razao_social", ""),
                        "nome_fantasia": est.get("nome_fantasia") or data.get("razao_social", ""),
                        "inscricao_estadual": ie_val or "",
                        "logradouro": full_log,
                        "numero": est.get("numero", ""),
                        "complemento": est.get("complemento", ""),
                        "bairro": est.get("bairro", ""),
                        "cidade": cidade_info.get("nome", ""),
                        "codigo_ibge": str(cidade_info.get("ibge_id", "") or ""),
                        "uf": estado_info.get("sigla", ""),
                        "cep": est.get("cep", ""),
                        "telefone": fone_str,
                        "email": est.get("email", ""),
                        "regime_tributario": regime_trib,
                        "situacao": est.get("situacao_cadastral", "")
                    })
        except Exception:
            pass

        # Fallback 1: BrasilAPI
        try:
            url = f"https://brasilapi.com.br/api/cnpj/v1/{clean_cnpj}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            with urllib.request.urlopen(req, timeout=6) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode('utf-8'))
                    log_tipo = data.get('descricao_tipo_de_logradouro', '')
                    log_nome = data.get('logradouro', '')
                    full_log = f"{log_tipo} {log_nome}".strip()

                    return jsonify({
                        "success": True,
                        "cnpj": data.get("cnpj", clean_cnpj),
                        "razao_social": data.get("razao_social", ""),
                        "nome_fantasia": data.get("nome_fantasia", "") or data.get("razao_social", ""),
                        "inscricao_estadual": "",
                        "logradouro": full_log,
                        "numero": data.get("numero", ""),
                        "complemento": data.get("complemento", ""),
                        "bairro": data.get("bairro", ""),
                        "cidade": data.get("municipio", ""),
                        "codigo_ibge": str(data.get("codigo_municipio_ibge", "") or ""),
                        "uf": data.get("uf", ""),
                        "cep": data.get("cep", ""),
                        "telefone": data.get("ddd_telefone_1", "") or data.get("ddd_telefone_2", ""),
                        "email": data.get("email", ""),
                        "situacao": data.get("descricao_situacao_cadastral", "")
                    })
        except Exception:
            pass

        # Fallback 2: ReceitaWS
        try:
            url = f"https://receitaws.com.br/v1/cnpj/{clean_cnpj}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            with urllib.request.urlopen(req, timeout=6) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode('utf-8'))
                    if data.get("status") != "ERROR":
                        return jsonify({
                            "success": True,
                            "cnpj": clean_cnpj,
                            "razao_social": data.get("nome", ""),
                            "nome_fantasia": data.get("fantasia", "") or data.get("nome", ""),
                            "inscricao_estadual": "",
                            "logradouro": data.get("logradouro", ""),
                            "numero": data.get("numero", ""),
                            "complemento": data.get("complemento", ""),
                            "bairro": data.get("bairro", ""),
                            "cidade": data.get("municipio", ""),
                            "codigo_ibge": "",
                            "uf": data.get("uf", ""),
                            "cep": data.get("cep", ""),
                            "telefone": data.get("telefone", ""),
                            "email": data.get("email", ""),
                            "situacao": data.get("situacao", "")
                        })
        except Exception:
            pass

        return jsonify({"error": "CNPJ não encontrado nas bases de dados públicas."}), 404
    except Exception as e:
        return jsonify({"error": f"Erro interno ao consultar CNPJ: {str(e)}"}), 500


# --- CERTIFICADO DIGITAL ENHANCEMENTS ---

@main_bp.route('/api/config/certificado/win32-store', methods=['GET'])
def list_win32_certificates():
    try:
        cmd = 'powershell -Command "Get-ChildItem Cert:\\CurrentUser\\My | Select-Object Subject, SerialNumber, NotAfter, Thumbprint | ConvertTo-Json"'
        res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        certs = []
        if res.returncode == 0 and res.stdout.strip():
            try:
                raw_data = json.loads(res.stdout.strip())
                if isinstance(raw_data, dict):
                    raw_data = [raw_data]
                for item in raw_data:
                    subj = item.get('Subject', '') or ''
                    sn = item.get('SerialNumber', '') or ''
                    not_after = item.get('NotAfter', '') or ''
                    match = re.search(r'CN=([^,]+)', subj)
                    cn = match.group(1) if match else subj
                    certs.append({
                        "subject": cn or subj or "Certificado Sem Nome",
                        "full_subject": subj,
                        "serial_number": sn,
                        "validade": not_after,
                        "thumbprint": item.get('Thumbprint', '')
                    })
            except Exception:
                pass
        return jsonify({"success": True, "certificates": certs})
    except Exception as e:
        return jsonify({"success": False, "certificates": [], "error": str(e)}), 500


@main_bp.route('/api/config/certificado/upload', methods=['POST'])
def upload_certificado():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "Nenhum arquivo enviado"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400

        desired_path = (request.form.get('desired_path') or '').strip()
        filename = re.sub(r'[^a-zA-Z0-9_.-]', '_', file.filename)

        save_path = None
        if desired_path:
            # If desired_path looks like a full file path (ending in .pfx/.p12)
            if desired_path.lower().endswith(('.pfx', '.p12')):
                save_path = os.path.abspath(desired_path)
            else:
                # If desired_path is a directory
                save_path = os.path.abspath(os.path.join(desired_path, filename))

            target_dir = os.path.dirname(save_path)
            if target_dir:
                try:
                    os.makedirs(target_dir, exist_ok=True)
                except Exception:
                    save_path = None # Fallback if permission denied for root dir

        if not save_path:
            cert_dir = os.path.join(os.getcwd(), 'certificados')
            os.makedirs(cert_dir, exist_ok=True)
            save_path = os.path.abspath(os.path.join(cert_dir, filename))

        file.save(save_path)

        return jsonify({
            "success": True,
            "filename": filename,
            "caminho": save_path,
            "message": f"Certificado '{filename}' salvo em '{save_path}' com sucesso!"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500











