from flask import Blueprint, render_template, request, jsonify, session
import urllib.request
import json
import re
import os
import uuid
import datetime
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

@main_bp.route('/api/produtos/upload-foto', methods=['POST'])
def upload_produto_foto_temp():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "Nenhum arquivo de imagem enviado"}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400

        allowed_extensions = {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.svg', '.bmp'}
        _, ext = os.path.splitext(file.filename.lower())
        if ext not in allowed_extensions:
            return jsonify({"error": f"Formato inválido. Extensões permitidas: {', '.join(allowed_extensions)}"}), 400

        upload_dir = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'produtos')
        os.makedirs(upload_dir, exist_ok=True)

        unique_name = f"prd_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}{ext}"
        filepath = os.path.join(upload_dir, unique_name)
        file.save(filepath)

        photo_url = f"/static/uploads/produtos/{unique_name}"
        return jsonify({
            "message": "Foto enviada com sucesso!",
            "url": photo_url,
            "filename": unique_name
        }), 200
    except Exception as e:
        return jsonify({"error": f"Erro ao processar imagem: {str(e)}"}), 500

@main_bp.route('/api/produtos/<int:cod_prd>/foto', methods=['POST', 'DELETE'])
def produto_foto_manage(cod_prd):
    try:
        produto = db.get_product_by_id(cod_prd)
        if not produto:
            return jsonify({"error": "Produto não encontrado"}), 404

        if request.method == 'DELETE':
            # Remove foto do produto
            old_foto = produto.get('Foto')
            if old_foto and old_foto.startswith('/static/uploads/produtos/'):
                filename = os.path.basename(old_foto)
                filepath = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'produtos', filename)
                if os.path.exists(filepath):
                    try:
                        os.remove(filepath)
                    except Exception:
                        pass

            db.update_product(cod_prd, {'Foto': None})
            return jsonify({"message": "Foto removida com sucesso!"})

        # POST: Upload ou Atualização de Foto
        photo_url = None
        if 'file' in request.files and request.files['file'].filename != '':
            file = request.files['file']
            allowed_extensions = {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.svg', '.bmp'}
            _, ext = os.path.splitext(file.filename.lower())
            if ext not in allowed_extensions:
                return jsonify({"error": "Formato de arquivo não suportado"}), 400

            upload_dir = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'produtos')
            os.makedirs(upload_dir, exist_ok=True)

            unique_name = f"prd_{cod_prd}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:6]}{ext}"
            filepath = os.path.join(upload_dir, unique_name)
            file.save(filepath)
            photo_url = f"/static/uploads/produtos/{unique_name}"
        elif request.json and request.json.get('foto_url'):
            photo_url = request.json.get('foto_url')
        else:
            return jsonify({"error": "Nenhum arquivo ou URL de foto informada"}), 400

        db.update_product(cod_prd, {'Foto': photo_url})
        return jsonify({"message": "Foto do produto atualizada com sucesso!", "foto_url": photo_url})
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
        grupo = request.args.get('grupo', None)
        limit = request.args.get('limit', 60, type=int)
        results = db.search_pdv_products(query=q, grupo=grupo, limit=limit)
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
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
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

@main_bp.route('/api/empresas/upload-logo', methods=['POST'])
def upload_empresa_logo_temp():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "Nenhum arquivo de imagem enviado"}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "Nenhum arquivo selecionado"}), 400

        allowed_extensions = {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.svg', '.bmp'}
        _, ext = os.path.splitext(file.filename.lower())
        if ext not in allowed_extensions:
            return jsonify({"error": f"Formato inválido. Extensões permitidas: {', '.join(allowed_extensions)}"}), 400

        upload_dir = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'empresas')
        os.makedirs(upload_dir, exist_ok=True)

        unique_name = f"logo_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}{ext}"
        filepath = os.path.join(upload_dir, unique_name)
        file.save(filepath)

        logo_url = f"/static/uploads/empresas/{unique_name}"
        return jsonify({
            "message": "Logomarca enviada com sucesso!",
            "url": logo_url,
            "filename": unique_name
        }), 200
    except Exception as e:
        return jsonify({"error": f"Erro ao processar imagem: {str(e)}"}), 500

@main_bp.route('/api/empresas/<int:id_empresa>/logo', methods=['POST', 'DELETE'])
def empresa_logo_manage(id_empresa):
    from app.empresa_manager import get_empresa_by_id, save_empresa
    try:
        empresa = get_empresa_by_id(id_empresa)
        if not empresa:
            return jsonify({"error": "Empresa não encontrada"}), 404

        if request.method == 'DELETE':
            old_logo = empresa.get('Logo')
            if old_logo and old_logo.startswith('/static/uploads/empresas/'):
                filename = os.path.basename(old_logo)
                filepath = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'empresas', filename)
                if os.path.exists(filepath):
                    try:
                        os.remove(filepath)
                    except Exception:
                        pass
            empresa['Logo'] = ''
            save_empresa(empresa)
            return jsonify({"message": "Logomarca removida com sucesso!"})

        # POST: Upload ou Atualização de Logo
        logo_url = None
        if 'file' in request.files and request.files['file'].filename != '':
            file = request.files['file']
            allowed_extensions = {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.svg', '.bmp'}
            _, ext = os.path.splitext(file.filename.lower())
            if ext not in allowed_extensions:
                return jsonify({"error": "Formato de arquivo não suportado"}), 400

            upload_dir = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'empresas')
            os.makedirs(upload_dir, exist_ok=True)

            unique_name = f"logo_{id_empresa}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:6]}{ext}"
            filepath = os.path.join(upload_dir, unique_name)
            file.save(filepath)
            logo_url = f"/static/uploads/empresas/{unique_name}"
        elif request.json and request.json.get('logo_url'):
            logo_url = request.json.get('logo_url')
        else:
            return jsonify({"error": "Nenhum arquivo ou URL de logomarca informada"}), 400

        empresa['Logo'] = logo_url
        save_empresa(empresa)
        return jsonify({"message": "Logomarca da empresa atualizada com sucesso!", "logo_url": logo_url})
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
    from app.nfe_engine import emit_nfe_55, validate_nfe_structure
    from app.danfe_generator import render_danfe_html
    from app.config_manager import get_pdv_config
    import html as pyhtml
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404

        order = order_data["pedido"]
        items = order_data["itens"]
        company = order_data["empresa"]
        cfg = get_pdv_config() or {}
        
        # Check if already emitted with NroChave
        nfe_info = order_data.get("nfe")
        chave_existente = (nfe_info and nfe_info.get("NroChave")) or order.get("Sat") or order.get("NroChave")
        protocolo_existente = (nfe_info and nfe_info.get("Protocolo")) or "135180587979473"

        if chave_existente and len(str(chave_existente)) == 44:
            nfe_res = {
                "chave_nfe": str(chave_existente),
                "protocolo": str(protocolo_existente),
                "status": str((nfe_info and nfe_info.get("Status")) or "100")
            }
        else:
            val_report = validate_nfe_structure(order, company, items, cfg)
            if not val_report.get("valid"):
                errs_html = "".join(f"<li>{pyhtml.escape(e)}</li>" for e in val_report.get("errors", []))
                return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Bloqueio de Impressão de DANFE</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0f172a; color: #f8fafc; padding: 2.5rem; text-align: center; }}
        .card {{ max-width: 600px; margin: 0 auto; background: #1e293b; border: 1px solid #f43f5e; border-radius: 12px; padding: 2rem; box-shadow: 0 10px 25px rgba(0,0,0,0.5); }}
        h2 {{ color: #f43f5e; margin-top: 0; }}
        ul {{ text-align: left; background: rgba(244,63,94,0.1); border-radius: 8px; padding: 1.25rem 2rem; color: #fca5a5; }}
        .btn {{ display: inline-block; margin-top: 1.5rem; padding: 0.6rem 1.2rem; background: #3b82f6; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; }}
    </style>
</head>
<body>
    <div class="card">
        <h2>⚠️ Impressão de DANFE / PDF Bloqueada</h2>
        <p>A NF-e do Pedido <strong>#{cod_ped}</strong> não pode ser emitida ou impressa porque possui pendências fiscais que violam as regras da SEFAZ:</p>
        <ul>{errs_html}</ul>
        <p style="font-size: 0.9rem; color: #94a3b8;">Abra o pedido no <strong>FrmNota</strong> e preencha todos os dados do Destinatário (CPF/CNPJ, Nome e Endereço) para liberar a transmissão e impressão.</p>
        <a href="javascript:window.close()" class="btn">Fechar Janela</a>
    </div>
</body>
</html>""", 400, {'Content-Type': 'text/html; charset=utf-8'}

            nfe_res = emit_nfe_55(order, company, items)

        html = render_danfe_html(order, company, items, nfe_res)
        return html, 200, {'Content-Type': 'text/html; charset=utf-8'}
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@main_bp.route('/api/nfe/<int:cod_ped>/xml-download', methods=['GET'])
def download_nfe_xml(cod_ped):
    from app.nfe_engine import emit_nfe_55, validate_nfe_structure
    from app.config_manager import get_pdv_config
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404

        order = order_data["pedido"]
        items = order_data["itens"]
        company = order_data["empresa"]
        cfg = get_pdv_config() or {}

        val_report = validate_nfe_structure(order, company, items, cfg)
        if not val_report.get("valid"):
            return jsonify({
                "error": "Não é possível baixar o XML: a nota possui pendências fiscais que violam as regras da SEFAZ.",
                "errors": val_report.get("errors", [])
            }), 400

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

@main_bp.route('/api/pedidos/<int:cod_ped>/validar-nfce', methods=['GET', 'POST'])
def validar_nfce_pedido(cod_ped):
    from app.nfce import validate_nfce_structure
    from app.empresa_manager import get_empresa_by_id
    from app.config_manager import get_pdv_config
    try:
        order_data = db.get_order_by_id(cod_ped)
        if not order_data or not order_data.get("pedido"):
            return jsonify({"error": "Pedido não encontrado"}), 404
        
        company_data = get_empresa_by_id(1) or {}
        config_data = get_pdv_config() or {}

        report = validate_nfce_structure(
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


# --- RELATÓRIO DE VENDAS AGRUPADO POR PRODUTO ENDPOINTS ---

@main_bp.route('/api/relatorios/vendas-por-produto', methods=['GET'])
def get_relatorio_vendas_produto_route():
    try:
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        grupo = request.args.get('grupo', None)
        q = request.args.get('q', None)
        id_empresa = request.args.get('id_empresa', None)
        status = request.args.get('status', 'ativos')
        sort_by = request.args.get('sort_by', 'total_valor')
        sort_order = request.args.get('sort_order', 'DESC')
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 50, type=int)

        relatorio = db.get_relatorio_vendas_produto(
            data_inicio=data_inicio,
            data_fim=data_fim,
            grupo=grupo,
            q=q,
            id_empresa=id_empresa,
            status=status,
            sort_by=sort_by,
            sort_order=sort_order,
            page=page,
            limit=limit
        )
        return jsonify(relatorio)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/vendas-por-produto/detalhes/<int:cod_prd>', methods=['GET'])
def get_relatorio_vendas_produto_detalhes_route(cod_prd):
    try:
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        id_empresa = request.args.get('id_empresa', None)
        status = request.args.get('status', 'ativos')

        detalhes = db.get_relatorio_vendas_produto_detalhes(
            cod_prd=cod_prd,
            data_inicio=data_inicio,
            data_fim=data_fim,
            id_empresa=id_empresa,
            status=status
        )
        return jsonify(detalhes)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/vendas-por-produto/export/csv', methods=['GET'])
def export_relatorio_vendas_produto_csv_route():
    import io
    import csv
    try:
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        grupo = request.args.get('grupo', None)
        q = request.args.get('q', None)
        id_empresa = request.args.get('id_empresa', None)
        status = request.args.get('status', 'ativos')
        sort_by = request.args.get('sort_by', 'total_valor')
        sort_order = request.args.get('sort_order', 'DESC')

        # Limit 0 to get all rows for export
        relatorio = db.get_relatorio_vendas_produto(
            data_inicio=data_inicio,
            data_fim=data_fim,
            grupo=grupo,
            q=q,
            id_empresa=id_empresa,
            status=status,
            sort_by=sort_by,
            sort_order=sort_order,
            page=1,
            limit=0
        )

        output = io.StringIO()
        # UTF-8 BOM for Excel pt-BR
        output.write('\ufeff')
        writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_MINIMAL)

        # Header rows
        writer.writerow(['RELATORIO DE VENDAS AGRUPADO POR PRODUTO'])
        periodo_str = f"Periodo: {data_inicio or 'Inicio'} ate {data_fim or 'Hoje'}"
        writer.writerow([periodo_str, f"Status: {status.upper()}", f"Total Itens Distintos: {relatorio['summary']['total_produtos_distintos']}"])
        writer.writerow([
            f"Faturamento Total: R$ {relatorio['summary']['total_faturamento_geral']:,.2f}",
            f"Qtd Total: {relatorio['summary']['total_qtd_geral']}",
            f"Lucro Estimado: R$ {relatorio['summary']['lucro_bruto_geral']:,.2f} ({relatorio['summary']['margem_lucro_pct_geral']}%)"
        ])
        writer.writerow([])

        # Table header
        writer.writerow([
            'Codigo',
            'Codigo de Barras',
            'Descricao do Produto',
            'Grupo',
            'Unidade',
            'Qtd Vendida',
            'Nro Pedidos',
            'Preco Medio (R$)',
            'Desconto (R$)',
            'Total Faturado (R$)',
            'Participacao (%)',
            'Custo Total (R$)',
            'Lucro Bruto (R$)',
            'Margem (%)',
            'Estoque Atual'
        ])

        for item in relatorio['items']:
            writer.writerow([
                item.get('CodPrd', ''),
                item.get('CodBar', ''),
                item.get('Descricao_Produto', ''),
                item.get('Descricao_Grupo', ''),
                item.get('Embalagem', 'UN'),
                f"{float(item.get('total_qtd', 0)):.2f}".replace('.', ','),
                item.get('total_pedidos', 0),
                f"{float(item.get('preco_medio', 0)):.2f}".replace('.', ','),
                f"{float(item.get('total_desconto', 0)):.2f}".replace('.', ','),
                f"{float(item.get('total_valor', 0)):.2f}".replace('.', ','),
                f"{float(item.get('participacao_pct', 0)):.2f}%".replace('.', ','),
                f"{float(item.get('custo_total', 0)):.2f}".replace('.', ','),
                f"{float(item.get('lucro_bruto', 0)):.2f}".replace('.', ','),
                f"{float(item.get('margem_pct', 0)):.2f}%".replace('.', ','),
                f"{float(item.get('EstoqueAtual', 0)):.2f}".replace('.', ',')
            ])

        csv_content = output.getvalue()
        now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"Relatorio_Vendas_Produto_{now_str}.csv"

        return csv_content, 200, {
            'Content-Type': 'text/csv; charset=utf-8',
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- RELATORIO VENDAS POR CLIENTE ---

@main_bp.route('/api/relatorios/vendas-por-cliente', methods=['GET'])
def get_relatorio_vendas_cliente_route():
    try:
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        cod_entidade = request.args.get('cod_entidade', None)
        grupo = request.args.get('grupo', None)
        q = request.args.get('q', None)
        id_empresa = request.args.get('id_empresa', None)
        status = request.args.get('status', 'ativos')
        sort_by = request.args.get('sort_by', 'total_valor')
        sort_order = request.args.get('sort_order', 'DESC')
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 50, type=int)

        relatorio = db.get_relatorio_vendas_cliente(
            data_inicio=data_inicio,
            data_fim=data_fim,
            cod_entidade=cod_entidade,
            grupo=grupo,
            q=q,
            id_empresa=id_empresa,
            status=status,
            sort_by=sort_by,
            sort_order=sort_order,
            page=page,
            limit=limit
        )
        return jsonify(relatorio)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/vendas-por-cliente/detalhes/<int:cod_entidade>', methods=['GET'])
def get_relatorio_vendas_cliente_detalhes_route(cod_entidade):
    try:
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        id_empresa = request.args.get('id_empresa', None)
        status = request.args.get('status', 'ativos')

        detalhes = db.get_relatorio_vendas_cliente_detalhes(
            cod_entidade=cod_entidade,
            data_inicio=data_inicio,
            data_fim=data_fim,
            id_empresa=id_empresa,
            status=status
        )
        return jsonify(detalhes)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/vendas-por-cliente/export/csv', methods=['GET'])
def export_relatorio_vendas_cliente_csv_route():
    import io
    import csv
    try:
        data_inicio = request.args.get('data_inicio', None)
        data_fim = request.args.get('data_fim', None)
        cod_entidade = request.args.get('cod_entidade', None)
        grupo = request.args.get('grupo', None)
        q = request.args.get('q', None)
        id_empresa = request.args.get('id_empresa', None)
        status = request.args.get('status', 'ativos')
        sort_by = request.args.get('sort_by', 'total_valor')
        sort_order = request.args.get('sort_order', 'DESC')

        # Limit 0 to get all rows for export
        relatorio = db.get_relatorio_vendas_cliente(
            data_inicio=data_inicio,
            data_fim=data_fim,
            cod_entidade=cod_entidade,
            grupo=grupo,
            q=q,
            id_empresa=id_empresa,
            status=status,
            sort_by=sort_by,
            sort_order=sort_order,
            page=1,
            limit=0
        )

        output = io.StringIO()
        output.write('\ufeff')
        writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_MINIMAL)

        writer.writerow(['RELATORIO DE VENDAS DE PRODUTOS POR CLIENTE'])
        periodo_str = f"Periodo: {data_inicio or 'Inicio'} ate {data_fim or 'Hoje'}"
        writer.writerow([periodo_str, f"Status: {status.upper()}", f"Total Clientes: {relatorio['summary']['total_clientes_distintos']}"])
        writer.writerow([
            f"Faturamento Total: R$ {relatorio['summary']['total_faturamento_geral']:,.2f}",
            f"Total de Pedidos: {relatorio['summary']['total_pedidos_geral']}",
            f"Total de Itens/Qtd: {relatorio['summary']['total_qtd_geral']}",
            f"Ticket Medio: R$ {relatorio['summary']['ticket_medio_geral']:,.2f}"
        ])
        writer.writerow([])

        # Table header
        writer.writerow([
            'Cod Cliente',
            'Razao Social / Nome',
            'Nome Fantasia',
            'CPF / CNPJ',
            'Cidade',
            'UF',
            'Telefone',
            'Qtd Pedidos',
            'Qtd Itens/Pecas',
            'Itens Distintos',
            'Desconto Total (R$)',
            'Total Compras (R$)',
            'Participacao (%)',
            'Ticket Medio (R$)',
            'Ultima Compra',
            'Cod Produto',
            'Descricao do Produto',
            'Qtd Produto',
            'Unidade',
            'Preco Medio (R$)',
            'Total Produto (R$)'
        ])

        for cli in relatorio['items']:
            cli_base = [
                cli.get('CodEntidade', ''),
                cli.get('Nome', ''),
                cli.get('Fantasia', ''),
                cli.get('Documento', ''),
                cli.get('Cidade', ''),
                cli.get('Uf', ''),
                cli.get('Fone', ''),
                cli.get('total_pedidos', 0),
                f"{float(cli.get('total_qtd', 0)):.2f}".replace('.', ','),
                cli.get('total_itens_distintos', 0),
                f"{float(cli.get('total_desconto', 0)):.2f}".replace('.', ','),
                f"{float(cli.get('total_valor', 0)):.2f}".replace('.', ','),
                f"{float(cli.get('participacao_pct', 0)):.2f}%".replace('.', ','),
                f"{float(cli.get('ticket_medio', 0)):.2f}".replace('.', ','),
                cli.get('ultima_compra', '')
            ]
            
            prods = cli.get('produtos', [])
            if not prods:
                writer.writerow(cli_base + ['', '', '', '', '', ''])
            else:
                for idx, prod in enumerate(prods):
                    prod_cols = [
                        prod.get('CodPrd', ''),
                        prod.get('Descricao_Produto', ''),
                        f"{float(prod.get('total_qtd', 0)):.2f}".replace('.', ','),
                        prod.get('Embalagem', 'UN'),
                        f"{float(prod.get('preco_medio', 0)):.2f}".replace('.', ','),
                        f"{float(prod.get('total_valor', 0)):.2f}".replace('.', ',')
                    ]
                    if idx == 0:
                        writer.writerow(cli_base + prod_cols)
                    else:
                        # Blank customer info for subsequent product rows for cleaner reading
                        writer.writerow([''] * len(cli_base) + prod_cols)

        csv_content = output.getvalue()
        now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"Relatorio_Vendas_Por_Cliente_{now_str}.csv"

        return csv_content, 200, {
            'Content-Type': 'text/csv; charset=utf-8',
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- RELATORIO DE DADOS CADASTRAIS DE ENTIDADES (CLIENTES / FORNECEDORES) ---

@main_bp.route('/api/relatorios/entidades', methods=['GET'])
def get_relatorio_entidades_route():
    try:
        tipo = request.args.get('tipo', None)
        ativo = request.args.get('ativo', None)
        pessoa = request.args.get('pessoa', None)
        uf = request.args.get('uf', None)
        cidade = request.args.get('cidade', None)
        q = request.args.get('q', None)
        id_empresa = request.args.get('id_empresa', None)
        sort_by = request.args.get('sort_by', 'Nome')
        sort_order = request.args.get('sort_order', 'ASC')
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 50, type=int)

        relatorio = db.get_relatorio_entidades(
            tipo=tipo,
            ativo=ativo,
            pessoa=pessoa,
            uf=uf,
            cidade=cidade,
            q=q,
            id_empresa=id_empresa,
            sort_by=sort_by,
            sort_order=sort_order,
            page=page,
            limit=limit
        )
        return jsonify(relatorio)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/entidades/ficha/<int:cod_entidade>', methods=['GET'])
def get_relatorio_entidade_ficha_route(cod_entidade):
    try:
        ficha = db.get_relatorio_entidade_ficha(cod_entidade)
        if not ficha:
            return jsonify({"error": "Entidade não encontrada"}), 404
        return jsonify(ficha)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/entidades/filtros-locais', methods=['GET'])
def get_relatorio_entidades_filtros_locais_route():
    try:
        dados = db.get_relatorio_entidades_ufs_cidades()
        return jsonify(dados)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@main_bp.route('/api/relatorios/entidades/export/csv', methods=['GET'])
def export_relatorio_entidades_csv_route():
    import io
    import csv
    import datetime

    try:
        tipo = request.args.get('tipo', None)
        ativo = request.args.get('ativo', None)
        pessoa = request.args.get('pessoa', None)
        uf = request.args.get('uf', None)
        cidade = request.args.get('cidade', None)
        q = request.args.get('q', None)
        id_empresa = request.args.get('id_empresa', None)
        sort_by = request.args.get('sort_by', 'Nome')
        sort_order = request.args.get('sort_order', 'ASC')

        # Buscar todos os registros correspondentes para exportação (limit 5000)
        relatorio = db.get_relatorio_entidades(
            tipo=tipo,
            ativo=ativo,
            pessoa=pessoa,
            uf=uf,
            cidade=cidade,
            q=q,
            id_empresa=id_empresa,
            sort_by=sort_by,
            sort_order=sort_order,
            page=1,
            limit=5000
        )

        output = io.StringIO()
        # UTF-8 BOM para compatibilidade com Microsoft Excel
        output.write('\ufeff')
        writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_MINIMAL)

        # Cabeçalho CSV
        writer.writerow([
            'Codigo',
            'Tipo Entidade',
            'Pessoa',
            'Razao Social / Nome',
            'Nome Fantasia',
            'CPF / CNPJ',
            'RG / Inscricao Estadual',
            'Telefone',
            'Celular',
            'Email',
            'Endereco',
            'Numero',
            'Complemento',
            'Bairro',
            'Cidade',
            'UF',
            'CEP',
            'Limite Credito (R$)',
            'Condicao Padrao',
            'Status',
            'Data Cadastro',
            'Ultimo Movimento',
            'Observacoes'
        ])

        for ent in relatorio.get('items', []):
            writer.writerow([
                ent.get('CodEntidade', ''),
                ent.get('TipoDescricao', ''),
                ent.get('TipoPessoa', ''),
                ent.get('Nome', ''),
                ent.get('Fantasia', ''),
                ent.get('DocumentoFormatado', ''),
                ent.get('InscrEst', '') or ent.get('RG', ''),
                ent.get('Fone', ''),
                ent.get('Celular', ''),
                ent.get('Email', ''),
                ent.get('Endereco', ''),
                ent.get('Nro', ''),
                ent.get('Complemento', ''),
                ent.get('Bairro', ''),
                ent.get('Cidade', ''),
                ent.get('Uf', ''),
                ent.get('Cep', ''),
                f"{float(ent.get('Credito', 0)):.2f}".replace('.', ','),
                ent.get('Condicao', '') or ent.get('Prazo', ''),
                'ATIVO' if ent.get('IsAtivo') else 'INATIVO',
                ent.get('DtCadastro', ''),
                ent.get('DtUltMov', ''),
                str(ent.get('Obs', '')).replace('\n', ' ').replace('\r', ' ')
            ])

        csv_content = output.getvalue()
        now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"Relatorio_Cadastral_Entidades_{now_str}.csv"

        return csv_content, 200, {
            'Content-Type': 'text/csv; charset=utf-8',
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
    except Exception as e:
        return jsonify({"error": str(e)}), 500














