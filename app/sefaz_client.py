import os
import ssl
import time
import tempfile
import urllib.request
import urllib.error
from lxml import etree
from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption
from app.nfe_signer import load_pfx_certificate, sign_xml_sefaz

SEFAZ_URLS = {
    "SP": {
        "55": {
            1: { # Produção
                "StatusServico": "https://nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx",
                "Autorizacao": "https://nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx",
                "RetAutorizacao": "https://nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx",
                "RecepcaoEvento": "https://nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx",
                "Inutilizacao": "https://nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx",
                "ConsultaProtocolo": "https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
            },
            2: { # Homologação
                "StatusServico": "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx",
                "Autorizacao": "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx",
                "RetAutorizacao": "https://homologacao.nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx",
                "RecepcaoEvento": "https://homologacao.nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx",
                "Inutilizacao": "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx",
                "ConsultaProtocolo": "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
            }
        },
        "65": { # NFC-e SP
            1: {
                "StatusServico": "https://nfce.fazenda.sp.gov.br/ws/nfestatusservico4.asmx",
                "Autorizacao": "https://nfce.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx",
                "RetAutorizacao": "https://nfce.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx",
                "RecepcaoEvento": "https://nfce.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx",
                "Inutilizacao": "https://nfce.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
            },
            2: {
                "StatusServico": "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfestatusservico4.asmx",
                "Autorizacao": "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx",
                "RetAutorizacao": "https://homologacao.nfce.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx",
                "RecepcaoEvento": "https://homologacao.nfce.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx",
                "Inutilizacao": "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
            }
        }
    },
    "SVRS": { # Sefaz Virtual RS (utilizado pela maioria dos estados)
        "55": {
            1: {
                "StatusServico": "https://nfe.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx",
                "Autorizacao": "https://nfe.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx",
                "RetAutorizacao": "https://nfe.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx",
                "RecepcaoEvento": "https://nfe.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx",
                "Inutilizacao": "https://nfe.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
            },
            2: {
                "StatusServico": "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx",
                "Autorizacao": "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx",
                "RetAutorizacao": "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx",
                "RecepcaoEvento": "https://nfe-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx",
                "Inutilizacao": "https://nfe-homologacao.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
            }
        },
        "65": {
            1: {
                "StatusServico": "https://nfce.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx",
                "Autorizacao": "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx",
                "RetAutorizacao": "https://nfce.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx",
                "RecepcaoEvento": "https://nfce.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
            },
            2: {
                "StatusServico": "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx",
                "Autorizacao": "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx",
                "RetAutorizacao": "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx",
                "RecepcaoEvento": "https://nfce-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
            }
        }
    }
}

CODIGOS_UF = {
    "AC": "12", "AL": "27", "AP": "16", "AM": "13", "BA": "29", "CE": "23",
    "DF": "53", "ES": "32", "GO": "52", "MA": "21", "MT": "51", "MS": "50",
    "MG": "31", "PA": "15", "PB": "25", "PR": "41", "PE": "26", "PI": "22",
    "RJ": "33", "RN": "24", "RS": "43", "RO": "11", "RR": "14", "SC": "42",
    "SP": "35", "SE": "28", "TO": "17"
}


def get_sefaz_endpoint(service_name, uf="SP", modelo="55", ambiente=2):
    """
    Retorna a URL do WebService da SEFAZ para o serviço, UF, modelo e ambiente solicitados.
    """
    uf_upper = str(uf or "SP").upper()
    mod_str = str(modelo or "55")
    amb_int = int(ambiente) if str(ambiente) in ("1", "2") else 2

    uf_group = SEFAZ_URLS.get(uf_upper, SEFAZ_URLS["SVRS"])
    mod_group = uf_group.get(mod_str, uf_group.get("55"))
    amb_group = mod_group.get(amb_int, mod_group.get(2))

    url = amb_group.get(service_name)
    if not url:
        # Fallback para SVRS
        url = SEFAZ_URLS["SVRS"][mod_str][amb_int].get(service_name)
    return url


def create_sefaz_ssl_context(pfx_path=None, password=None):
    """
    Cria um contexto SSL mTLS configurado com o Certificado Digital A1 do cliente.
    """
    ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    if pfx_path and os.path.exists(pfx_path):
        with open(pfx_path, 'rb') as f:
            pfx_bytes = f.read()

        pwd = password.encode('utf-8') if password else None
        private_key, certificate, additional_certs = pkcs12.load_key_and_certificates(pfx_bytes, pwd)

        if private_key and certificate:
            # Gerar arquivo PEM temporário para configurar mTLS na conexão SSL
            cert_pem = certificate.public_bytes(Encoding.PEM)
            key_pem = private_key.private_bytes(
                Encoding.PEM,
                PrivateFormat.PKCS8,
                NoEncryption()
            )

            chain_pem = b"".join(c.public_bytes(Encoding.PEM) for c in (additional_certs or []))

            # Cria arquivos temporários seguros para o SSLContext
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pem") as cert_file, \
                 tempfile.NamedTemporaryFile(delete=False, suffix=".pem") as key_file:
                cert_file.write(cert_pem + chain_pem)
                key_file.write(key_pem)
                cert_file_path = cert_file.name
                key_file_path = key_file.name

            try:
                ctx.load_cert_chain(certfile=cert_file_path, keyfile=key_file_path)
            finally:
                try:
                    os.remove(cert_file_path)
                    os.remove(key_file_path)
                except Exception:
                    pass

    return ctx


def send_sefaz_soap_request(url, soap_action, xml_body, pfx_path=None, password=None, timeout=15):
    """
    Executa a requisição HTTP POST SOAP 1.2 com mTLS para o WebService SEFAZ.
    """
    soap_envelope = f"""<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <nfeDadosMsg xmlns="{soap_action}">
      {xml_body}
    </nfeDadosMsg>
  </soap12:Body>
</soap12:Envelope>"""

    headers = {
        "Content-Type": "application/soap+xml; charset=utf-8",
        "User-Agent": "SistemaVendas/1.0 (Windows NT 10.0; Win64; x64)"
    }

    ssl_ctx = create_sefaz_ssl_context(pfx_path, password)
    req = urllib.request.Request(url, data=soap_envelope.encode('utf-8'), headers=headers, method="POST")

    start_time = time.time()
    try:
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=timeout) as response:
            resp_bytes = response.read()
            elapsed_ms = int((time.time() - start_time) * 1000)
            return {
                "success": True,
                "status_code": response.getcode(),
                "response_xml": resp_bytes.decode('utf-8', errors='replace'),
                "elapsed_ms": elapsed_ms
            }
    except urllib.error.HTTPError as e:
        err_bytes = e.read() if hasattr(e, 'read') else b""
        elapsed_ms = int((time.time() - start_time) * 1000)
        return {
            "success": False,
            "status_code": e.code,
            "error": f"Erro HTTP {e.code}: {e.reason}",
            "response_xml": err_bytes.decode('utf-8', errors='replace'),
            "elapsed_ms": elapsed_ms
        }
    except Exception as e:
        elapsed_ms = int((time.time() - start_time) * 1000)
        return {
            "success": False,
            "status_code": 0,
            "error": f"Falha de conexão com a SEFAZ: {str(e)}",
            "response_xml": "",
            "elapsed_ms": elapsed_ms
        }


def parse_sefaz_retorno_autorizacao(resp_xml_str):
    """
    Faz o parse do XML de retorno da SEFAZ (<retEnviNFe> / <protNFe>).
    """
    try:
        parser = etree.XMLParser(recover=True)
        root = etree.fromstring(resp_xml_str.encode('utf-8'), parser=parser)

        # Buscar protNFe
        prot = root.find(".//{http://www.portalfiscal.inf.br/nfe}protNFe")
        if prot is None:
            prot = root.find(".//protNFe")

        if prot is not None:
            inf_prot = prot.find(".//{http://www.portalfiscal.inf.br/nfe}infProt")
            if inf_prot is None:
                inf_prot = prot.find(".//infProt")

            if inf_prot is not None:
                c_stat = inf_prot.findtext(".//{http://www.portalfiscal.inf.br/nfe}cStat") or inf_prot.findtext(".//cStat") or ""
                x_motivo = inf_prot.findtext(".//{http://www.portalfiscal.inf.br/nfe}xMotivo") or inf_prot.findtext(".//xMotivo") or ""
                ch_nfe = inf_prot.findtext(".//{http://www.portalfiscal.inf.br/nfe}chNFe") or inf_prot.findtext(".//chNFe") or ""
                n_prot = inf_prot.findtext(".//{http://www.portalfiscal.inf.br/nfe}nProt") or inf_prot.findtext(".//nProt") or ""
                dh_recbto = inf_prot.findtext(".//{http://www.portalfiscal.inf.br/nfe}dhRecbto") or inf_prot.findtext(".//dhRecbto") or ""
                dig_val = inf_prot.findtext(".//{http://www.portalfiscal.inf.br/nfe}digVal") or inf_prot.findtext(".//digVal") or ""

                prot_xml_str = etree.tostring(prot, encoding="utf-8").decode('utf-8')

                return {
                    "cStat": c_stat,
                    "xMotivo": x_motivo,
                    "chNFe": ch_nfe,
                    "nProt": n_prot,
                    "dhRecbto": dh_recbto,
                    "digVal": dig_val,
                    "prot_xml": prot_xml_str,
                    "autorizada": c_stat in ("100", "150")
                }

        # Buscar retEnviNFe geral (cStat de lote)
        ret_envi = root.find(".//{http://www.portalfiscal.inf.br/nfe}retEnviNFe")
        if ret_envi is None:
            ret_envi = root.find(".//retEnviNFe")
        if ret_envi is not None:
            c_stat = ret_envi.findtext(".//{http://www.portalfiscal.inf.br/nfe}cStat") or ret_envi.findtext(".//cStat") or ""
            x_motivo = ret_envi.findtext(".//{http://www.portalfiscal.inf.br/nfe}xMotivo") or ret_envi.findtext(".//xMotivo") or ""
            n_rec = ret_envi.findtext(".//{http://www.portalfiscal.inf.br/nfe}infRec/nRec") or ret_envi.findtext(".//nRec") or ""

            return {
                "cStat": c_stat,
                "xMotivo": x_motivo,
                "nRec": n_rec,
                "autorizada": False,
                "prot_xml": ""
            }

        return {
            "cStat": "999",
            "xMotivo": "Formato de retorno SEFAZ não reconhecido",
            "autorizada": False,
            "prot_xml": ""
        }

    except Exception as e:
        return {
            "cStat": "999",
            "xMotivo": f"Erro ao analisar XML de resposta SEFAZ: {str(e)}",
            "autorizada": False,
            "prot_xml": ""
        }


def build_envi_nfe_batch(signed_nfe_xml_str, id_lote=1, ind_sinc=1):
    """
    Envelopa a NF-e / NFC-e assinada no lote oficial <enviNFe versao="4.00">.
    """
    # Limpar declaração XML interna se houver
    clean_nfe = signed_nfe_xml_str
    if clean_nfe.startswith("<?xml"):
        clean_nfe = clean_nfe.split("?>", 1)[-1].strip()

    envi_xml = f"""<enviNFe xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
  <idLote>{id_lote}</idLote>
  <indSinc>{ind_sinc}</indSinc>
  {clean_nfe}
</enviNFe>"""
    return envi_xml


def build_nfeproc_authorized(signed_nfe_xml_str, prot_nfe_xml_str):
    """
    Gera o XML final de distribuição <nfeProc versao="4.00"> (NFe assinada + Protocolo de Autorização).
    """
    clean_nfe = signed_nfe_xml_str.split("?>", 1)[-1].strip() if signed_nfe_xml_str.startswith("<?xml") else signed_nfe_xml_str.strip()
    clean_prot = prot_nfe_xml_str.split("?>", 1)[-1].strip() if prot_nfe_xml_str.startswith("<?xml") else prot_nfe_xml_str.strip()

    proc_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
  {clean_nfe}
  {clean_prot}
</nfeProc>"""
    return proc_xml
