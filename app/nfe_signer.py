import base64
import os
import io
from lxml import etree
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import pkcs12, Encoding
from cryptography import x509

XMLDSIG_NS = "http://www.w3.org/2000/09/xmldsig#"
C14N_ALGO = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315"
ENVELOPED_SIG = "http://www.w3.org/2000/09/xmldsig#enveloped-signature"
RSA_SHA1_ALGO = "http://www.w3.org/2000/09/xmldsig#rsa-sha1"
SHA1_ALGO = "http://www.w3.org/2000/09/xmldsig#sha1"

class CertificateError(Exception):
    pass

class SignatureError(Exception):
    pass


def load_pfx_certificate(pfx_path, password=None):
    """
    Carrega o certificado A1 (.pfx / .p12) e retorna:
    (private_key, certificate, cert_chain, cert_der_b64)
    """
    if not os.path.exists(pfx_path):
        raise CertificateError(f"Arquivo de certificado digital não encontrado: '{pfx_path}'")

    with open(pfx_path, 'rb') as f:
        pfx_data = f.read()

    pwd_bytes = password.encode('utf-8') if password else None

    try:
        private_key, certificate, additional_certs = pkcs12.load_key_and_certificates(
            pfx_data,
            pwd_bytes
        )
    except Exception as e:
        raise CertificateError(f"Falha ao abrir certificado digital A1. Verifique a senha informada. Erro: {str(e)}")

    if not private_key:
        raise CertificateError("Chave privada RSA não encontrada no certificado digital.")
    if not certificate:
        raise CertificateError("Certificado X509 não encontrado no arquivo PFX.")

    cert_der = certificate.public_bytes(Encoding.DER)
    cert_b64 = base64.b64encode(cert_der).decode('ascii')

    return {
        "private_key": private_key,
        "certificate": certificate,
        "additional_certs": additional_certs or [],
        "cert_der": cert_der,
        "cert_b64": cert_b64,
        "subject": certificate.subject.rfc4514_string(),
        "not_valid_after": certificate.not_valid_after_utc if hasattr(certificate, 'not_valid_after_utc') else certificate.not_valid_after,
        "not_valid_before": certificate.not_valid_before_utc if hasattr(certificate, 'not_valid_before_utc') else certificate.not_valid_before
    }


def get_configured_or_fallback_certificate(id_empresa=1, pfx_path=None, password=None):
    """
    Busca o certificado configurado na tabela CfgPdv ou no caminho informado.
    Caso não exista no caminho informado, faz fallback transparente para o certificado de homologação certs/demo_a1.pfx.
    """
    if pfx_path and os.path.exists(pfx_path):
        return load_pfx_certificate(pfx_path, password)

    try:
        from app.config_manager import get_pdv_config
        cfg = get_pdv_config(id_empresa)
        cfg_cert = cfg.get("Certificado", {})
        cfg_path = cfg_cert.get("Caminho", "")
        cfg_pwd = cfg_cert.get("Senha", "")
        if cfg_path and os.path.exists(cfg_path):
            return load_pfx_certificate(cfg_path, cfg_pwd)
    except Exception as e:
        pass

    # Fallback para certs/demo_a1.pfx
    demo_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "certs", "demo_a1.pfx")
    if os.path.exists(demo_path):
        return load_pfx_certificate(demo_path, "123456")

def sign_xml_sefaz(xml_content, pfx_path=None, password=None, cert_info=None, id_empresa=1):
    """
    Assina digitalmente um XML de NF-e, NFC-e ou Evento SEFAZ seguindo rigorosamente
    o padrão XML-DSig / Manual de Orientação do Contribuinte (MOC 7.00).

    Tags assinadas: <infNFe Id="NFe...">, <infNFCe Id="NFe..."> ou <infEvento Id="ID...">.
    """
    if cert_info is None:
        cert_info = get_configured_or_fallback_certificate(id_empresa=id_empresa, pfx_path=pfx_path, password=password)

    private_key = cert_info["private_key"]
    cert_b64 = cert_info["cert_b64"]


    # Parse XML com lxml
    if isinstance(xml_content, str):
        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(xml_content.encode('utf-8'), parser=parser)
    else:
        root = xml_content

    # Localizar nó a ser assinado
    inf_node = None
    ref_uri = None

    # Verificar se é NFe / NFC-e (<infNFe>)
    for tag_name in ["{http://www.portalfiscal.inf.br/nfe}infNFe", "infNFe", 
                     "{http://www.portalfiscal.inf.br/nfe}infEvento", "infEvento",
                     "{http://www.portalfiscal.inf.br/nfe}infInut", "infInut"]:
        found = root.find(f".//{tag_name}")
        if found is not None and "Id" in found.attrib:
            inf_node = found
            ref_uri = f"#{found.attrib['Id']}"
            break

    if inf_node is None:
        if "Id" in root.attrib:
            inf_node = root
            ref_uri = f"#{root.attrib['Id']}"
        else:
            raise SignatureError("Nó raiz ou filho com atributo 'Id' (ex: infNFe, infEvento) não encontrado para assinatura.")

    # Remover assinaturas antigas se existirem
    for old_sig in root.findall(f".//{{{XMLDSIG_NS}}}Signature"):
        parent = old_sig.getparent()
        if parent is not None:
            parent.remove(old_sig)

    # 1. Canonicalizar o nó <infNFe> / <infEvento> e calcular DigestValue (SHA-1)
    c14n_inf_bytes = etree.tostring(inf_node, method="c14n", exclusive=False, with_comments=False)
    digest = hashes.Hash(hashes.SHA1())
    digest.update(c14n_inf_bytes)
    digest_val = base64.b64encode(digest.finalize()).decode('ascii')

    # 2. Construir <SignedInfo>
    nsmap = {None: XMLDSIG_NS}
    signed_info = etree.Element(f"{{{XMLDSIG_NS}}}SignedInfo", nsmap=nsmap)

    can_method = etree.SubElement(signed_info, f"{{{XMLDSIG_NS}}}CanonicalizationMethod")
    can_method.set("Algorithm", C14N_ALGO)

    sig_method = etree.SubElement(signed_info, f"{{{XMLDSIG_NS}}}SignatureMethod")
    sig_method.set("Algorithm", RSA_SHA1_ALGO)

    reference = etree.SubElement(signed_info, f"{{{XMLDSIG_NS}}}Reference")
    reference.set("URI", ref_uri)

    transforms = etree.SubElement(reference, f"{{{XMLDSIG_NS}}}Transforms")
    t1 = etree.SubElement(transforms, f"{{{XMLDSIG_NS}}}Transform")
    t1.set("Algorithm", ENVELOPED_SIG)
    t2 = etree.SubElement(transforms, f"{{{XMLDSIG_NS}}}Transform")
    t2.set("Algorithm", C14N_ALGO)

    digest_method = etree.SubElement(reference, f"{{{XMLDSIG_NS}}}DigestMethod")
    digest_method.set("Algorithm", SHA1_ALGO)

    digest_value_el = etree.SubElement(reference, f"{{{XMLDSIG_NS}}}DigestValue")
    digest_value_el.text = digest_val

    # 3. Construir o nó <Signature> e anexar na árvore ANTES de canonicalizar SignedInfo
    signature_node = etree.Element(f"{{{XMLDSIG_NS}}}Signature", nsmap=nsmap)
    signature_node.append(signed_info)

    sig_val_el = etree.SubElement(signature_node, f"{{{XMLDSIG_NS}}}SignatureValue")
    key_info = etree.SubElement(signature_node, f"{{{XMLDSIG_NS}}}KeyInfo")
    x509_data = etree.SubElement(key_info, f"{{{XMLDSIG_NS}}}X509Data")
    x509_cert = etree.SubElement(x509_data, f"{{{XMLDSIG_NS}}}X509Certificate")
    x509_cert.text = cert_b64

    parent_of_inf = inf_node.getparent()
    if parent_of_inf is not None:
        parent_of_inf.append(signature_node)
    else:
        root.append(signature_node)

    # 4. Canonicalizar <SignedInfo> exatamente na posição final e assinar com RSA (PKCS#1 v1.5 + SHA-1)
    c14n_signed_info = etree.tostring(signed_info, method="c14n", exclusive=False, with_comments=False)

    signature_bytes = private_key.sign(
        c14n_signed_info,
        padding.PKCS1v15(),
        hashes.SHA1()
    )
    signature_val = base64.b64encode(signature_bytes).decode('ascii')
    sig_val_el.text = signature_val

    signed_xml_bytes = etree.tostring(root, encoding="utf-8", xml_declaration=True)
    return signed_xml_bytes.decode('utf-8')


def verify_xml_signature(signed_xml_content):
    """
    Valida matematicamente a assinatura digital de um XML SEFAZ (DigestValue e SignatureValue).
    Retorna (is_valid: bool, reason: str).
    """
    try:
        if isinstance(signed_xml_content, str):
            parser = etree.XMLParser(remove_blank_text=False)
            root = etree.fromstring(signed_xml_content.encode('utf-8'), parser=parser)
        else:
            root = signed_xml_content

        sig_node = root.find(f".//{{{XMLDSIG_NS}}}Signature")
        if sig_node is None:
            return False, "Nenhum elemento <Signature> encontrado no XML."

        signed_info = sig_node.find(f"{{{XMLDSIG_NS}}}SignedInfo")
        sig_val_node = sig_node.find(f"{{{XMLDSIG_NS}}}SignatureValue")
        x509_node = sig_node.find(f".//{{{XMLDSIG_NS}}}X509Certificate")

        if signed_info is None or sig_val_node is None or x509_node is None:
            return False, "Estrutura <Signature> incompleta (falta SignedInfo, SignatureValue ou X509Certificate)."

        # Carregar chave pública do certificado X509 contido na assinatura
        cert_der = base64.b64decode(x509_node.text.strip())
        cert = x509.load_der_x509_certificate(cert_der)
        pub_key = cert.public_key()

        # Verificar SignatureValue sobre SignedInfo
        sig_bytes = base64.b64decode(sig_val_node.text.strip())
        c14n_si = etree.tostring(signed_info, method="c14n", exclusive=False, with_comments=False)

        pub_key.verify(
            sig_bytes,
            c14n_si,
            padding.PKCS1v15(),
            hashes.SHA1()
        )

        # Verificar DigestValue do infNFe
        ref_node = signed_info.find(f"{{{XMLDSIG_NS}}}Reference")
        expected_digest = ref_node.find(f"{{{XMLDSIG_NS}}}DigestValue").text.strip()
        ref_uri = ref_node.attrib.get("URI", "").lstrip("#")

        inf_node = root.find(f".//*[@Id='{ref_uri}']")
        if inf_node is None:
            return False, f"Nó referenciado '{ref_uri}' não encontrado no documento."

        c14n_inf = etree.tostring(inf_node, method="c14n", exclusive=False, with_comments=False)
        d = hashes.Hash(hashes.SHA1())
        d.update(c14n_inf)
        actual_digest = base64.b64encode(d.finalize()).decode('ascii')

        if actual_digest != expected_digest:
            return False, f"DigestValue incorreto (Esperado: {expected_digest}, Calculado: {actual_digest})"

        return True, "Assinatura digital XML-DSig SEFAZ 100% Válida e Autêntica."

    except Exception as e:
        return False, f"Falha na validação criptográfica da assinatura: {str(e)}"
