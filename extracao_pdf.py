import pdfplumber
import os
import re

def extrair_informacoes_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_completo += pagina.extract_text() + "\n"
            print(texto_completo)
        info = {
            'PI': '',
            'AP': '',
            'Data de Emissão': '',
            'Nome do Arquivo': os.path.basename(caminho_pdf),
            'Caminho do Arquivo': caminho_pdf,
            'Veículo': '',
            'Endereço Veículo': '',
            'Cidade Veículo': '',
            'Estado Veículo': '',
            'CNPJ Veículo': '',
            'CEP Veículo': '',
            'Telefone Veículo': '',
            'Cliente': '',
            'Endereço Cliente': '',
            'Cidade Cliente': '',
            'Estado Cliente': '',
            'CNPJ Cliente': '',
            'CEP Cliente': '',
            'Período Veíc.': '',
            'Vencimento': '',
            'Produto': '',
            'Campanha': '',
            'Total Inserções': '',
            'Negociado': '',
            'Cachê': '',
            'Total Item': '',
            'Total Bruto': '',
            'Desconto da agência': '',
            'Comissão de Agência': '',
            'Comissão da SBC': '',
            'Comissão da ESSIÊ': '',
            'Comissão da empresa': '',
            'Total Líquido': '',
            'Controle Interno': '',
            'Emitido Por': ''
        }

        # ====== Extrações com regex ======
        m = re.search(r'VE[IÍ]CULO:\s*([^\n]+?)\s*FONE:', texto_completo)
        if m: info['Veículo'] = m.group(1).strip()

        m = re.search(r'FONE:\s*([^\n]+?)\s*CLIENTE:', texto_completo)
        if m: info['Telefone Veículo'] = m.group(1).strip()

        m = re.search(r'CLIENTE:\s*([^\n]+?)\s*PI:', texto_completo)
        if m: info['Cliente'] = m.group(1).strip()

        m = re.search(r'PI:\s*(\d+)', texto_completo)
        if m: info['PI'] = m.group(1).strip()

        m = re.search(r'ENDEREÇO:\s*([^\n]+?)\s*ENDEREÇO:', texto_completo)
        if m: info['Endereço Veículo'] = m.group(1).strip()

        m = re.search(r'ENDEREÇO:.*ENDEREÇO:\s*([^\n]+?)\s*AP:', texto_completo)
        if m: info['Endereço Cliente'] = m.group(1).strip()

        m = re.search(r'AP:\s*([\w\d]+)', texto_completo)
        if m: info['AP'] = m.group(1).strip()

        m = re.search(r'CIDADE:\s*([^\n]+?)\s*ESTADO:', texto_completo)
        if m: info['Cidade Veículo'] = m.group(1).strip()

        m = re.search(r'ESTADO:\s*([^\n]+?)\s*CIDADE:', texto_completo)
        if m: info['Estado Veículo'] = m.group(1).strip()

        m = re.search(r'CIDADE:.*CIDADE:\s*([^\n]+?)\s*ESTADO:', texto_completo)
        if m: info['Cidade Cliente'] = m.group(1).strip()

        m = re.search(r'ESTADO:.*ESTADO:\s*([^\n]+?)\s*PER[IÍ]ODO', texto_completo)
        if m: info['Estado Cliente'] = m.group(1).strip()

        m = re.search(r'PER[IÍ]ODO VE[IÍ]C\.?:\s*([^\n]+)', texto_completo)
        if m: info['Período Veíc.'] = m.group(1).strip()

        m = re.search(r'VENCIMENTO:\s*([^\n]+)', texto_completo)
        if m: info['Vencimento'] = m.group(1).strip()

        m = re.search(r'CNPJ:\s*([^\n]+?)\s*CEP:', texto_completo)
        if m: info['CNPJ Veículo'] = m.group(1).strip()

        m = re.search(r'CEP:\s*([^\n]+?)\s*CNPJ:', texto_completo)
        if m: info['CEP Veículo'] = m.group(1).strip()

        m = re.search(r'CNPJ:.*CNPJ:\s*([^\n]+?)\s*CEP:', texto_completo)
        if m: info['CNPJ Cliente'] = m.group(1).strip()

        m = re.search(r'CEP:.*CEP:\s*([\d\- ]{8,9})', texto_completo)
        if m: info['CEP Cliente'] = m.group(1).strip()

        m = re.search(r'PRODUTO:\s*([^\n]+?)\s*EMISSÃO:', texto_completo)
        if m: info['Produto'] = m.group(1).strip()

        m = re.search(r'EMISSÃO:\s*([\d/]+)', texto_completo)
        if m: info['Data de Emissão'] = m.group(1).strip()

        m = re.search(r'CAMPANHA:\s*([^\n]+?)\s*PÁGINA:', texto_completo)
        if m: info['Campanha'] = m.group(1).strip()

        m = re.search(r'TOTAL GERAL\s*(\d+)\s*([\d.,]+)', texto_completo)
        if m:
            info['Total Inserções'] = m.group(1).strip()
            info['Total Item'] = m.group(2).strip()

        m = re.search(r'NEGOC\s*([\d.,]+)', texto_completo)
        if m: info['Negociado'] = m.group(1).strip()

        m = re.search(r'CACH[ÊE]\s*([\d.,]+)', texto_completo)
        if m: info['Cachê'] = m.group(1).strip()

        m = re.search(r'Total Bruto\s*([\d.,]+)', texto_completo)
        if m: info['Total Bruto'] = m.group(1).strip()

        m = re.search(r'Comissão de Agência \(20%\)\s*([\d.,]+)', texto_completo)
        if m: info['Comissão de Agência'] = m.group(1).strip()

        m = re.search(r'Comiss[aã]o da SBC.*?([\d.,]+)', texto_completo, re.IGNORECASE)
        if m: info['Comissão da SBC'] = m.group(1).strip()

        m = re.search(r'Comiss[aã]o da\s*ESSI[ÊE]?.*?([\d.,]+)', texto_completo, re.IGNORECASE)
        if m: info['Comissão da ESSIÊ'] = m.group(1).strip()

        # --- Comissão da Empresa (Agência ou SBC) ---
        m = re.search(
        r'Comiss[aã]o\s*da\s*ESSI[ÊE]?(?:\s*\(20%.*?\))?[\s:–-]*[R$ ]*([\d.,]+)',
        texto_completo,
        re.IGNORECASE | re.DOTALL
        )
        if not m:
            m = re.search(
                r'Comiss[aã]o\s*da\s*SBC(?:\s*\(20%.*?\))?[\s:–-]*[R$ ]*([\d.,]+)',
                texto_completo,
                re.IGNORECASE | re.DOTALL
            )

        if m:
            info['Comissão da empresa'] = m.group(1).strip()
        else:
            info['Comissão da empresa'] = ''


        # === Captura Desconto da Agência ===
        linhas = texto_completo.splitlines()
        for i, linha in enumerate(linhas):
            if re.search(r'Desconto.*Ag[eê]ncia', linha, re.IGNORECASE):
        # procura número na mesma linha primeiro
                m = re.search(r'([\d]{1,3}(?:[\.\d]{0,3})*,\d{2})', linha)
                if not m:
            # se não encontrar, procura nas 3 linhas acima
                    for j in range(i - 1, max(i - 4, -1), -1):
                     m = re.search(r'([\d]{1,3}(?:[\.\d]{0,3})*,\d{2})', linhas[j])
                    if m:
                        break
                if m:
                    info['Desconto da agência'] = m.group(1).strip()
                break  # sai do loop depois de encontrar

        # === Captura Comissão da ESSIÊ ===
        for linha in linhas:
            if re.search(r'Comiss[aã]o.*ESSI[ÊE]?', linha, re.IGNORECASE):
                m = re.search(r'([\d]{1,3}(?:[\.\d]{0,3})*,\d{2})', linha)
                if m:
                    info['Comissão da ESSIÊ'] = m.group(1).strip()
                break



        m = re.search(r'Total Líquido\s*([\d.,]+)', texto_completo)
        if m: info['Total Líquido'] = m.group(1).strip()

        m = re.search(r'Controle Interno SBC: PI - (\d+)', texto_completo)
        if m: info['Controle Interno'] = m.group(1).strip()

        m = re.search(r'Emitido por:\s*([^\n]+)', texto_completo)
        if m: info['Emitido Por'] = m.group(1).strip()

        # Retorna o dicionário de informações
        return info

    except Exception as e:
        print(f"❌ Erro ao processar o PDF {caminho_pdf}: {str(e)}")
        return None


def processar_pdf_unico(caminho_pdf):
    print(f"\n📄 Processando arquivo: {caminho_pdf}")
    info = extrair_informacoes_pdf(caminho_pdf)

    if info:
        print("\n========= RESULTADO DA EXTRAÇÃO =========\n")
        for chave, valor in info.items():
            print(f"{chave}: {valor}")
        print("\n=========================================\n")
    else:
        print("❌ Nenhuma informação foi extraída.")


if __name__ == "__main__":
    # 👇 Coloque o caminho do seu PDF aqui
    caminho_pdf = r"C:\Users\karen.takara\OneDrive - Essie Publicidade e Comunicacao Ltda\Documentos\extracao de pdf\arquivo.pdf"
    processar_pdf_unico(caminho_pdf)
