import os
import re
import pdfplumber
import pandas as pd

# Base para buscar os PDFs
base_path = r"Z:\Comercial"

# Regex para arquivos PDF com o padrão desejado
pattern = re.compile(r'\d+-\d+\.pdf$', re.IGNORECASE)

def localizar_pdfs():
    pdf_list = []
    total_geral = 0

    for year in range (2019, 2027):
        year_folder = os.path.join(base_path, f"PI {year}")
        print(f"\n🔍 Verificando: {year_folder}")

        contador_ano = 0

        if os.path.exists(year_folder):
            for root, dirs, files in os.walk(year_folder):
                for file in files:
                    if file.lower().endswith(".pdf") and pattern.search(file):
                        full_path = os.path.join(root, file)
                        pdf_list.append(full_path)
                        contador_ano += 1
                        total_geral += 1
                        print(f"{total_geral:05d} ➜ {full_path}")

            print(f" Total de PDFs encontrados em {year_folder}: {contador_ano}")
        else:
            print(" Pasta não encontrada")

    print(f"\n Total geral de PDFs encontrados: {total_geral}\n")
    return pdf_list

def extrair_informacoes_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"

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
            'Total Geral': '',
            'Negociado': '',
            'Cachê': '',
            'Total Bruto': '',
            'Desconto da agência': '',
            'Comissão de Agência': '',
            'Comissão da empresa': '',
            'Total Líquido': '',
            'Emitido Por': ''
        }

        # Aqui entram seus regex para preencher info, igual no seu código original
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

        m = re.search(r'CNPJ:.*CNPJ:\s*([^\n]+?)\s*CEP:', texto_completo)
        if m: info['CNPJ Cliente'] = m.group(1).strip()

        cep = re.findall(r'CEP:\s*(\d{5}-?\d{3})(?!\d)', texto_completo)
        if len(cep) >= 1:
            info['CEP Veículo'] = cep[0].strip()
        if len(cep) >= 2:
            info['CEP Cliente'] = cep[1].strip()

        m = re.search(r'PRODUTO:\s*([^\n]+?)\s*EMISSÃO:', texto_completo)
        if m: info['Produto'] = m.group(1).strip()

        m = re.search(r'EMISSÃO:\s*([\d/]+)', texto_completo)
        if m: info['Data de Emissão'] = m.group(1).strip()

        m = re.search(r'CAMPANHA:\s*([^\n]+?)\s*PÁGINA:', texto_completo)
        if m: info['Campanha'] = m.group(1).strip()

        m = re.search(r'TOTAL GERAL\s+(\d+)', texto_completo)
        if m:
            info['Total Geral'] = m.group(1).strip()
        else:
            info['Total Geral'] = '0'

        m = re.search(r'NEGOC\s*([\d.,]+)', texto_completo)
        if m: info['Negociado'] = m.group(1).strip()

        m = re.search(r'CACH[ÊE]\s*([\d.,]+)', texto_completo)
        if m: info['Cachê'] = m.group(1).strip()

        m = re.search(r'Total Bruto\s*([\d.,]+)', texto_completo)
        if m: info['Total Bruto'] = m.group(1).strip()

        m = re.search(r'Comissão de Agência \(20%\)\s*([\d.,]+)', texto_completo)
        if m: info['Comissão de Agência'] = m.group(1).strip()

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

        linhas = texto_completo.splitlines()
        for i, linha in enumerate(linhas):
            if re.search(r'Desconto.*Ag[eê]ncia', linha, re.IGNORECASE):
                m = re.search(r'([\d]{1,3}(?:[\.\d]{0,3})*,\d{2})', linha)
                if not m:
                    for j in range(i - 1, max(i - 4, -1), -1):
                        m = re.search(r'([\d]{1,3}(?:[\.\d]{0,3})*,\d{2})', linhas[j])
                        if m:
                            break
                if m:
                    info['Desconto da agência'] = m.group(1).strip()
                break

        m = re.search(r'Total Líquido\s*([\d.,]+)', texto_completo)
        if m: info['Total Líquido'] = m.group(1).strip()

        m = re.search(r'Emitido por:\s*([^\n]+)', texto_completo)
        if m: info['Emitido Por'] = m.group(1).strip()

        return info

    except Exception as e:
        print(f"❌ Erro ao processar o PDF {caminho_pdf}: {str(e)}")
        return None

def main():
    # Passo 1: localizar todos os PDFs
    lista_pdfs = localizar_pdfs()

    if not lista_pdfs:
        print("❌ Nenhum PDF foi encontrado para processar.")
        return

    # Passo 2: extrair as informações de cada PDF
    lista_infos = []
    for pdf_path in lista_pdfs:
        print(f"\n Processando: {pdf_path}")
        info = extrair_informacoes_pdf(pdf_path)
        if info:
            lista_infos.append(info)
        else:
            print(f"⚠️ Falha ao extrair dados do arquivo: {pdf_path}")

    # Passo 3: salvar no Excel
    if lista_infos:
        df = pd.DataFrame(lista_infos)
        caminho_excel = r"C:\projetos\extra-o-de-pdf-1\resultado_extracao.xlsx"
        df.to_excel(caminho_excel, index=False)
        print(f"\n✅ Planilha gerada com sucesso: {caminho_excel}")
    else:
        print("⚠️ Nenhuma informação foi extraída de nenhum PDF.")

if __name__ == "__main__":
    main()
