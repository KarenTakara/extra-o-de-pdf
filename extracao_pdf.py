import pdfplumber
import os
import re
import sys
import glob
import pandas as pd

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
            'Total Geral': '',
            'Negociado': '',
            'Cachê': '',
            'Total Geral': '',
            'Total Bruto': '',
            'Desconto da agência': '',
            'Comissão de Agência': '',
            'Comissão da empresa': '',
            'Total Líquido': '',
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

        m = re.search(r'CNPJ:.*CNPJ:\s*([^\n]+?)\s*CEP:', texto_completo)
        if m: info['CNPJ Cliente'] = m.group(1).strip()

       # Pega todos os CEPs
        # Corrige a captura dos CEPs com validação
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



        m = re.search(r'Total Líquido\s*([\d.,]+)', texto_completo)
        if m: info['Total Líquido'] = m.group(1).strip()

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
    # Uso:
    #   python extracao_pdf.py <caminho_para_pdf>
    # Se nenhum caminho for passado, o script tentará (na ordem):
    # 1) o arquivo 'arquivo.pdf' na mesma pasta do script
    # 2) o primeiro arquivo .pdf encontrado no diretório atual
    # 3) o caminho hardcoded usado anteriormente (OneDrive)

    # caminho hardcoded antigo (mantido como última tentativa)
    caminho_hardcoded = r"C:\Users\karen.takara\OneDrive - Essie Publicidade e Comunicacao Ltda\Documentos\extracao de pdf\arquivo.pdf"

    caminho_pdf = None

    # 1) argumento de linha de comando
    if len(sys.argv) > 1:
        caminho_pdf = sys.argv[1]

    # 2) se não fornecido, tenta arquivo.pdf no mesmo diretório do script
    if not caminho_pdf:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidato = os.path.join(script_dir, 'arquivo.pdf')
        if os.path.exists(candidato):
            caminho_pdf = candidato

    # 3) se ainda não, tenta o primeiro .pdf no cwd
    if not caminho_pdf:
        lista_pdf = glob.glob(os.path.join(os.getcwd(), '*.pdf'))
        if lista_pdf:
            caminho_pdf = lista_pdf[0]

    # 4) por fim, tenta o caminho hardcoded
    if not caminho_pdf and os.path.exists(caminho_hardcoded):
        caminho_pdf = caminho_hardcoded

    # Se não encontramos nenhum caminho válido, mostra diagnóstico e sai
    if not caminho_pdf or not os.path.exists(caminho_pdf):
        print("\n❌ Arquivo PDF não encontrado.")
        print("Dicas:")
        print(" - Passe o caminho para o PDF como argumento: python extracao_pdf.py 'C:\\caminho\\para\\arquivo.pdf'")
        print(" - Coloque o PDF chamado 'arquivo.pdf' na mesma pasta do script:")
        print(f"   {os.path.dirname(os.path.abspath(__file__))}")
        print(" - Ou mova/ponha o PDF no diretório de trabalho atual:")
        print(f"   {os.getcwd()}")
        # lista arquivos .pdf no cwd para ajudar
        encontrados = glob.glob(os.path.join(os.getcwd(), '*.pdf'))
        if encontrados:
            print("\nArquivos .pdf encontrados no diretório atual:")
            for p in encontrados:
                print(f" - {p}")
        else:
            print("\nNenhum arquivo .pdf encontrado no diretório atual.")
        sys.exit(1)

    processar_pdf_unico(caminho_pdf)


def processar_varios_pdfs(diretorio):
    lista_infos = []
    arquivos_pdf = glob.glob(os.path.join(diretorio, "*.pdf"))

    if not arquivos_pdf:
        print("❌ Nenhum PDF encontrado no diretório.")
        return

    for caminho_pdf in arquivos_pdf:
        print(f"\n📄 Processando arquivo: {caminho_pdf}")
        info = extrair_informacoes_pdf(caminho_pdf)
        if info:
            lista_infos.append(info)
        else:
            print(f"⚠️ Falha ao extrair dados de: {caminho_pdf}")

    if lista_infos:
        # Converter para DataFrame e salvar em Excel
        df = pd.DataFrame(lista_infos)
        caminho_excel = os.path.join(diretorio, "resultado_extracao.xlsx")
        df.to_excel(caminho_excel, index=False)
        print(f"\n✅ Arquivo Excel gerado com sucesso: {caminho_excel}")
    else:
        print("⚠️ Nenhuma informação extraída de nenhum PDF.")


if __name__ == "__main__":
    # Se quiser passar um diretório como argumento, pode fazer: python extracao_pdf.py C:\pasta
    if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
        diretorio_pdf = sys.argv[1]
    else:
        diretorio_pdf = os.getcwd()  # Diretório atual como padrão

    processar_varios_pdfs(diretorio_pdf)