import os
import re
import zipfile
from datetime import datetime

# Caminhos das pastas
pastas = {
    "PDF": r"V:\DBCorp\Engenharia\PDF",
    "DXF": r"V:\DBCorp\Engenharia\DXF",
    "STEP": r"V:\DBCorp\Engenharia\STEP"
}

# Pasta de destino
pasta_saida = r"S:\\"


def formatar_codigo(codigo):
    return f"{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}"


def encontrar_maior_revisao(pasta, codigo_formatado, extensao):
    arquivos = os.listdir(pasta)
    padrao = re.compile(rf"{re.escape(codigo_formatado)} - REV (\d{{2}})\.{extensao}$", re.IGNORECASE)

    maior_rev = -1
    arquivo_mais_recente = None

    for arquivo in arquivos:
        match = padrao.match(arquivo)
        if match:
            revisao = int(match.group(1))
            if revisao > maior_rev:
                maior_rev = revisao
                arquivo_mais_recente = os.path.join(pasta, arquivo)

    return arquivo_mais_recente


def main():
    print("Digite os códigos dos produtos, um por linha (pressione Enter duas vezes para terminar):")
    codigos = []
    while True:
        entrada = input()
        if entrada == "":
            break
        codigos.append(entrada.strip())

    arquivos_para_zipar = []
    resumo_revisoes = {}  # chave: código original, valor: maior revisão (como string)
    erros = []
    # Dicionário para verificar se o PDF foi encontrado para cada código
    pdf_encontrado_por_codigo = {codigo: False for codigo in codigos}


    for codigo in codigos:
        cod_formatado = formatar_codigo(codigo)

        # Primeiro, verificar o PDF
        arquivo_pdf = encontrar_maior_revisao(pastas["PDF"], cod_formatado, "pdf")
        if arquivo_pdf:
            arquivos_para_zipar.append(arquivo_pdf)
            print(f"[OK] PDF Encontrado: {arquivo_pdf}")
            pdf_encontrado_por_codigo[codigo] = True

            # Extrai a revisão do nome do arquivo PDF
            match_rev = re.search(r'REV (\d{2})', os.path.basename(arquivo_pdf), re.IGNORECASE)
            if match_rev:
                rev = match_rev.group(1)
                resumo_revisoes[codigo] = rev # Armazena a revisão do PDF como a principal para o resumo

            # Agora, procurar os outros tipos de arquivo se o PDF foi encontrado
            for tipo, pasta in pastas.items():
                if tipo != "PDF": # Já lidamos com o PDF
                    ext = tipo.lower()
                    arquivo = encontrar_maior_revisao(pasta, cod_formatado, ext)
                    if arquivo:
                        arquivos_para_zipar.append(arquivo)
                        print(f"[OK] Encontrado: {arquivo}")
                    else:
                        erros.append(f"[AVISO] Arquivo {tipo} não encontrado para {codigo}")
        else:
            erros.append(f"[ERRO] PDF não encontrado para {codigo}")
            # Se o PDF não for encontrado, não precisamos procurar os outros tipos para este código.

    # Verificar se todos os códigos têm um PDF associado
    todos_pdfs_encontrados = all(pdf_encontrado_por_codigo.values())

    if erros:
        print("\nOcorreram problemas:")
        for erro in erros:
            print(erro)

    # Só gera o ZIP se todos os PDFs necessários foram encontrados
    if todos_pdfs_encontrados:
        agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_zip = f"EngenhariaAirzap_{agora}.zip"
        caminho_zip = os.path.join(pasta_saida, nome_zip)

        try:
            with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                for arquivo in arquivos_para_zipar:
                    zipf.write(arquivo, os.path.basename(arquivo))
            print(f"\n[OK] Arquivo ZIP criado com sucesso: {caminho_zip}")

            # Mostra o resumo final
            print("\nResumo dos arquivos incluídos:")
            for cod, rev in resumo_revisoes.items():
                print(f"  {formatar_codigo(cod)} - REV {rev}")

        except Exception as e:
            print(f"\n[ERRO] Ocorreu um erro ao criar o arquivo ZIP: {e}")
    else:
        print("\n[AVISO] O arquivo ZIP não foi gerado porque um ou mais PDFs não foram encontrados para os códigos fornecidos.")


    input("\nPressione Enter para sair...")

if __name__ == "__main__":
    main()