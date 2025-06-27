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

    for codigo in codigos:
        cod_formatado = formatar_codigo(codigo)

        for tipo, pasta in pastas.items():
            ext = tipo.lower()
            arquivo = encontrar_maior_revisao(pasta, cod_formatado, ext)

            if arquivo:
                arquivos_para_zipar.append(arquivo)
                print(f"[OK] Encontrado: {arquivo}")

                # Extrai a revisão do nome do arquivo
                match_rev = re.search(r'REV (\d{2})', os.path.basename(arquivo), re.IGNORECASE)
                if match_rev:
                    rev = match_rev.group(1)
                    chave_resumo = f"{codigo}"  # ou cod_formatado se preferir
                    if chave_resumo not in resumo_revisoes:
                        resumo_revisoes[chave_resumo] = rev
                    else:
                        # Guarda a maior revisão entre as encontradas
                        if int(rev) > int(resumo_revisoes[chave_resumo]):
                            resumo_revisoes[chave_resumo] = rev

            else:
                erros.append(f"[ERRO] Arquivo não encontrado para {codigo} ({tipo})")

    if erros:
        print("\nOcorreram erros:")
        for erro in erros:
            print(erro)

    agora = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_zip = f"EngenhariaAirzap_{agora}.zip"
    caminho_zip = os.path.join(pasta_saida, nome_zip)

    with zipfile.ZipFile(caminho_zip, 'w') as zipf:
        for arquivo in arquivos_para_zipar:
            zipf.write(arquivo, os.path.basename(arquivo))

    print(f"\n[OK] Arquivo ZIP criado com sucesso: {caminho_zip}")

    # Mostra o resumo final
    print("\nResumo dos arquivos incluídos:")
    for cod, rev in resumo_revisoes.items():
        print(f"  {formatar_codigo(cod)} - REV {rev}")

    input("\nPressione Enter para sair...")

if __name__ == "__main__":
    main()
