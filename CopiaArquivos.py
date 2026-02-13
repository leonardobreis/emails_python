import os
import re
import zipfile
from datetime import datetime

# --- CONFIGURAÇÕES DE CAMINHOS ---
pasta_saida = os.path.join(os.path.expanduser("~"), "Downloads")

pastas = {
    "ETIQUETAS": r"V:\DBCorp\Engenharia\Etiquetas",
    "PDF": r"V:\DBCorp\Engenharia\PDF",
    "DXF": r"V:\DBCorp\Engenharia\DXF",
    "STEP": r"V:\DBCorp\Engenharia\STEP"
}


def formatar_codigo(codigo):
    return f"{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}"


def encontrar_maior_revisao(pasta, codigo_formatado, extensao):
    if not os.path.exists(pasta):
        return None

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
    print(f"Destino dos arquivos: {pasta_saida}")
    print("Digite os códigos (Enter duas vezes para finalizar):")

    codigos = []
    while True:
        entrada = input()
        if entrada == "": break
        codigos.append(entrada.strip())

    arquivos_para_zipar = []
    resumo_revisoes = {}
    erros = []

    # Dicionário para validar se o código tem o "mínimo necessário" (Etiqueta OU PDF)
    item_validado = {codigo: False for codigo in codigos}

    for codigo in codigos:
        cod_formatado = formatar_codigo(codigo)
        encontrou_etiqueta_neste_codigo = False

        # 1. Busca Etiqueta
        arquivo_etiqueta = encontrar_maior_revisao(pastas["ETIQUETAS"], cod_formatado, "pdf")
        if arquivo_etiqueta:
            arquivos_para_zipar.append(arquivo_etiqueta)
            item_validado[codigo] = True
            encontrou_etiqueta_neste_codigo = True
            print(f"[OK] Etiqueta encontrada para {codigo}")

        # 2. Busca PDF
        arquivo_pdf = encontrar_maior_revisao(pastas["PDF"], cod_formatado, "pdf")
        if arquivo_pdf:
            arquivos_para_zipar.append(arquivo_pdf)
            item_validado[codigo] = True  # Valida se ainda não foi validado pela etiqueta
            print(f"[OK] PDF encontrado para {codigo}")

            # Coleta dados de revisão para o resumo (baseado no PDF)
            match_rev = re.search(r'REV (\d{2})', os.path.basename(arquivo_pdf), re.IGNORECASE)
            if match_rev:
                rev = match_rev.group(1)
                data_mod = datetime.fromtimestamp(os.path.getmtime(arquivo_pdf)).strftime("%d/%m/%Y %H:%M")
                resumo_revisoes[codigo] = (rev, data_mod)

        # 3. Busca DXF e STEP (Apenas se o item já foi validado por Etiqueta ou PDF)
        for tipo in ["DXF", "STEP"]:
            arquivo = encontrar_maior_revisao(pastas[tipo], cod_formatado, tipo.lower())
            if arquivo:
                arquivos_para_zipar.append(arquivo)
                print(f"[OK] {tipo} encontrado para {codigo}")

        # 4. Verificação de Erro: Se não achou Etiqueta NEM PDF
        if not item_validado[codigo]:
            erros.append(f"[ERRO] Nem Etiqueta nem PDF encontrados para {codigo}")

    # Gera o ZIP se todos os códigos digitados tiverem ao menos um arquivo base (Etiqueta ou PDF)
    pode_gerar_zip = all(item_validado.values())

    if erros:
        print("\nRelatório de Inconsistências:")
        for erro in erros:
            print(erro)

    if pode_gerar_zip and arquivos_para_zipar:
        agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_zip = os.path.join(pasta_saida, f"EngenhariaAirzap_{agora}.zip")

        try:
            with zipfile.ZipFile(caminho_zip, 'w') as zipf:
                for arquivo in set(arquivos_para_zipar):
                    zipf.write(arquivo, os.path.basename(arquivo))
            print(f"\n[SUCESSO] ZIP criado: {caminho_zip}")
        except Exception as e:
            print(f"\n[ERRO] Falha ao criar ZIP: {e}")
    else:
        print("\n[AVISO] ZIP não gerado. Alguns códigos não possuem arquivos base (Etiqueta ou PDF).")

    input("\nPressione Enter para sair...")


if __name__ == "__main__":
    main()