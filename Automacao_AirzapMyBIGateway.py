import os
import uuid  # <<< 1. ADICIONADO PARA FORÇAR O EMPACOTADOR A INCLUIR A BIBLIOTECA
from decimal import Decimal
import ConectDBcorp


def ler_arquivo_sql(caminho):
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        with open(caminho, "r", encoding="latin-1") as f:
            return f.read()


def executar_pipeline_etl(caminho_sql, tabela_destino, colunas_indice_str=""):
    """Executa o processo de ETL individual para uma consulta e tabela destino"""
    print(f"\n{'-' * 60}\nProcessando: {tabela_destino}\n{'-' * 60}")

    if not os.path.exists(caminho_sql):
        print(f"ERRO: Arquivo SQL não encontrado em: {caminho_sql}")
        return

    query_original = ler_arquivo_sql(caminho_sql)
    query_execucao = (
            "SET NOCOUNT ON;\nSET ANSI_WARNINGS OFF;\n" + query_original
    )

    try:
        print("Conectando na ORIGEM (DBCORP) e extraindo dados...")
        cursor_origem = ConectDBcorp.ConectaSQL()
        cursor_origem.execute(query_execucao)

        while cursor_origem.description is None:
            if not cursor_origem.nextset():
                break

        descricao_colunas = cursor_origem.description
        colunas_nomes = [col[0] for col in descricao_colunas]
        linhas = cursor_origem.fetchall()
        cursor_origem.connection.close()

        total_linhas = len(linhas)
        print(f"Sucesso! Extraídas {total_linhas} linhas e {len(colunas_nomes)} colunas.")

        if total_linhas == 0:
            print(f"Nenhum dado retornado para {tabela_destino}. Pulando...")
            return

        print("Mapeando os tipos de dados originais...")
        lista_definicao_colunas = []

        for col in descricao_colunas:
            nome_col, tipo_python, _, tamanho, precisao, escala = col[:6]

            if tipo_python == str:
                tipo_sql = (
                    f"VARCHAR({tamanho})"
                    if tamanho and 0 < tamanho <= 8000
                    else "VARCHAR(8000)"
                )
            elif tipo_python == int:
                tipo_sql = "INT"
            elif tipo_python == float or tipo_python == Decimal:
                tipo_sql = (
                    f"NUMERIC({precisao}, {escala})"
                    if precisao and escala
                    else "FLOAT"
                )
            elif tipo_python == bool:
                tipo_sql = "BIT"
            elif tipo_python == uuid.UUID or "uuid" in str(tipo_python).lower():
                tipo_sql = "UNIQUEIDENTIFIER"
            elif "datetime" in str(tipo_python).lower() or "date" in str(tipo_python).lower():
                tipo_sql = "DATETIME"
            else:
                tipo_sql = "VARCHAR(8000)"

            lista_definicao_colunas.append(f"[{nome_col}] {tipo_sql}")

        string_colunas_sql = ", ".join(lista_definicao_colunas)

        print(f"Conectando no DESTINO (SW2019) e gravando em {tabela_destino}...")
        conn_destino = ConectDBcorp.ConectaAirzapBI()
        cursor_destino = conn_destino.cursor()
        cursor_destino.fast_executemany = True

        cursor_destino.execute(
            f"IF OBJECT_ID('{tabela_destino}', 'U') IS NOT NULL DROP TABLE {tabela_destino}"
        )
        cursor_destino.execute(
            f"CREATE TABLE {tabela_destino} ({string_colunas_sql})"
        )

        placeholders = ", ".join(["?" for _ in colunas_nomes])
        query_insert = f"INSERT INTO {tabela_destino} VALUES ({placeholders})"

        linhas_limpas = [
            tuple(float(v) if isinstance(v, Decimal) else v for v in linha)
            for linha in linhas
        ]
        cursor_destino.executemany(query_insert, linhas_limpas)

        # ==========================================
        # NOVO BLOCO DINÂMICO DE CRIAÇÃO DE ÍNDICES
        # ==========================================
        print("Criando índices de performance...")

        # Só tenta criar índices se o atributo existir e não estiver vazio
        if colunas_indice_str:
            # Quebra a string "Status, ItemEmpresaId" numa lista e remove espaços vazios
            lista_indices_pedidos = [col.strip() for col in colunas_indice_str.split(',')]

            for coluna_pedida in lista_indices_pedidos:
                # Confirma se a coluna realmente veio do SELECT original
                if coluna_pedida in colunas_nomes:
                    # Gera um nome de índice limpo. Ex: idx_status_tbconsultaproducao
                    nome_indice = f"idx_{coluna_pedida.lower()}_{tabela_destino.lower()}"

                    try:
                        cursor_destino.execute(
                            f"CREATE INDEX {nome_indice} ON {tabela_destino} ([{coluna_pedida}])"
                        )
                        print(f" -> Índice '{nome_indice}' criado com sucesso na coluna [{coluna_pedida}].")
                    except Exception as err:
                        print(f" -> ERRO ao criar índice na coluna [{coluna_pedida}]: {err}")
                else:
                    print(f" -> AVISO: A coluna [{coluna_pedida}] configurada no XML não existe na tabela. Pulando...")
        else:
            print(" -> Nenhum índice configurado no XML para esta tabela.")

        # ==========================================

        conn_destino.commit()
        conn_destino.close()

        print(f"Fim do processo para a tabela: {tabela_destino}")

    except Exception as e:
        print(f"\nERRO CRÍTICO NA EXECUÇÃO DA TABELA {tabela_destino}:\n{repr(e)}")


def main():
    print("Buscando consultas configuradas no XML...")

    consultas_xml = ConectDBcorp.config.getElementsByTagName("ConsultaETL")

    if not consultas_xml:
        print("Nenhuma tag <ConsultaETL> foi encontrada no XML.")
        return

    print(f"Foram encontradas {len(consultas_xml)} consultas para executar.")

    for item in consultas_xml:
        caminho_sql = item.getAttribute("CaminhoSQL")
        tabela_destino = item.getAttribute("TabelaDestino")

        # NOVO: Pega o atributo de índices. Se a tag não existir, retorna uma string vazia ("")
        colunas_indice = item.getAttribute("ColunasIndice")

        if caminho_sql and tabela_destino:
            # Passa a string de índices para a função principal
            executar_pipeline_etl(caminho_sql, tabela_destino, colunas_indice)
        else:
            print("Aviso: Configuração incompleta detectada no XML (Caminho ou Tabela ausente).")

    print("\n=== TODO O PROCESSO EM LOTE FOI FINALIZADO ===")

if __name__ == "__main__":
    main()
    print("\n")
    os.system("pause")