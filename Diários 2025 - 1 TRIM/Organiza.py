import os
import shutil

def organizar_arquivos(diretorio):
    disciplinas = [
#         "Laboratório LGG", "Biologia", "Laboratório CHS", "Ed. Física", "Estudo Orientado", "Física", "Geografia", "História", "Inglês", "Matemática", "Português", "Química", "Reforço Escolar", "Filosofia"
          "Artes", "Biologia", "CA1", "CA2", "Laboratório CHS", "Ed. Física", "Estudo Orientado", "Física", "Geografia", "História", "Inglês", "Matemática", "Português", "Química", "Reforço Escolar", "Sociologia"
#          "Biologia", "CA1", "CA2", "CA3", "Ed. Física", "Estudo Orientado", "Filosofia", "Física", "Geografia", "História", "Inglês", "Matemática", "Português", "Química", "Reforço Escolar", "Sociologia"
    ]

    if not os.path.exists(diretorio):
        print("Diretório não encontrado.")
        return

    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(".xlsx"):
            partes = arquivo.split()
            if len(partes) > 4:
                codigo_turma = partes[-1].split('.')[0]  # Pega as últimas 4 posições antes do .xlsx
                pasta_turma = os.path.join(diretorio, codigo_turma)
                os.makedirs(pasta_turma, exist_ok=True)

                for disciplina in disciplinas:
                    novo_nome = f"{arquivo[:-5]} - {disciplina}.xlsx"
                    destino = os.path.join(pasta_turma, novo_nome)
                    shutil.copy(os.path.join(diretorio, arquivo), destino)
                    print(f"Copiado: {arquivo} -> {destino}")

    print("Organização concluída!")

# Defina o caminho do diretório onde estão os arquivos
diretorio_raiz = os.path.abspath("Diários 2025 - 2TRIM/2a série")
organizar_arquivos(diretorio_raiz)