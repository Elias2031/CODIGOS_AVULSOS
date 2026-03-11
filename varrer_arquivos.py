from pathlib import Path
import shutil

def mover_todos_arquivos_para_raiz(pasta_raiz: str):
    raiz = Path(pasta_raiz).resolve()

    if not raiz.exists() or not raiz.is_dir():
        raise ValueError(f"Pasta inválida: {raiz}")

    for item in list(raiz.rglob("*")):
        if item.is_file() and item.parent != raiz:
            destino = raiz / item.name

            # Trata arquivos com mesmo nome
            if destino.exists():
                contador = 1
                novo_nome = f"{item.stem}_{contador}{item.suffix}"
                destino = raiz / novo_nome
                while destino.exists():
                    contador += 1
                    novo_nome = f"{item.stem}_{contador}{item.suffix}"
                    destino = raiz / novo_nome

            shutil.move(str(item), str(destino))

    print("Todos os arquivos foram movidos para a pasta raiz.")

# EXEMPLO DE USO
if __name__ == "__main__":
    mover_todos_arquivos_para_raiz(r"C:\Users\elias.junior\Desktop\automacao-backup\NOVA_FASE\AUTOMACAO_RECUPERACAO_SN\PGDAS")
