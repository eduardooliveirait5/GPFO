from pathlib import Path

# Pasta raiz do projeto
BASE_DIR = Path(__file__).resolve().parent.parent

# Banco de dados
DATABASE_DIR = BASE_DIR / "database"
DATABASE_PATH = DATABASE_DIR / "presencas.db"

# Arquivos enviados
UPLOAD_DIR = BASE_DIR / "uploads"

# Relatórios gerados
EXPORT_DIR = BASE_DIR / "exports"

# Criação automatica das pastas necessárias
for pasta in (DATABASE_DIR, UPLOAD_DIR, EXPORT_DIR):
    pasta.mkdir(parents=True, exist_ok=True)