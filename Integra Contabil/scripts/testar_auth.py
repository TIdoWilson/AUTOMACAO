"""
Script de linha de comando para testar o carregamento do .env,
as configurações e a autenticação com o SERPRO.
"""

import sys
from pathlib import Path

# Descobre a pasta raiz do projeto (um nível acima de /scripts)
ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"

# Garante que a pasta src/ esteja no sys.path
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from integra_contabil_core.config import get_settings
from integra_contabil_core.auth.service import obter_tokens
from integra_contabil_core.exceptions import IntegraContabilError


def main() -> None:
    settings = get_settings()

    print("=== Configurações carregadas ===")
    print(f"ENVIRONMENT         = {settings.ENVIRONMENT}")
    print(f"SERPRO_AUTH_URL     = {settings.SERPRO_AUTH_URL}")
    print(f"INTEGRA_BASE_URL    = {settings.INTEGRA_BASE_URL}")
    print(f"HTTP_TIMEOUT (s)    = {settings.HTTP_TIMEOUT}")
    print(f"SERPRO_CONSUMER_KEY = {settings.SERPRO_CONSUMER_KEY!r}")
    print(
        f"SERPRO_CONSUMER_SECRET definido? "
        f"{'SIM' if bool(settings.SERPRO_CONSUMER_SECRET) else 'NÃO'}"
    )

    print("\n=== Testando autenticação no SERPRO ===")
    try:
        tokens = obter_tokens()
    except IntegraContabilError as exc:
        print("Falha na autenticação:")
        print(f"  {exc}")
        return

    # Não exibimos os tokens completos por segurança.
    access_preview = tokens.access_token[:10] + "..." if tokens.access_token else ""
    jwt_preview = tokens.jwt_token[:10] + "..." if tokens.jwt_token else ""

    print("Autenticação bem-sucedida!")
    print(f"token_type          = {tokens.token_type}")
    print(f"expires_in (s)      = {tokens.expires_in}")
    print(f"access_token (ini)  = {access_preview}")
    print(f"jwt_token (ini)     = {jwt_preview}")


if __name__ == "__main__":
    main()
