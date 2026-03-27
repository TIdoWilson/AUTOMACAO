"""
Configurações do projeto Integra Contabil.

Carrega variáveis de ambiente do arquivo .env (na raiz do projeto)
usando python-dotenv, e expõe uma função get_settings() para
obter essas configurações em qualquer lugar do código.
"""

import os
from dataclasses import dataclass
from functools import lru_cache

from dotenv import load_dotenv


@dataclass
class Settings:
    # Credenciais SERPRO
    SERPRO_CONSUMER_KEY: str
    SERPRO_CONSUMER_SECRET: str

    # URLs
    SERPRO_AUTH_URL: str
    INTEGRA_BASE_URL: str

    # Ambiente: homolog ou producao (uso interno seu)
    ENVIRONMENT: str

    # Timeout padrão em segundos para requisições HTTP
    HTTP_TIMEOUT: int


@lru_cache()
def get_settings() -> Settings:
    """
    Carrega o arquivo .env (se existir) e devolve uma instância
    única de Settings (cached).
    """
    # Carrega o .env da pasta atual / raiz do projeto
    load_dotenv()

    serpro_consumer_key = os.getenv("SERPRO_CONSUMER_KEY", "")
    serpro_consumer_secret = os.getenv("SERPRO_CONSUMER_SECRET", "")

    serpro_auth_url = os.getenv(
        "SERPRO_AUTH_URL",
        "https://autenticacao.sapi.serpro.gov.br/authenticate",
    )

    integra_base_url = os.getenv(
        "INTEGRA_BASE_URL",
        "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1",
    )

    environment = os.getenv("ENVIRONMENT", "homolog")

    http_timeout_str = os.getenv("HTTP_TIMEOUT", "30")
    try:
        http_timeout = int(http_timeout_str)
    except ValueError:
        http_timeout = 30

    return Settings(
        SERPRO_CONSUMER_KEY=serpro_consumer_key,
        SERPRO_CONSUMER_SECRET=serpro_consumer_secret,
        SERPRO_AUTH_URL=serpro_auth_url,
        INTEGRA_BASE_URL=integra_base_url,
        ENVIRONMENT=environment,
        HTTP_TIMEOUT=http_timeout,
    )
