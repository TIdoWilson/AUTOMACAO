"""
Serviços de autenticação com o SERPRO (obtenção de tokens).
"""

import base64
from typing import Dict, Any

import requests

from integra_contabil_core.config import get_settings
from integra_contabil_core.exceptions import IntegraContabilError
from integra_contabil_core.auth.models import AuthTokens


def _build_basic_auth_header(consumer_key: str, consumer_secret: str) -> str:
    """
    Monta o valor do header Authorization para autenticação Basic.

    Formato: Base64("consumer_key:consumer_secret")
    """
    token_bytes = f"{consumer_key}:{consumer_secret}".encode("utf-8")
    return base64.b64encode(token_bytes).decode("ascii")


def obter_tokens() -> AuthTokens:
    """
    Obtém access_token e jwt_token junto à API de autenticação do SERPRO.

    Fluxo básico (padrão APIs SERPRO):
    - POST para SERPRO_AUTH_URL
    - Header:
        Authorization: Basic <base64(consumer_key:consumer_secret)>
        role-type: TERCEIROS
        Content-Type: application/x-www-form-urlencoded
        Accept: application/json
    - Body:
        grant_type=client_credentials

    Retorna:
        AuthTokens com access_token, jwt_token e demais informações.
    """
    settings = get_settings()

    if not settings.SERPRO_CONSUMER_KEY or not settings.SERPRO_CONSUMER_SECRET:
        raise IntegraContabilError(
            "Credenciais SERPRO não configuradas. "
            "Verifique SERPRO_CONSUMER_KEY e SERPRO_CONSUMER_SECRET no .env."
        )

    basic_token = _build_basic_auth_header(
        settings.SERPRO_CONSUMER_KEY,
        settings.SERPRO_CONSUMER_SECRET,
    )

    headers = {
        "Authorization": f"Basic {basic_token}",
        "role-type": "TERCEIROS",
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    }

    data = {
        "grant_type": "client_credentials",
        # se a doc do SERPRO pedir mais campos (scope, etc.), adiciona aqui
    }

    try:
        resp = requests.post(
            settings.SERPRO_AUTH_URL,
            headers=headers,
            data=data,
            timeout=settings.HTTP_TIMEOUT or 30,
        )
    except requests.RequestException as exc:
        raise IntegraContabilError(
            f"Erro de comunicação com o servidor de autenticação SERPRO: {exc}"
        ) from exc

    if resp.status_code != 200:
        # Traz o texto da resposta para ajudar debug
        raise IntegraContabilError(
            f"Erro na autenticação SERPRO "
            f"(status {resp.status_code}): {resp.text}"
        )

    try:
        payload: Dict[str, Any] = resp.json()
    except ValueError as exc:
        raise IntegraContabilError(
            f"Resposta de autenticação não é um JSON válido: {resp.text}"
        ) from exc

    access_token = payload.get("access_token")
    # Alguns serviços podem usar "jwt_token" ou "jwt" → tentamos ambos
    jwt_token = payload.get("jwt_token") or payload.get("jwt")

    if not access_token or not jwt_token:
        raise IntegraContabilError(
            "Resposta de autenticação não contém access_token e/ou jwt_token. "
            f"Payload: {payload}"
        )

    return AuthTokens(
        access_token=access_token,
        jwt_token=jwt_token,
        token_type=payload.get("token_type", "Bearer"),
        expires_in=payload.get("expires_in"),
        raw=payload,
    )
