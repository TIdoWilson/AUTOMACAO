"""
Modelos de dados relacionados à autenticação no SERPRO.
"""

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class AuthTokens:
    """
    Representa o resultado da autenticação no SERPRO.

    - access_token: usado no header Authorization: Bearer ...
    - jwt_token: usado no header jwt_token
    - token_type: geralmente "Bearer"
    - expires_in: tempo de expiração em segundos (se informado)
    - raw: resposta original completa (JSON) para debug/log
    """

    access_token: str
    jwt_token: str
    token_type: str = "Bearer"
    expires_in: Optional[int] = None
    raw: Optional[Dict[str, Any]] = None
