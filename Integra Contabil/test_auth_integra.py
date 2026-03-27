# test_auth_integra.py
"""
Script de teste para auth_integra.py.
Tenta executar obter_tokens() e captura exceções para validação.
"""

import sys
import os
from pathlib import Path

# Adiciona o diretório atual ao path para importar auth_integra
sys.path.insert(0, str(Path(__file__).parent))

try:
    from auth_integra import obter_tokens, _limpar_caminho_pfx
    print("✅ Imports bem-sucedidos.")
except ImportError as e:
    print(f"❌ Erro de importação: {e}")
    sys.exit(1)

# Teste da função _limpar_caminho_pfx
print("\n--- Teste da função _limpar_caminho_pfx ---")
test_cases = [
    '="C:\\path\\to\\file.pfx"',
    "'C:\\path\\to\\file.pfx'",
    "  =W:\\path\\file.pfx  ",
    "normal_path.pfx",
    None
]
for case in test_cases:
    result = _limpar_caminho_pfx(case)
    print(f"Input: {case!r} -> Output: {result!r}")

# Teste da função obter_tokens
print("\n--- Teste da função obter_tokens ---")
try:
    access_token, jwt_token = obter_tokens()
    print("✅ Autenticação bem-sucedida!")
    print(f"Access Token: {access_token[:20]}...")  # Mostra apenas início
    print(f"JWT Token: {jwt_token[:20]}...")
except RuntimeError as e:
    print(f"❌ Erro de configuração: {e}")
    print("Verifique se o arquivo .env existe e contém as variáveis necessárias.")
except FileNotFoundError as e:
    print(f"❌ Arquivo PFX não encontrado: {e}")
except Exception as e:
    print(f"❌ Erro inesperado: {e}")

print("\nTeste concluído.")