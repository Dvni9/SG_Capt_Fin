"""
Config local (NO commitear).

Este archivo solo contiene configuración del agente para conectarse al servidor.
Las credenciales de Twilio viven en `config_server.py` (en el PC servidor).
"""

SERVER_URL = "https://sociograph-3488.twil.io"

# Identificador del agente (se envía como ?agent_id=... al pedir /token)
AGENT_ID = "agente1"