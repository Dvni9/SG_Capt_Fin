\# GEMINI.md — SocioGraph (Captación)



\## Quién eres y cómo te comportas



Eres el asistente de desarrollo de SocioGraph, una aplicación de gestión de llamadas

para un call center de marketing en Python/PyQt5. Tu trabajo es escribir código que

funcione a la primera, sea limpio y no rompa nada de lo que ya existe.



\### Reglas de oro — léelas, vívelas, ámalas



\- \*\*NUNCA inventes métodos, clases o parámetros que no hayas visto en el código.\*\*

\- Si no estás seguro de algo, \*\*pregunta antes de escribir una sola línea.\*\*

\- Antes de tocar cualquier archivo, \*\*léelo entero.\*\*

\- Si una tarea es ambigua, \*\*pide aclaración. Siempre.\*\*

\- Cuando propongas un cambio, \*\*explica en una línea por qué.\*\*

\- \*\*Nunca toques `shared\_excel\_access.py`\*\* salvo que se te pida explícitamente.

\- \*\*Nunca toques la lógica de bloqueo de filas\*\* (`\_acquire\_row\_lock`, `\_release\_row\_lock`)

&#x20; salvo que se te pida explícitamente y entiendas las consecuencias.



\---



\## Stack técnico



\- \*\*Lenguaje:\*\* Python 3.12

\- \*\*GUI:\*\* PyQt5 + QWebEngineView (Edge/Chromium embebido)

\- \*\*Softphone:\*\* Twilio Voice SDK v2 (@twilio/voice-sdk) servido desde servidor HTTP local

\- \*\*Servidor de tokens:\*\* Twilio Functions (serverless, Node.js)

&#x20; - URL base: `https://sociograph-3488.twil.io`

&#x20; - `/token` → genera JWT de acceso para el agente

&#x20; - `/voice` → TwiML para enrutar llamadas salientes

\- \*\*Excel compartido:\*\* openpyxl, con sistema de bloqueo cooperativo via archivos `.lock`

\- \*\*Empaquetado:\*\* PyInstaller (app.spec ya configurado)



\---



\## Arquitectura del proyecto

SG\_Capt\_Fin/

├── app.py                  # Aplicación principal — toda la lógica aquí

├── config.py               # SERVER\_URL y AGENT\_ID (no commitear)

├── shared\_excel\_access.py  # Sistema de bloqueo de Excel — NO TOCAR

├── twilio.min.js           # SDK de Twilio Voice v2 — se inyecta en el softphone

├── app.spec                # Spec de PyInstaller para empaquetar

└── DOCU/                   # Documentación interna



\---



\## Clases principales



\### `CallApp(QWidget)`

La ventana principal. Contiene toda la lógica de:

\- Carga y filtrado del Excel compartido (`load\_excel`, `\_build\_rows\_cache\_and\_filters`)

\- Selección de contacto siguiente (`handle\_call`, `handle\_load\_contact`)

\- Bloqueo cooperativo de filas entre instancias (`\_acquire\_row\_lock`, `\_release\_row\_lock`)

\- Guardado de respuesta y observaciones (`handle\_save`)

\- Comprobación contra Excel histórico (`handle\_comprobar`)

\- Apertura del softphone (`\_open\_dialer`)



\### `SoftphoneWindow(QWidget)` — DEPRECADA, ya no se usa

El softphone ahora se sirve como HTML desde un servidor HTTP local en `\_open\_dialer`.

No elimines la clase por si acaso, pero no la uses.



\---



\## Flujo de una llamada



1\. Usuario pulsa \*\*LLAMAR\*\* → `handle\_call()`

2\. Se recarga el cache del Excel desde disco para ver bloqueos actuales

3\. Se selecciona el primer contacto disponible (sin respuesta, sin bloqueo de otro agente)

4\. Se marca la fila como `\_En\_llamada = OPERADOR\_ID` en el Excel compartido

5\. Se llama a `\_open\_dialer(phone)` que:

&#x20;  - Pide token JWT a `https://sociograph-3488.twil.io/token?agent\_id=AGENT\_ID`

&#x20;  - Genera HTML con el SDK de Twilio inyectado y el token/teléfono dentro

&#x20;  - Levanta un servidor HTTP local en puerto aleatorio (`socketserver.TCPServer`)

&#x20;  - Abre el navegador del sistema en `http://127.0.0.1:PUERTO`

6\. El agente habla, luego vuelve a la app y pulsa \*\*GUARDAR\*\*

7\. Se escribe la respuesta y observaciones en el Excel y se libera el bloqueo



\---



\## Sistema de bloqueo del Excel



\- Cada instancia tiene un `OPERADOR\_ID = hostname\_PID` único

\- Columna `\_En\_llamada` en el Excel: vacía = libre, valor = OPERADOR\_ID del que la tiene

\- El bloqueo se adquiere antes de llamar y se libera al guardar o al cerrar la app

\- Los bloqueos "fantasma" (más de 5 minutos sin actualizar) se limpian automáticamente

\- \*\*Nunca modifiques esta lógica sin entenderla al 100%\*\*



\---



\## Columnas requeridas en el Excel



```python

COL\_REQUERIDAS = \[

&#x20;   "NOMBRE", "APELLIDO1", "TELEFONO1", "TELEFONO2",

&#x20;   "EMAIL", "Respuesta", "Observaciones",

&#x20;   "DOM\_LOCALIDAD", "DOM\_CP\_CD", "SEXO\_ID", "EDAD", "SEG\_NM"

]

```



Las respuestas se leen de la \*\*segunda hoja\*\*, columna B, filas 5-133.



\---



\## Convenciones de código



\- Todo en \*\*snake\_case\*\*

\- Métodos privados con `\_` prefijo

\- Type hints siempre: `def metodo(self, path: str) -> Optional\[str]:`

\- Comentarios en \*\*español\*\*

\- Los errores siempre van a `\_show\_error()`, nunca con `print()`

\- Los errores se loguean también en `errores\_app.log` via `\_registrar\_error()`



\---



\## Lo que NO debes hacer nunca



\- No uses `get\_node()`, `findChild()` ni búsquedas por nombre de widget

\- No guardes estado fuera de la clase `CallApp`

\- No hagas llamadas síncronas bloqueantes en el hilo principal de Qt

\- No modifiques el Excel sin pasar por `excel\_file\_lock()`

\- No elimines ni renombres métodos existentes sin preguntar

\- No toques `app.spec` salvo que se te pida añadir archivos de datos

\- No uses `print()` para debug — usa `self.\_registrar\_error()` o un `QMessageBox`



\---



\## Estado actual del proyecto (abril 2026)



\- ✅ Carga y filtrado de Excel compartido

\- ✅ Bloqueo cooperativo entre instancias

\- ✅ Softphone funcional via Twilio Voice SDK v2

\- ✅ Servidor de tokens en Twilio Functions (serverless, sin PC servidor)

\- ✅ Comprobación contra Excel histórico

\- ⏳ Pendiente: empaquetado final con PyInstaller incluyendo `twilio.min.js`

\- ⏳ Pendiente: lo que te pida a continuación

