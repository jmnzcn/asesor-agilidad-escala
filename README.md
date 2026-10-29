# Asesor de Agilidad a Escala (Streamlit + Gemini)

Agente que recomienda **Nexus, LeSS, SAFe o Disciplined Agile** con la matriz de decisión del equipo (PF-3872, UCR).

- **Interpretar:** Gemini lee el caso, cita la evidencia y llena la ficha de captura (nivel + estado por criterio).
- **Calcular:** `motor_matriz.py` aplica las reglas R1–R9 con pesos, niveles y puntajes leídos de `Matriz_Decision_Agilidad_Escala.xlsx`.
- **Auditar:** el botón «Descargar Excel» entrega la matriz con la hoja *Caso_Evaluado* llena; Excel recalcula y muestra los mismos totales.

Sin clave de Gemini la app funciona igual en modo manual (ficha editable, cálculo, gráfico y Excel).

## Archivos

| Archivo | Qué es |
|---|---|
| `app.py` | Interfaz Streamlit |
| `motor_matriz.py` | Motor de cálculo (el mismo verificado contra las 5 hojas de caso del Excel) |
| `Matriz_Decision_Agilidad_Escala.xlsx` | Fuente única de reglas |
| `requirements.txt` | Dependencias |
| `.streamlit/secrets.toml.example` | Plantilla para la clave de Gemini |

## Publicarla gratis (≈ 15 minutos)

### 1. Clave de Gemini (gratis)
1. Entrar a **https://aistudio.google.com/apikey** con una cuenta de Google.
2. **Create API key** y copiarla. No la comparta ni la suba a GitHub.

### 2. Repositorio en GitHub
1. En **github.com › New repository**: nombre `asesor-agilidad-escala`, **Public**, sin README.
2. Subir el contenido de esta carpeta (`App_Streamlit`): en el repositorio vacío, *uploading an existing file* y arrastrar `app.py`, `motor_matriz.py`, `Matriz_Decision_Agilidad_Escala.xlsx`, `requirements.txt`, `README.md` y `.gitignore`. **No** suba `secrets.toml`.

   O desde una terminal en esta carpeta:
   ```bash
   git init && git add . && git commit -m "Asesor de Agilidad a Escala" && git branch -M main
   git remote add origin https://github.com/<su-usuario>/asesor-agilidad-escala.git && git push -u origin main
   ```

### 3. Streamlit Community Cloud
1. Entrar a **https://share.streamlit.io** › *Continue with GitHub* y autorizar.
2. **Create app › Deploy a public app from GitHub**: repositorio `asesor-agilidad-escala`, rama `main`, archivo `app.py`.
3. *App URL*: elegir algo como `asesor-agilidad-pf3872`.
4. **Advanced settings › Secrets**: pegar
   ```toml
   GEMINI_API_KEY = "su-clave"
   ```
5. **Deploy**. En 2–3 minutos queda en `https://asesor-agilidad-pf3872.streamlit.app`, abierta para cualquiera sin registrarse.

### 4. Probar antes de entregar
1. Elegir **Caso 3 · Servicios profesionales**: la ficha de referencia del equipo da **Disciplined Agile 4,70**.
2. Pulsar **Analizar con IA**: Gemini llena la ficha con evidencia citada; la recomendación debe seguir siendo DA (el puntaje puede variar si la IA deja algún criterio como faltante en vez de inferido).
3. **Descargar Excel** y abrir la hoja *Caso_Evaluado*: los totales deben coincidir con la app.

> Las apps gratuitas de Streamlit se «duermen» tras unos días sin visitas. Si el profesor ve el botón *Yes, get this app back up!*, basta pulsarlo y esperar unos segundos. Conviene abrirla el día antes de la revisión.

## Ejecutar en la computadora
```bash
pip install -r requirements.txt
streamlit run app.py
```
Para usar la IA localmente, copie `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` y ponga su clave.
