# Sistema de Registro de Reportes de Operación de Equipos

Sube fotos de reportes físicos → IA extrae los datos → Se guardan en Excel acumulativo → Se envía imagen al validador por email.

---

## Pasos para desplegar en Render (15 minutos)

### 1. Obtener API Key de Groq (gratis)
1. Ve a https://console.groq.com
2. Crea una cuenta gratuita
3. Ve a "API Keys" → "Create API Key"
4. Copia la key (empieza con `gsk_...`)

### 2. Subir el código a GitHub
1. Crea cuenta en https://github.com si no tienes
2. Crea un repositorio nuevo (ej: `reporte-equipos`)
3. Sube todos los archivos de esta carpeta

### 3. Desplegar en Render
1. Ve a https://render.com y crea cuenta gratuita
2. Click en "New" → "Web Service"
3. Conecta tu repositorio de GitHub
4. Configura:
   - **Name:** reporte-equipos
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT`

### 4. Configurar variables de entorno en Render
En el panel de Render → "Environment" → agrega estas variables:

| Variable | Valor |
|---|---|
| `GROQ_API_KEY` | Tu key de Groq |
| `SMTP_EMAIL` | Email que envía (ej: campamento@empresa.com) |
| `SMTP_PASSWORD` | Contraseña del email |
| `VALIDATOR_EMAIL` | Email del validador |
| `ADMIN_PIN` | PIN de 4 dígitos para el validador (ej: 7391) |

> **Nota Outlook:** Si usas Outlook/Office 365, puede que necesites generar una "contraseña de aplicación" en la configuración de seguridad de tu cuenta Microsoft.

### 5. ¡Listo!
Render te dará una URL pública tipo: `https://reporte-equipos.onrender.com`

Esa URL es la que comparte con los proveedores.

---

## Uso del sistema

**Los proveedores:**
- Abren la URL en su navegador (celular o PC)
- Escriben su nombre/empresa
- Suben la foto del reporte
- El sistema confirma que se procesó

**El validador:**
- Recibe la imagen por email automáticamente
- Puede descargar el Excel actualizado en cualquier momento desde la misma URL
- Puede descargar el ZIP con todas las imágenes de la semana

---

## Limitaciones del plan gratuito de Render
- La app se "duerme" si no la usan por 15 minutos
- Al despertar tarda ~30 segundos la primera vez
- Para el piloto de una semana es completamente aceptable

---

## Estructura del proyecto
```
reporte-equipos/
├── app.py              # Servidor Flask principal
├── requirements.txt    # Dependencias Python
├── render.yaml         # Configuración de Render
├── templates/
│   └── index.html      # Interfaz web
├── uploads/            # Imágenes recibidas (se crea automático)
└── outputs/            # Excel generado (se crea automático)
```
