# Extractor de Correos – Automatizado y Fácil de Usar

Este proyecto permite **exportar correos de Outlook**, junto con sus anexos e información relevante, hacia archivos Excel y carpetas organizadas por fecha.  
Está diseñado especialmente para que pueda ser usado por personas sin conocimientos técnicos, gracias a un archivo `ejecutar.bat` que:

- Actualiza el programa automáticamente
- Instala las dependencias necesarias
- Ejecuta el script principal en un solo clic

---

## 📨 ¿Qué hace este programa?

- Lee correos de una carpeta específica en Outlook  
- Extrae:
  - Remitente
  - Asunto
  - Fecha
  - Cuerpo del mensaje
  - Anexos
- Guarda todo en un archivo Excel organizado
- Crea una carpeta por correo y almacena allí los anexos
- Controla errores comunes para evitar interrupciones

---

## 🔧 Requisitos (para Windows):

1. **Outlook (classic) instalado y con sesión iniciada**
2. **Python 3.10 o superior**  
   Descargar en: https://www.python.org/downloads/  
   *Activar “Add Python to PATH” durante la instalación.*
3. **Git para Windows**
   Descargar en: https://git-scm.com/download/win

---

## 🛠 Instalación (solo la primera vez)

1. Instalar Python y Git.
2. Abrir CMD (o símbolo del sistema).
3. Copiar, pegar, y ejecutar lo siguiente:
   
cd Documents

y luego:

git clone https://github.com/thnylpz/extractor-correos.git



