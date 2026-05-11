# Extractor Correos +

Aplicacion interna para exportar correos de Outlook clasico a carpetas organizadas, con PDF del correo, anexos y Excel resumen.

## Requisitos

- Windows.
- Outlook clasico configurado con la cuenta institucional.
- Microsoft Word clasico instalado.
- Python 3.x.
- Dependencias de `docs/requirements.txt`.

Instalacion de dependencias:

```powershell
python -m pip install -r docs/requirements.txt
```

## Uso

1. Ejecutar `Crear acceso directo.bat` una vez para crear el acceso directo del escritorio.
2. Abrir `Extractor Correos +`.
3. Elegir `Correos recibidos` o `Correos enviados`.
4. Elegir dia, mes y anio, o marcar `Exportar mes completo`.
5. Presionar `Exportar correos`.
6. Al finalizar, usar `Abrir carpeta de exportacion`.

## Configuracion recomendada de Outlook

El programa depende de Outlook clasico mediante COM/MAPI. Para evitar bloqueos:

- Abrir Outlook clasico antes de exportar, si es posible.
- En el icono de Outlook de la bandeja de Windows, desmarcar `Mostrar advertencias de red`.
- Si se necesitan correos antiguos, configurar la cuenta Exchange para descargar `Todo` el historial.
- Esperar a que Outlook termine de sincronizar antes de exportar meses o fechas antiguas.

Si Outlook muestra el mensaje de conexion de uso medido, el programa intenta pulsar `Conectar` automaticamente usando `pywinauto`. Si no lo consigue, pulse `Conectar` manualmente y revise `logs/extractor_correos.log`.

## Salida generada

La exportacion se crea en `Downloads`, con una estructura similar a:

```txt
Downloads/
  Correos Enviados Enero - 2025/
    2025-01-09/
      2025-01-09_CorreosEnviadosExportados.xlsx
      115114_DESTINATARIO/
        asunto.pdf
        Anexos/
```

Cada correo exportado debe tener PDF. Los archivos `.mht` se usan solo como temporales y se eliminan automaticamente. No se genera respaldo `.msg`.

## Personas y cargos

Las personas se administran en:

```txt
src/personas.json
```

Desde la ventana principal se puede abrir `Configurar personas` para agregar, editar, eliminar o duplicar cargos historicos. Antes de guardar, el configurador crea un backup en:

```txt
src/backups_personas/
```

## Logs

Los errores y detalles tecnicos se guardan en:

```txt
logs/extractor_correos.log
```

Ese archivo sirve para diagnosticar fallos de Outlook, Word, conversion a PDF, anexos o filtros de fecha.

## Problemas comunes

- `No esta conectado`: Outlook no termino de conectarse o mostro una advertencia de red.
- No aparecen correos antiguos: Outlook no tiene sincronizado todo el historial.
- Exportacion lenta: convertir correos a PDF con Word puede tardar, especialmente si tienen imagenes o anexos grandes.
- PDF fallido: revisar `logs/extractor_correos.log`.
