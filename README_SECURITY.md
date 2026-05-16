# Kardex System - Cruz Roja Mexicana Delegación Durango

Sistema de gestión y consulta de Kardex para el área de capacitación.

### ⚠️ SEGURIDAD CRÍTICA (IMPORTANTE)

Para proteger la integridad de los datos de la Delegación, **NO SUBAS** los siguientes archivos a repositorios públicos de GitHub:

1. `firebase-applet-config.json`: Contiene las llaves de acceso a la base de datos de Firebase.
2. `firebase-blueprint.json`: Define la estructura de seguridad de los datos.
3. `.env`: Contiene variables de entorno privadas.

Estos archivos ya están incluidos en el `.gitignore` para prevenir subidas accidentales.

### Instrucciones para Despliegue en GitHub Pages

1. Ejecuta `npm run build` en tu computadora.
2. Sube únicamente el contenido de la carpeta `dist/` a tu repositorio de GitHub Pages.
3. Asegúrate de que en `vite.config.ts` la propiedad `base` esté configurada como `'./'` (ya configurado).

### Características
- **Carga Masiva**: Soporte para múltiples archivos Excel (.xlsx, .xls) simultáneos.
- **Consulta Inteligente**: Búsqueda por nombre o folio con autocompletado.
- **Exportación a PDF**: Generación de boletas individuales con formato oficial.
- **Filtros Avanzados**: Segmentación por edad, género y fecha de nacimiento.
