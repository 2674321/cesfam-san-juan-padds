# Instalación — Sistema PADI (V4)

El sistema vive en **Google Apps Script** vinculado a una hoja de cálculo de
Google Sheets, que actúa como base de datos. Al primer arranque el sistema
crea sus propias hojas y formatos automáticamente.

## Requisitos

- Cuenta de Google
- Node.js ≥ 16 y npm (para [clasp](https://github.com/google/clasp))

## Pasos

1. **Instalar clasp**
   ```bash
   npm install -g @google/clasp
   ```

2. **Autenticarse**
   ```bash
   clasp login
   ```

3. **Crear el contenedor**
   Opción A: crea una hoja en Google Sheets → *Extensiones → Apps Script*.
   Opción B: desde una carpeta con este repositorio:
   ```bash
   clasp create --title "Sistema PADI" --type sheets --parentId ID_DE_TU_HOJA
   ```

4. **Subir el código**
   Copia los archivos `*.js` y `appsscript.json` de este repositorio a la
   carpeta donde hiciste `clasp create` (o haz `clasp clone <scriptId>`),
   luego:
   ```bash
   clasp push
   ```

5. **Primer arranque**
   Abre la hoja de cálculo y recarga: aparecerá el menú del sistema.
   En la primera ejecución Google pedirá autorizar los permisos.

## Notas

- El archivo `appsscript.json` de este repo contiene el manifiesto (zonas
  horarias, scopes, menú web si aplica).
- Los datos se guardan en las hojas que el propio sistema genera; no hace
  falta crear nada a mano.
- Para desplegar interfaz web: *Implementar → Nueva implementación →
  Aplicación web*, según cómo tengas configurado el manifiesto.
