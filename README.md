# AutorIA Corrector para Word

Prototipo de complemento lateral para Microsoft Word. Incluye identidad visual old tech, analisis de voz, correcciones editoriales, ritmo narrativo, asistente de reformulacion y panel de sesion.

## Probar en navegador

```powershell
npm run dev
```

Luego abre:

```text
http://localhost:3000
```

## Usar como complemento de Word

1. Levanta el servidor local con `npm run dev`.
2. Para sideload en Word, usa `manifest.xml`.
3. El manifiesto apunta a `https://localhost:3000/src/taskpane.html`, que es el formato esperado por Office Add-ins. Para una demo en navegador, el servidor local tambien responde en `http://localhost:3000`.

## Publicar con GitHub Pages

1. Crea un repositorio en GitHub, por ejemplo `autoria-corrector-word`.
2. Sube estos archivos al repositorio.
3. En GitHub, entra a `Settings > Pages`.
4. En `Build and deployment`, selecciona `Deploy from a branch`.
5. Elige la rama `main` y la carpeta `/root`.
6. GitHub publicara el panel en una URL como:

```text
https://USUARIO.github.io/REPO/src/taskpane.html
```

Para crear el manifiesto final de Word:

```powershell
npm run manifest:github -- USUARIO REPO
```

Eso genera `manifest.github-pages.xml`, que es el archivo que debes cargar en Word.

`manifest.xml` queda apuntando a `localhost` para pruebas locales. `manifest.github-pages.template.xml` queda preparado para GitHub Pages.

## Alcance del prototipo

- Modulo 1: correccion editorial contextual simulada con reglas locales.
- Modulo 2: visualizador de ritmo, palabras repetidas y tension lexica.
- Modulo 3: texto fantasma y tres variantes de reformulacion sin llamar a un LLM.
- Modulo 4: contador de palabras nuevas, tiempo activo y resumen de sesion.

La base ya esta preparada para conectar un backend Node/Express y una API de IA cuando exista una clave propia del proyecto.
