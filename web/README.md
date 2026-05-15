# EditorHTML PCF Local

Harness web para probar el editor sin importar el PCF en Power Apps en cada iteracion.

## Ejecutar

```powershell
npm.cmd install
npm.cmd run dev:web
```

Abrir `http://127.0.0.1:5173`.

## Paridad con PCF

- Reutiliza `EditorHTML/lifecycle/EditorComponent.ts`.
- Reutiliza `EditorHTML/Orquestador/Paginator.ts`.
- Reutiliza `EditorHTML/Orquestador/CaretManager.ts`.
- Reutiliza `EditorHTML/Resize/Toolbar.ts`.
- Reutiliza `EditorHTML/css/editor.css`.

La unica diferencia intencionada es el origen del HTML:

- En PCF se carga y guarda con `EditorHTML/execCommand/fileApi.ts` contra Dataverse.
- En local se carga un archivo `.html` desde el navegador y se guarda en `localStorage`; tambien se puede descargar el HTML resultante.

Esto permite reproducir aqui los mismos problemas de DOM, paginacion, caret, tablas e imagenes que luego afectan al PCF.
