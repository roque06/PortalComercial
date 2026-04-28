# Validador de Cedulas

App local para:

- seleccionar un Excel con columna `Cedula`
- elegir la instancia Oracle desde `dbConnections.json`
- consultar `CLIENTES_B2000` con `CODIGO_TIPO_IDENTIFICACION = 1`
- exportar un Excel con las cedulas no encontradas

## Abrir la app

```powershell
npm run cedulas:gui
```

o ejecutar:

```powershell
scripts\run-cedulas-gui.cmd
```

## Requisitos

- `dbConnections.json` disponible en `C:\CierreBancarioQA\src\config\dbConnections.json`
- `sqlplus.exe` disponible en `C:\instantclient_21_13\sqlplus.exe`
- Node.js disponible en `PATH`

## Salida

La app genera un Excel con:

- hoja `No_Existen`
- hoja `Resumen`

La comparación normaliza las cédulas removiendo guiones y otros caracteres no numéricos, para evitar falsos no encontrados por diferencias de formato.
