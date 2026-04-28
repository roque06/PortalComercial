# Runner Oficial

La unica app soportada para ejecutar automatizaciones es:

- `scripts/run-robot-gui.cmd`

## Uso

### Opcion A: doble clic

1. Abrir la carpeta `scripts`.
2. Ejecutar `run-robot-gui.cmd`.
3. Elegir el flujo desde la interfaz:
   - `Cuenta Efectivo - Cliente existente`
   - `Cuenta Efectivo - Cliente nuevo`
   - `Cuenta Efectivo - Ambos`
   - `Deposito a plazos - existente`

### Opcion B: consola

```powershell
npm run robot
```

o

```powershell
npm run robot:gui
```

## Notas

- Si no existe `node_modules`, el runner instala dependencias automaticamente.
- `run-robot-gui.cmd` delega en `run-robot-gui.ps1`; no hay otros launchers soportados.
