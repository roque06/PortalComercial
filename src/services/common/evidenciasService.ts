import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';

export function limpiarEvidenciasTemporales() {
  const evidenciasDir = path.resolve(process.cwd(), 'artifacts', 'evidencias_tmp');
  if (!fs.existsSync(evidenciasDir)) return;

  for (const file of fs.readdirSync(evidenciasDir)) {
    if (file.toLowerCase().endsWith('.png')) {
      fs.unlinkSync(path.join(evidenciasDir, file));
    }
  }
}

export async function abrirArchivoAlFinal(filePath: string) {
  if (!filePath || !fs.existsSync(filePath)) return false;

  try {
    if (process.platform === 'win32') {
      const child = spawn('cmd', ['/c', 'start', '', filePath], {
        detached: true,
        stdio: 'ignore',
      });
      child.unref();
      return true;
    }

    if (process.platform === 'darwin') {
      const child = spawn('open', [filePath], { detached: true, stdio: 'ignore' });
      child.unref();
      return true;
    }

    const child = spawn('xdg-open', [filePath], { detached: true, stdio: 'ignore' });
    child.unref();
    return true;
  } catch {
    return false;
  }
}
