import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { spawn } from 'child_process';
import path from 'path';

// Plugin Vite qui démarre le backend FastAPI automatiquement
function backendPlugin() {
  let backendProcess: ReturnType<typeof spawn> | null = null;

  return {
    name: 'backend-auto-start',
    configureServer() {
      const pythonRoot = path.resolve(__dirname, '..', '..', '..');
      const uvicorn = path.resolve(
        process.env.LOCALAPPDATA ?? '',
        'Programs', 'Python', 'Python314', 'Scripts', 'uvicorn.exe'
      );

      console.log('\x1b[36m[CADI]\x1b[0m Démarrage du backend FastAPI...');

      backendProcess = spawn(
        uvicorn,
        ['apps.cadi_web.back.main:app', '--reload', '--port', '8001'],
        {
          cwd: pythonRoot,
          env: { ...process.env, PYTHONPATH: pythonRoot },
          stdio: 'pipe',
          shell: false,
        }
      );

      backendProcess.stdout?.on('data', (data: Buffer) => {
        const lines = data.toString().split('\n').filter(Boolean);
        lines.forEach(line => console.log(`\x1b[36m[back]\x1b[0m ${line}`));
      });

      backendProcess.stderr?.on('data', (data: Buffer) => {
        const lines = data.toString().split('\n').filter(Boolean);
        lines.forEach(line => {
          // uvicorn écrit ses infos sur stderr — on filtre les vraies erreurs
          if (line.includes('ERROR') || line.includes('Traceback')) {
            console.error(`\x1b[31m[back ERR]\x1b[0m ${line}`);
          } else {
            console.log(`\x1b[36m[back]\x1b[0m ${line}`);
          }
        });
      });

      backendProcess.on('exit', (code) => {
        if (code !== 0 && code !== null) {
          console.error(`\x1b[31m[CADI]\x1b[0m Backend arrêté (code ${code})`);
        }
      });

      // Arrête le backend proprement quand Vite s'arrête
      process.on('exit', () => backendProcess?.kill());
      process.on('SIGINT', () => { backendProcess?.kill(); process.exit(); });
      process.on('SIGTERM', () => { backendProcess?.kill(); process.exit(); });
    },
  };
}

export default defineConfig({
  plugins: [
    react(),
    ...(process.env.CADI_AUTO_BACKEND === '1' ? [backendPlugin()] : []),
  ],
  optimizeDeps: {
    exclude: ['lucide-react'],
  },
});
