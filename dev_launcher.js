/**
 * dev_launcher.js — Unified Process Orchestrator for Antigravity Office Studio.
 * Spawns both the Python Web API Bridge (port 8000) and the Vite Web App (port 5173).
 */

import fs from 'fs';
import path from 'path';
import net from 'net';
import { fileURLToPath } from 'url';
import { spawn, spawnSync, execSync } from 'child_process';
import process from 'process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
let pkgVersion = '4.6.0';
try {
  const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
  if (pkg && pkg.version) {
    pkgVersion = pkg.version;
  }
} catch {
  // Use fallback if package.json read fails
}

console.log('\x1b[36m%s\x1b[0m', '═══════════════════════════════════════════════════════════════════════════');
console.log('\x1b[1m\x1b[34m%s\x1b[0m', `  🚀 Antigravity Universal Office & Presentation Web Studio (v${pkgVersion})`);
console.log('\x1b[36m%s\x1b[0m', '═══════════════════════════════════════════════════════════════════════════');
console.log('\x1b[90m%s\x1b[0m', '  Starting local development services (Press Ctrl+C to stop all cleanly)...');

let isShuttingDown = false;
let apiProcess = null;
let viteProcess = null;

// Wait for a port to be listening before proceeding (prevents ECONNREFUSED race conditions)
function waitForPort(port, host = '127.0.0.1', timeoutMs = 12000) {
  const startTime = Date.now();
  return new Promise((resolve) => {
    const tryConnect = () => {
      if (isShuttingDown) return resolve(false);
      const socket = new net.Socket();
      socket.setTimeout(400);

      socket.once('connect', () => {
        socket.destroy();
        resolve(true);
      });

      socket.once('error', () => {
        socket.destroy();
        if (Date.now() - startTime > timeoutMs) {
          resolve(false);
        } else {
          setTimeout(tryConnect, 150);
        }
      });

      socket.once('timeout', () => {
        socket.destroy();
        if (Date.now() - startTime > timeoutMs) {
          resolve(false);
        } else {
          setTimeout(tryConnect, 150);
        }
      });

      socket.connect(port, host);
    };

    tryConnect();
  });
}

// Clean up any lingering process on Windows listening on a port
function cleanupPort(port) {
  if (process.platform !== 'win32') return;
  try {
    const stdout = execSync(`netstat -ano | findstr :${port}`, {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    const lines = stdout.trim().split('\n');
    const killedPids = new Set();
    for (const line of lines) {
      const parts = line.trim().split(/\s+/);
      const pid = parts[parts.length - 1];
      if (pid && pid !== '0' && pid !== String(process.pid) && !killedPids.has(pid)) {
        killedPids.add(pid);
        spawnSync('taskkill', ['/pid', pid, '/f', '/t'], { stdio: 'ignore' });
      }
    }
  } catch {
    // Port not in use
  }
}

function killProcessTree(pid) {
  if (!pid) return;
  try {
    if (process.platform === 'win32') {
      spawnSync('taskkill', ['/pid', String(pid), '/f', '/t'], { stdio: 'ignore' });
    } else {
      process.kill(-pid, 'SIGKILL');
    }
  } catch {
    // Process might have already terminated
  }
}

// 1. Spawn Python Web API Server with auto-restart on exit
const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

function startApiServer() {
  if (isShuttingDown) return;
  apiProcess = spawn(pythonCmd, ['web_api_server.py', '8000'], {
    stdio: ['ignore', 'inherit', 'inherit'],
    shell: false,
  });

  apiProcess.on('error', (err) => {
    if (!isShuttingDown) {
      console.error('\x1b[31m[API Server Error]\x1b[0m', err.message);
    }
  });

  apiProcess.on('exit', (code) => {
    if (!isShuttingDown) {
      console.log('\x1b[33m%s\x1b[0m', `[API Server] Process exited (code ${code}), restarting in 1s...`);
      setTimeout(startApiServer, 1000);
    }
  });
}

// 2. Spawn Vite Dev Server directly via Node (No cmd.exe / npx.cmd batch wrapper)
// Pipe stdin instead of inheriting so Vite never puts the parent terminal into raw mode or steals Ctrl+C.
function startViteServer() {
  if (isShuttingDown) return;
  const viteBin = path.join(__dirname, 'node_modules', 'vite', 'bin', 'vite.js');
  if (fs.existsSync(viteBin)) {
    viteProcess = spawn(process.execPath, [viteBin], {
      stdio: ['pipe', 'inherit', 'inherit'],
      shell: false,
    });
  } else {
    viteProcess = spawn('npx', ['vite'], {
      stdio: ['pipe', 'inherit', 'inherit'],
      shell: process.platform === 'win32',
    });
  }

  viteProcess.on('error', (err) => {
    if (!isShuttingDown) {
      console.error('\x1b[31m[Vite Error]\x1b[0m', err.message);
    }
  });

  viteProcess.on('exit', (code) => {
    if (!isShuttingDown) {
      cleanExit();
    }
  });
}

// Graceful, one-touch cleanup on exit
const cleanExit = () => {
  if (isShuttingDown) return;
  isShuttingDown = true;
  try {
    if (process.stdin.isTTY && typeof process.stdin.setRawMode === 'function') {
      process.stdin.setRawMode(false);
    }
    process.stdin.pause();
  } catch {
    // Ignore stdin restore errors
  }
  console.log('\n\x1b[33m%s\x1b[0m', '  [Shutting down] Closing API Bridge and Vite servers...');
  try {
    if (apiProcess && apiProcess.pid) killProcessTree(apiProcess.pid);
    if (viteProcess && viteProcess.pid) killProcessTree(viteProcess.pid);

    cleanupPort(8000);
    cleanupPort(5173);
  } catch (e) {
    // Ignore cleanup errors
  }
  console.log('\x1b[32m%s\x1b[0m', '  ✓ All services stopped cleanly. Ready.\n');
  process.exit(0);
};

// Exclusive Terminal Keypress & Signal Handler:
// Keeps full ownership of stdin so Ctrl+C (\x03) or 'q' triggers instant cleanExit(),
// while forwarding other shortcuts ('h', 'r', 'u', 'c', 'o') to Vite's stdin pipe.
if (process.stdin.isTTY) {
  process.stdin.setRawMode(true);
  process.stdin.resume();
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', (key) => {
    // \u0003 is Ctrl+C, \u0004 is Ctrl+D
    if (key === '\u0003' || key === '\u0004' || key === 'q' || key === 'Q') {
      cleanExit();
      return;
    }
    if (viteProcess && viteProcess.stdin && !viteProcess.stdin.destroyed) {
      viteProcess.stdin.write(key);
    }
  });
}

process.on('SIGINT', cleanExit);
process.on('SIGTERM', cleanExit);
process.on('SIGBREAK', cleanExit);

// Coordinated bootstrap: start Python API Bridge first, wait until port 8000 is listening, then start Vite
async function bootstrap() {
  startApiServer();
  await waitForPort(8000, '127.0.0.1', 12000);
  if (!isShuttingDown) {
    startViteServer();
  }
}

bootstrap();
