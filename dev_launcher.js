/**
 * dev_launcher.js — Unified Process Orchestrator for Antigravity Office Studio.
 * Spawns both the Python Web API Bridge (port 8000) and the Vite Web App (port 5173).
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawn, spawnSync, execSync } from 'child_process';
import process from 'process';
import readline from 'readline';

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
    stdio: 'inherit',
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

startApiServer();

// 2. Spawn Vite Dev Server directly via Node (No cmd.exe / npx.cmd batch wrapper)
// This eliminates the Windows "Terminate batch job (Y/N)?" prompt entirely.
const viteBin = path.join(__dirname, 'node_modules', 'vite', 'bin', 'vite.js');
if (fs.existsSync(viteBin)) {
  viteProcess = spawn(process.execPath, [viteBin], {
    stdio: 'inherit',
    shell: false,
  });
} else {
  viteProcess = spawn('npx', ['vite'], {
    stdio: 'inherit',
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

// Graceful, one-touch cleanup on exit
const cleanExit = () => {
  if (isShuttingDown) return;
  isShuttingDown = true;
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

// Windows SIGINT Hook via Readline to intercept Ctrl+C cleanly before OS batch break handler
if (process.platform === 'win32' && process.stdin.isTTY) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.on('SIGINT', () => {
    cleanExit();
  });
}

process.on('SIGINT', cleanExit);
process.on('SIGTERM', cleanExit);
