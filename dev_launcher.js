/**
 * dev_launcher.js — Unified Process Orchestrator for Antigravity Office Studio.
 * Spawns both the Python Web API Bridge (port 8000) and the Vite Web App (port 5173).
 */

import { spawn } from 'child_process';
import process from 'process';

console.log('\x1b[36m%s\x1b[0m', '═══════════════════════════════════════════════════════════════════════════');
console.log('\x1b[1m\x1b[34m%s\x1b[0m', '  🚀 Antigravity Universal Office & Presentation Web Studio (v3.2)');
console.log('\x1b[36m%s\x1b[0m', '═══════════════════════════════════════════════════════════════════════════');
console.log('\x1b[90m%s\x1b[0m', '  Starting local development services...');

let isShuttingDown = false;
let apiProcess = null;

// 1. Spawn Python Web API Server with auto-restart on exit
const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

function startApiServer() {
  if (isShuttingDown) return;
  apiProcess = spawn(pythonCmd, ['web_api_server.py', '8000'], {
    stdio: 'inherit',
    shell: true,
  });

  apiProcess.on('error', (err) => {
    console.error('\x1b[31m[API Server Error]\x1b[0m', err.message);
  });

  apiProcess.on('exit', (code) => {
    if (!isShuttingDown) {
      console.log('\x1b[33m%s\x1b[0m', `[API Server] Process exited (code ${code}), restarting in 1s...`);
      setTimeout(startApiServer, 1000);
    }
  });
}

startApiServer();

// 2. Spawn Vite Dev Server
const npxCmd = process.platform === 'win32' ? 'npx.cmd' : 'npx';
const viteProcess = spawn(npxCmd, ['vite'], {
  stdio: 'inherit',
  shell: true,
});

viteProcess.on('error', (err) => {
  console.error('\x1b[31m[Vite Error]\x1b[0m', err.message);
});

// Graceful cleanup on exit
const cleanExit = () => {
  isShuttingDown = true;
  console.log('\n\x1b[33m%s\x1b[0m', '  [Shutting down] Closing API Bridge and Vite servers...');
  try {
    if (process.platform === 'win32') {
      if (apiProcess && apiProcess.pid) spawn('taskkill', ['/pid', apiProcess.pid, '/f', '/t']);
      if (viteProcess && viteProcess.pid) spawn('taskkill', ['/pid', viteProcess.pid, '/f', '/t']);
    } else {
      if (apiProcess) apiProcess.kill('SIGTERM');
      if (viteProcess) viteProcess.kill('SIGTERM');
    }
  } catch (e) {
    // Ignore cleanup errors
  }
  process.exit(0);
};

process.on('SIGINT', cleanExit);
process.on('SIGTERM', cleanExit);
