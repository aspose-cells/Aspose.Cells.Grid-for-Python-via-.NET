import { spawn, spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import process from 'node:process';

const windows = process.platform === 'win32';
const venvPython = windows ? '.venv\\Scripts\\python.exe' : '.venv/bin/python';

function findSystemPython() {
  const candidates = windows ? [['py', ['-3']], ['python', []]] : [['python3', []], ['python', []]];
  for (const [command, prefix] of candidates) {
    const probe = spawnSync(command, [...prefix, '--version'], { stdio: 'ignore' });
    if (probe.status === 0) return { command, prefix };
  }
  throw new Error('Python 3 was not found. Install Python 3.11 or newer.');
}

function runSync(command, args) {
  const result = spawnSync(command, args, { stdio: 'inherit' });
  if (result.status !== 0) process.exit(result.status ?? 1);
}

function ensureEnvironment() {
  if (!existsSync(venvPython)) {
    const systemPython = findSystemPython();
    runSync(systemPython.command, [...systemPython.prefix, '-m', 'venv', '.venv']);
  }
}

const action = process.argv[2] ?? 'run';

if (action === 'setup') {
  ensureEnvironment();
  runSync(venvPython, ['-m', 'pip', 'install', '--upgrade', 'pip']);
  runSync(venvPython, ['-m', 'pip', 'install', '-r', 'requirements.txt']);
  process.exit(0);
}

ensureEnvironment();

if (action === 'check') {
  runSync(venvPython, ['-m', 'py_compile', 'app.py']);
  process.exit(0);
}

if (action !== 'run') {
  throw new Error(`Unknown Python action: ${action}`);
}

const child = spawn(venvPython, ['app.py'], { stdio: 'inherit' });
for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => child.kill(signal));
}
child.on('exit', (code, signal) => {
  if (signal) process.kill(process.pid, signal);
  process.exit(code ?? 1);
});
