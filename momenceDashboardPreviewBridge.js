import fs from 'fs';
import http from 'http';
import path from 'path';
import { spawn } from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const HOST = process.env.SCHEDULE_PREVIEW_HOST || '127.0.0.1';
const PORT = Number(process.env.SCHEDULE_PREVIEW_PORT || 3210);
const COMMAND = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const COMMAND_ARGS = ['run', 'update', '--', '--static'];
const ALLOWED_FILES = new Set(['Kemps.html', 'Bandra.html']);
const LOG_LIMIT = 300;

const state = {
  status: 'idle',
  startedAt: null,
  finishedAt: null,
  exitCode: null,
  command: `${COMMAND} ${COMMAND_ARGS.join(' ')}`,
  logLines: [],
  currentProcess: null,
  runCount: 0,
  error: null
};

function log(message) {
  const line = `[${new Date().toISOString()}] ${message}`;
  state.logLines.push(line);
  if (state.logLines.length > LOG_LIMIT) {
    state.logLines.splice(0, state.logLines.length - LOG_LIMIT);
  }
  console.log(message);
}

function getFileInfo(fileName) {
  const filePath = path.join(__dirname, fileName);

  if (!fs.existsSync(filePath)) {
    return {
      name: fileName,
      exists: false,
      path: filePath,
      updatedAt: null,
      size: 0
    };
  }

  const stats = fs.statSync(filePath);
  return {
    name: fileName,
    exists: true,
    path: filePath,
    updatedAt: stats.mtime.toISOString(),
    size: stats.size
  };
}

function getStatusPayload() {
  return {
    ok: true,
    status: state.status,
    startedAt: state.startedAt,
    finishedAt: state.finishedAt,
    exitCode: state.exitCode,
    command: state.command,
    runCount: state.runCount,
    error: state.error,
    files: {
      kemps: getFileInfo('Kemps.html'),
      bandra: getFileInfo('Bandra.html')
    },
    logs: state.logLines.slice(-80)
  };
}

function writeJson(res, statusCode, payload) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(payload, null, 2));
}

function writeText(res, statusCode, body, contentType = 'text/plain; charset=utf-8') {
  res.writeHead(statusCode, {
    'Content-Type': contentType,
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      if (chunks.length === 0) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch (error) {
        reject(error);
      }
    });
    req.on('error', reject);
  });
}

function startUpdateRun() {
  if (state.currentProcess) {
    return { started: false, reason: 'already-running' };
  }

  state.status = 'running';
  state.startedAt = new Date().toISOString();
  state.finishedAt = null;
  state.exitCode = null;
  state.error = null;
  state.runCount += 1;
  state.logLines = [];

  log(`Starting background update: ${state.command}`);

  const child = spawn(COMMAND, COMMAND_ARGS, {
    cwd: __dirname,
    env: process.env,
    stdio: ['ignore', 'pipe', 'pipe']
  });

  state.currentProcess = child;

  child.stdout.on('data', chunk => {
    const text = chunk.toString('utf8').trimEnd();
    if (text) {
      text.split(/\r?\n/).forEach(line => log(line));
    }
  });

  child.stderr.on('data', chunk => {
    const text = chunk.toString('utf8').trimEnd();
    if (text) {
      text.split(/\r?\n/).forEach(line => log(`stderr: ${line}`));
    }
  });

  child.on('error', error => {
    state.status = 'error';
    state.finishedAt = new Date().toISOString();
    state.error = error.message;
    state.currentProcess = null;
    log(`Update process failed to start: ${error.message}`);
  });

  child.on('close', code => {
    state.exitCode = code;
    state.finishedAt = new Date().toISOString();
    state.status = code === 0 ? 'success' : 'error';
    state.error = code === 0 ? null : `Update exited with code ${code}`;
    state.currentProcess = null;
    log(code === 0 ? 'Background update completed successfully.' : `Background update failed with exit code ${code}.`);
  });

  return { started: true };
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url || '/', `http://${req.headers.host || `${HOST}:${PORT}`}`);

  if (req.method === 'OPTIONS') {
    writeText(res, 204, '');
    return;
  }

  if (req.method === 'GET' && url.pathname === '/') {
    writeJson(res, 200, {
      ok: true,
      message: 'Momence schedule preview bridge is running.',
      command: state.command,
      endpoints: ['/status', '/run', '/html/Kemps.html', '/html/Bandra.html']
    });
    return;
  }

  if (req.method === 'GET' && url.pathname === '/status') {
    writeJson(res, 200, getStatusPayload());
    return;
  }

  if (req.method === 'POST' && url.pathname === '/run') {
    try {
      await readBody(req);
    } catch {
      writeJson(res, 400, {
        ok: false,
        error: 'Request body must be valid JSON.'
      });
      return;
    }

    const result = startUpdateRun();
    const payload = getStatusPayload();
    payload.started = result.started;
    payload.reason = result.reason || null;
    writeJson(res, result.started ? 202 : 200, payload);
    return;
  }

  if (req.method === 'GET' && url.pathname.startsWith('/html/')) {
    const fileName = decodeURIComponent(url.pathname.replace('/html/', ''));

    if (!ALLOWED_FILES.has(fileName)) {
      writeJson(res, 404, { ok: false, error: 'Unknown preview file.' });
      return;
    }

    const filePath = path.join(__dirname, fileName);
    if (!fs.existsSync(filePath)) {
      writeJson(res, 404, { ok: false, error: `${fileName} does not exist yet.` });
      return;
    }

    writeText(res, 200, fs.readFileSync(filePath, 'utf8'), 'text/html; charset=utf-8');
    return;
  }

  writeJson(res, 404, { ok: false, error: 'Not found.' });
});

server.listen(PORT, HOST, () => {
  log(`Momence schedule preview bridge listening on http://${HOST}:${PORT}`);
  log(`Ready to run: ${state.command}`);
});
