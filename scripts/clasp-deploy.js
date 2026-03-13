const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

function parseArgs(argv) {
  const parsed = {
    deploymentId: process.env.CLASP_DEPLOYMENT_ID || '',
    description: process.env.CLASP_DEPLOY_DESCRIPTION || '',
  };
  const bareArgs = [];

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if ((arg === '--deploymentId' || arg === '-i') && i + 1 < argv.length) {
      parsed.deploymentId = argv[++i];
      continue;
    }
    if ((arg === '--description' || arg === '-d') && i + 1 < argv.length) {
      parsed.description = argv[++i];
      continue;
    }
    bareArgs.push(arg);
  }

  if (!parsed.deploymentId && bareArgs.length >= 2 && looksLikeDeploymentId(bareArgs[0])) {
    parsed.deploymentId = bareArgs.shift();
  }

  if (!parsed.description && bareArgs.length > 0) {
    parsed.description = bareArgs.join(' ');
  }

  return parsed;
}

function looksLikeDeploymentId(value) {
  return /^[A-Za-z0-9_-]{10,}$/.test(String(value || '').trim());
}

function getClaspEntrypoint() {
  if (process.platform === 'win32') {
    const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || '', 'AppData', 'Roaming');
    const candidate = path.join(appData, 'npm', 'node_modules', '@google', 'clasp', 'build', 'src', 'index.js');
    if (fs.existsSync(candidate)) return candidate;
  }

  const npmBin = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const npmRoot = spawnSync(npmBin, ['root', '-g'], {
    encoding: 'utf8',
    cwd: process.cwd(),
    shell: process.platform === 'win32',
  });

  if (npmRoot.status === 0) {
    const rootDir = String(npmRoot.stdout || '').trim();
    const candidate = path.join(rootDir, '@google', 'clasp', 'build', 'src', 'index.js');
    if (fs.existsSync(candidate)) return candidate;
  }

  return '';
}

const parsed = parseArgs(process.argv.slice(2));
const claspEntrypoint = getClaspEntrypoint();
const claspArgs = [];

if (claspEntrypoint) {
  claspArgs.push(claspEntrypoint);
} else {
  console.error('Could not locate the global clasp installation.');
  process.exit(1);
}

claspArgs.push('-P', 'src', 'deploy');

if (parsed.deploymentId) {
  claspArgs.push('-i', parsed.deploymentId);
}

if (parsed.description) {
  claspArgs.push('-d', parsed.description);
}

const result = spawnSync(process.execPath, claspArgs, {
  stdio: 'inherit',
  cwd: process.cwd(),
});

if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}

process.exit(result.status === null ? 1 : result.status);
