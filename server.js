const http = require('http');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const PORT = Number(process.env.PORT || 4141);
const HOST = process.env.HOST || '127.0.0.1';
const jobs = new Map();

function sendJson(res, status, body) {
  const payload = JSON.stringify(body, null, 2);
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
  });
  res.end(payload);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => {
      data += chunk;
      if (data.length > 2_000_000) {
        reject(new Error('Request body too large'));
        req.destroy();
      }
    });
    req.on('end', () => {
      if (!data.trim()) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(data));
      } catch (error) {
        reject(new Error(`Invalid JSON: ${error.message}`));
      }
    });
    req.on('error', reject);
  });
}

function runPac(args, onLine) {
  return new Promise((resolve) => {
    const child = spawn(process.env.PAC_PATH || 'pac.exe', args, {
      cwd: process.cwd(),
      shell: false,
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';

    function collect(chunk, target) {
      const text = chunk.toString();
      if (target === 'stdout') stdout += text;
      if (target === 'stderr') stderr += text;
      if (onLine) {
        text.split(/\r?\n/).filter(Boolean).forEach(line => onLine(line));
      }
    }

    child.stdout.on('data', chunk => collect(chunk, 'stdout'));
    child.stderr.on('data', chunk => collect(chunk, 'stderr'));
    child.on('error', error => {
      resolve({ code: 1, stdout, stderr: `${stderr}\n${error.message}`.trim() });
    });
    child.on('close', code => {
      resolve({ code: code ?? 1, stdout, stderr });
    });
  });
}

function runFolderPicker(initialPath) {
  const script = [
    'Add-Type -AssemblyName System.Windows.Forms',
    '$dialog = New-Object System.Windows.Forms.FolderBrowserDialog',
    "$dialog.Description = 'Select export folder'",
    '$dialog.ShowNewFolderButton = $true',
    '$initialPath = $env:PAC_EXPORTER_INITIAL_DIR',
    'if ($initialPath -and (Test-Path -LiteralPath $initialPath)) { $dialog.SelectedPath = $initialPath }',
    'if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {',
    '  [Console]::Out.WriteLine($dialog.SelectedPath)',
    '  exit 0',
    '}',
    'exit 2',
  ].join('; ');

  return new Promise((resolve) => {
    const child = spawn('powershell.exe', ['-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-Command', script], {
      cwd: process.cwd(),
      shell: false,
      windowsHide: false,
      env: {
        ...process.env,
        PAC_EXPORTER_INITIAL_DIR: typeof initialPath === 'string' ? initialPath.trim() : '',
      },
    });

    let stdout = '';
    let stderr = '';
    child.stdout.on('data', chunk => { stdout += chunk.toString(); });
    child.stderr.on('data', chunk => { stderr += chunk.toString(); });
    child.on('error', error => {
      resolve({ code: 1, stdout, stderr: `${stderr}\n${error.message}`.trim() });
    });
    child.on('close', code => {
      resolve({ code: code ?? 1, stdout, stderr });
    });
  });
}

function buildEnvironmentArgs(environment) {
  const value = typeof environment === 'string' ? environment.trim() : '';
  return value ? ['--environment', value] : [];
}

function parsePacTable(output, columns) {
  const lines = String(output || '').split(/\r?\n/);
  const headerIndex = lines.findIndex(line => columns.every(column => line.includes(column)));
  if (headerIndex === -1) {
    return [];
  }

  const header = lines[headerIndex];
  const positions = columns
    .map(column => ({ column, at: header.indexOf(column) }))
    .sort((a, b) => a.at - b.at);

  return lines.slice(headerIndex + 1)
    .map(line => {
      if (!line.trim()) return null;
      const row = {};
      positions.forEach((position, index) => {
        const next = positions[index + 1];
        row[position.column] = line.slice(position.at, next ? next.at : undefined).trim();
      });
      const first = row[positions[0].column];
      if (!first || first.startsWith('---')) return null;
      return row;
    })
    .filter(Boolean);
}

function parseSolutionList(output) {
  const lines = output.split(/\r?\n/);
  const headerIndex = lines.findIndex(line =>
    line.includes('Unique Name') &&
    line.includes('Friendly Name') &&
    line.includes('Version') &&
    line.includes('Managed')
  );

  if (headerIndex === -1) {
    return [];
  }

  const header = lines[headerIndex];
  const uniqueAt = header.indexOf('Unique Name');
  const friendlyAt = header.indexOf('Friendly Name');
  const versionAt = header.indexOf('Version');
  const managedAt = header.indexOf('Managed');

  return lines.slice(headerIndex + 1)
    .map(line => {
      if (!line.trim()) return null;
      const uniqueName = line.slice(uniqueAt, friendlyAt).trim();
      const friendlyName = line.slice(friendlyAt, versionAt).trim();
      const version = line.slice(versionAt, managedAt).trim();
      const managedRaw = line.slice(managedAt).trim();
      if (!uniqueName || uniqueName.startsWith('---')) return null;
      return {
        uniqueName,
        friendlyName,
        version,
        managed: /^true$/i.test(managedRaw),
        system: isLikelySystemSolution(uniqueName, friendlyName),
      };
    })
    .filter(Boolean);
}

function parseSolutionMetadata(output) {
  const rows = parsePacTable(output, [
    'friendlyname',
    'modifiedon',
    'ismanaged',
    'uniquename',
    'modifiedby',
    'version',
    'solutionid',
  ]);

  const metadata = new Map();
  rows.forEach(row => {
    if (!row.uniquename) return;
    metadata.set(row.uniquename, {
      friendlyName: row.friendlyname,
      modifiedOn: row.modifiedon,
      modifiedBy: row.modifiedby,
      version: row.version,
      managed: /^true$/i.test(row.ismanaged),
    });
  });
  return metadata;
}

function parseAuthProfiles(output) {
  const emails = [...String(output || '').matchAll(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi)]
    .map(match => match[0]);
  return {
    authenticated: emails.length > 0,
    user: emails[0] || '',
  };
}

function isLikelySystemSolution(uniqueName, friendlyName) {
  const text = `${uniqueName || ''} ${friendlyName || ''}`.toLowerCase();
  return (
    text.startsWith('msdyn_') ||
    text.includes('microsoft') ||
    text.includes('dataverse') ||
    text.includes('common data services default solution') ||
    text.includes('default solution') ||
    text.includes('active solution') ||
    text.includes('creator kit') ||
    text.includes('power cat') ||
    text.includes('sample data') ||
    text.includes('fundraiser')
  );
}

async function fetchSolutionMetadata(environment) {
  const fetchXml = [
    "<fetch count='5000'>",
    "  <entity name='solution'>",
    "    <attribute name='uniquename' />",
    "    <attribute name='friendlyname' />",
    "    <attribute name='version' />",
    "    <attribute name='ismanaged' />",
    "    <attribute name='modifiedon' />",
    "    <attribute name='modifiedby' />",
    "    <order attribute='modifiedon' descending='true' />",
    "  </entity>",
    '</fetch>',
  ].join('');

  const result = await runPac([
    'org',
    'fetch',
    ...buildEnvironmentArgs(environment),
    '--xml',
    fetchXml,
  ]);

  if (result.code !== 0) {
    return {
      ok: false,
      error: result.stderr || result.stdout || `pac exited with ${result.code}`,
      metadata: new Map(),
    };
  }

  return {
    ok: true,
    metadata: parseSolutionMetadata(result.stdout),
  };
}

function safeFileName(name) {
  return String(name).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').slice(0, 150);
}

function defaultExportDirectory() {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').slice(0, 19);
  return path.join(process.cwd(), 'exports', stamp);
}

async function listSolutions(body) {
  const args = [
    'solution',
    'list',
    ...buildEnvironmentArgs(body.environment),
  ];
  if (body.includeSystemSolutions) {
    args.push('--includeSystemSolutions');
  }

  const result = await runPac(args);
  if (result.code !== 0) {
    return {
      ok: false,
      error: result.stderr || result.stdout || `pac exited with ${result.code}`,
      raw: result.stdout,
    };
  }

  const metadataResult = await fetchSolutionMetadata(body.environment);
  const metadata = metadataResult.metadata;
  const solutions = parseSolutionList(result.stdout)
    .map(solution => {
      const item = metadata.get(solution.uniqueName);
      return {
        ...solution,
        friendlyName: solution.friendlyName || (item && item.friendlyName) || '',
        version: solution.version || (item && item.version) || '',
        modifiedOn: item ? item.modifiedOn : '',
        modifiedBy: item ? item.modifiedBy : '',
      };
    })
    .sort((a, b) => {
      const left = Date.parse(a.modifiedOn) || 0;
      const right = Date.parse(b.modifiedOn) || 0;
      return right - left;
    });

  return {
    ok: true,
    solutions,
    metadataWarning: metadataResult.ok ? '' : metadataResult.error,
    raw: result.stdout,
  };
}

async function listAuthProfiles() {
  const result = await runPac(['auth', 'list']);
  const auth = parseAuthProfiles(result.stdout || result.stderr || '');
  return {
    ok: result.code === 0,
    output: (result.stdout || result.stderr || '').trim(),
    authenticated: auth.authenticated,
    user: auth.user,
    error: result.code === 0 ? '' : (result.stderr || result.stdout || `pac exited with ${result.code}`),
  };
}

async function createAuthProfile(body) {
  const args = ['auth', 'create'];
  const environmentArgs = buildEnvironmentArgs(body.environment);
  if (environmentArgs.length) {
    args.push(...environmentArgs);
  }

  const result = await runPac(args);
  const auth = parseAuthProfiles(result.stdout || result.stderr || '');
  return {
    ok: result.code === 0,
    output: (result.stdout || result.stderr || '').trim(),
    authenticated: auth.authenticated,
    user: auth.user,
    error: result.code === 0 ? '' : (result.stderr || result.stdout || `pac exited with ${result.code}`),
  };
}

async function selectExportFolder(body) {
  const result = await runFolderPicker(body.initialPath);
  if (result.code === 2) {
    return { ok: true, cancelled: true, path: '' };
  }
  if (result.code !== 0) {
    return {
      ok: false,
      error: result.stderr || result.stdout || `folder picker exited with ${result.code}`,
    };
  }
  return {
    ok: true,
    path: result.stdout.trim(),
  };
}

async function runExportJob(job, options) {
  job.status = 'running';
  job.startedAt = new Date().toISOString();

  const environmentArgs = buildEnvironmentArgs(options.environment);
  const outputDir = options.outputDir && String(options.outputDir).trim()
    ? path.resolve(String(options.outputDir).trim())
    : defaultExportDirectory();

  job.outputDir = outputDir;
  fs.mkdirSync(outputDir, { recursive: true });

  const manifest = {
    jobId: job.id,
    startedAt: job.startedAt,
    environment: options.environment || '(active PAC environment)',
    outputDir,
    exports: [],
  };

  for (const name of options.solutionNames) {
    const zipPath = path.join(outputDir, `${safeFileName(name)}.zip`);
    const item = { name, zipPath, status: 'running', startedAt: new Date().toISOString() };
    manifest.exports.push(item);
    job.current = name;
    job.logs.push(`Exporting ${name} -> ${zipPath}`);

    const args = [
      'solution',
      'export',
      ...environmentArgs,
      '--name',
      name,
      '--path',
      zipPath,
      '--overwrite',
    ];

    const result = await runPac(args, line => job.logs.push(`[${name}] ${line}`));
    item.finishedAt = new Date().toISOString();
    item.exitCode = result.code;

    if (result.code === 0) {
      item.status = 'completed';
      job.completed += 1;
      job.logs.push(`Completed ${name}`);
    } else {
      item.status = 'failed';
      item.error = result.stderr || result.stdout || `pac exited with ${result.code}`;
      job.failed += 1;
      job.logs.push(`Failed ${name}: ${item.error}`);
    }
  }

  manifest.finishedAt = new Date().toISOString();
  manifest.status = job.failed ? 'completed_with_failures' : 'completed';
  fs.writeFileSync(path.join(outputDir, 'export-manifest.json'), JSON.stringify(manifest, null, 2));

  job.current = '';
  job.finishedAt = manifest.finishedAt;
  job.status = manifest.status;
  job.manifestPath = path.join(outputDir, 'export-manifest.json');
}

function createExportJob(body) {
  const solutionNames = Array.isArray(body.solutionNames)
    ? body.solutionNames.map(String).map(s => s.trim()).filter(Boolean)
    : [];

  if (!solutionNames.length) {
    return { ok: false, error: 'Select at least one solution to export.' };
  }

  const id = crypto.randomUUID();
  const job = {
    id,
    status: 'queued',
    current: '',
    logs: [],
    completed: 0,
    failed: 0,
    total: solutionNames.length,
    outputDir: '',
    manifestPath: '',
    createdAt: new Date().toISOString(),
  };
  jobs.set(id, job);

  runExportJob(job, {
    environment: body.environment,
    outputDir: body.outputDir,
    solutionNames,
  }).catch(error => {
    job.status = 'failed';
    job.finishedAt = new Date().toISOString();
    job.logs.push(error.stack || error.message);
  });

  return { ok: true, job };
}

function indexHtml() {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PAC Solution Exporter</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f4f6f8;
      --panel: #ffffff;
      --panel-2: #f8fafc;
      --panel-3: #eef3f8;
      --text: #16202c;
      --muted: #66758a;
      --line: #dce3ec;
      --accent: #245bd8;
      --accent-2: #1747ad;
      --accent-soft: #eef4ff;
      --danger: #b42318;
      --ok: #147a46;
      --warning: #a15c00;
      --shadow: 0 16px 40px rgba(21, 33, 51, .08);
      --focus: 0 0 0 3px rgba(36, 91, 216, .18);
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Segoe UI, Arial, sans-serif;
      font-size: 14px;
      letter-spacing: 0;
    }

    .app-shell {
      min-height: 100vh;
      transition: filter .18s ease, opacity .18s ease;
    }

    body.auth-checking .app-shell {
      filter: blur(5px);
      opacity: .55;
      pointer-events: none;
      user-select: none;
    }

    .auth-gate {
      position: fixed;
      inset: 0;
      z-index: 20;
      display: grid;
      place-items: center;
      background: rgba(244, 246, 248, .45);
      backdrop-filter: blur(2px);
    }

    .auth-gate[hidden] { display: none; }

    .auth-gate-panel {
      width: min(420px, calc(100vw - 32px));
      background: #ffffff;
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: 0 24px 80px rgba(21, 33, 51, .18);
      padding: 20px;
    }

    .auth-gate-title {
      font-size: 16px;
      font-weight: 700;
      margin-bottom: 6px;
    }

    .auth-gate-copy {
      color: var(--muted);
      line-height: 1.45;
    }

    .progress-bar {
      height: 4px;
      overflow: hidden;
      background: var(--panel-3);
      border-radius: 999px;
      margin-top: 16px;
    }

    .progress-bar::before {
      content: "";
      display: block;
      width: 38%;
      height: 100%;
      border-radius: inherit;
      background: var(--accent);
      animation: auth-progress 1s ease-in-out infinite;
    }

    @keyframes auth-progress {
      0% { transform: translateX(-105%); }
      100% { transform: translateX(275%); }
    }

    header {
      background: #ffffff;
      border-bottom: 1px solid var(--line);
      padding: 16px 24px;
      position: sticky;
      top: 0;
      z-index: 3;
      box-shadow: 0 1px 0 rgba(22, 32, 44, .02);
    }

    .header-row {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: flex-start;
    }

    .header-actions {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 10px;
      flex-wrap: wrap;
    }

    .setup-link {
      display: inline-flex;
      align-items: center;
      min-height: 30px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 6px 10px;
      background: #fff;
      color: var(--text);
      font-size: 12px;
      font-weight: 700;
      text-decoration: none;
    }

    .setup-link:hover {
      background: var(--panel-2);
      border-color: #cbd5e1;
    }

    h1 {
      margin: 0;
      font-size: 20px;
      line-height: 1.2;
      font-weight: 700;
    }

    .subtitle {
      margin-top: 5px;
      color: var(--muted);
      font-size: 13px;
    }

    .auth-chip {
      flex: 0 1 auto;
      max-width: 360px;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 6px 10px;
      color: var(--muted);
      background: #f8fafc;
      font-size: 12px;
      font-weight: 700;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .auth-chip.signed-in {
      color: var(--ok);
      background: #e9f7ef;
      border-color: #bce5cc;
    }

    main {
      max-width: 1500px;
      margin: 0 auto;
      padding: 20px;
      display: grid;
      grid-template-columns: minmax(0, 1fr) 330px;
      gap: 20px;
    }

    section, aside {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }

    .toolbar, .side-block {
      padding: 16px;
      border-bottom: 1px solid var(--line);
    }

    .toolbar {
      display: grid;
      grid-template-columns: minmax(260px, 1fr) 150px auto;
      gap: 12px;
      align-items: end;
      background: linear-gradient(180deg, #fff, #fbfcfe);
    }

    label {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 600;
      margin-bottom: 6px;
    }

    input, textarea {
      width: 100%;
      min-height: 36px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font: inherit;
      color: var(--text);
      background: #fff;
    }

    input:focus, textarea:focus, button:focus-visible {
      outline: none;
      border-color: var(--accent);
      box-shadow: var(--focus);
    }

    textarea {
      min-height: 70px;
      resize: vertical;
    }

    button {
      min-height: 36px;
      border: 1px solid transparent;
      border-radius: 6px;
      background: var(--accent);
      color: #fff;
      font: inherit;
      font-weight: 700;
      padding: 8px 12px;
      cursor: pointer;
      transition: background .12s ease, border-color .12s ease, box-shadow .12s ease, transform .12s ease;
    }

    button:hover { background: var(--accent-2); }
    button:active { transform: translateY(1px); }
    button.secondary {
      background: #fff;
      color: var(--text);
      border-color: var(--line);
    }
    button.secondary:hover { background: var(--panel-2); }
    button:disabled { opacity: .55; cursor: not-allowed; }
    button.loading {
      position: relative;
      padding-left: 34px;
      opacity: .9;
    }
    button.loading::before {
      content: "";
      position: absolute;
      left: 12px;
      top: 50%;
      width: 14px;
      height: 14px;
      margin-top: -7px;
      border: 2px solid rgba(255, 255, 255, .45);
      border-top-color: #ffffff;
      border-radius: 50%;
      animation: spin .7s linear infinite;
    }
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
    input[type="checkbox"] {
      width: 16px;
      min-height: 16px;
      height: 16px;
      accent-color: var(--accent);
      cursor: pointer;
    }

    .solution-list {
      position: relative;
      padding: 0;
      max-height: calc(100vh - 230px);
      overflow: auto;
      contain: content;
    }

    .solution-progress {
      display: none;
      align-items: center;
      gap: 10px;
      padding: 10px 16px;
      border-bottom: 1px solid var(--line);
      background: #f7fbff;
      color: var(--accent-2);
      font-size: 13px;
      font-weight: 700;
    }

    .solution-progress::before {
      content: "";
      width: 16px;
      height: 16px;
      border: 2px solid rgba(36, 91, 216, .22);
      border-top-color: var(--accent);
      border-radius: 50%;
      animation: spin .7s linear infinite;
      flex: 0 0 auto;
    }

    .solution-progress-bar {
      position: absolute;
      left: 0;
      right: 0;
      top: 0;
      height: 3px;
      overflow: hidden;
      background: transparent;
      z-index: 4;
      display: none;
    }

    .solution-progress-bar::before {
      content: "";
      display: block;
      width: 34%;
      height: 100%;
      background: var(--accent);
      border-radius: 999px;
      animation: table-progress 1s ease-in-out infinite;
    }

    @keyframes table-progress {
      0% { transform: translateX(-110%); }
      100% { transform: translateX(310%); }
    }

    .solution-list.loading .solution-progress,
    .solution-list.loading .solution-progress-bar {
      display: flex;
    }

    .solution-list.loading table {
      opacity: .55;
      pointer-events: none;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }

    th, td {
      padding: 11px 12px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: middle;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    th {
      background: var(--panel-2);
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      position: sticky;
      top: 0;
      z-index: 2;
    }

    td.num { width: 50px; color: var(--muted); font-variant-numeric: tabular-nums; }
    td.select, th.select { width: 48px; text-align: center; }
    td.managed { width: 112px; }
    tbody tr { cursor: pointer; }
    tbody tr:hover { background: #fbfdff; }
    .solution-cell {
      position: relative;
      min-width: 320px;
    }
    .name {
      font-weight: 700;
      overflow-wrap: anywhere;
      word-break: break-word;
      line-height: 1.35;
    }
    .friendly {
      color: var(--muted);
      margin-top: 3px;
      overflow-wrap: anywhere;
      word-break: break-word;
      line-height: 1.35;
    }
    .hover-context {
      display: none;
      position: absolute;
      left: 10px;
      right: 10px;
      top: calc(100% - 2px);
      z-index: 8;
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #ffffff;
      box-shadow: 0 18px 48px rgba(21, 33, 51, .18);
      pointer-events: none;
    }
    .hover-context div + div { margin-top: 6px; }
    .hover-context span {
      display: block;
      color: var(--muted);
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
    }
    .hover-context strong {
      display: block;
      margin-top: 2px;
      color: var(--text);
      font-size: 12px;
      line-height: 1.35;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    tbody tr:hover .hover-context,
    tbody tr:focus-within .hover-context {
      display: block;
    }
    .pill {
      display: inline-block;
      border-radius: 999px;
      padding: 2px 8px;
      font-size: 12px;
      font-weight: 650;
      background: var(--panel-2);
      color: var(--muted);
    }
    .pill.unmanaged { color: var(--ok); background: #e9f7ef; }
    .pill.system { color: #7a4b00; background: #fff4d6; }
    .filters {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
      padding: 12px 16px;
      border-bottom: 1px solid var(--line);
      background: #fff;
      position: sticky;
      top: 0;
      z-index: 1;
    }
    .filter-check {
      display: inline-flex;
      gap: 6px;
      align-items: center;
      color: var(--text);
      font-size: 13px;
      user-select: none;
    }
    .filter-search {
      flex: 1;
      min-width: 220px;
    }
    tr.selected-row { background: var(--accent-soft); }
    tr.selected-row:hover { background: #e7f0ff; }

    .side-block:last-child { border-bottom: 0; }
    .hint { color: var(--muted); font-size: 12px; line-height: 1.45; margin-top: 7px; }
    .summary {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
      max-height: 118px;
      overflow: auto;
    }

    .log {
      min-height: 220px;
      max-height: 360px;
      overflow: auto;
      background: #111827;
      color: #d7e0ea;
      border-radius: 6px;
      padding: 10px;
      font: 12px Consolas, monospace;
      white-space: pre-wrap;
    }

    .mini-log {
      min-height: 92px;
      max-height: 180px;
      overflow: auto;
      background: #f8fafc;
      color: var(--text);
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px;
      font: 12px Consolas, monospace;
      white-space: pre-wrap;
      margin-top: 10px;
    }

    .signin-callout {
      display: none;
      margin-top: 10px;
      padding: 10px;
      border: 1px solid #f4cf91;
      border-radius: 6px;
      background: #fff8e8;
      color: var(--warning);
      font-size: 12px;
      line-height: 1.4;
    }

    body.auth-needed .signin-callout { display: block; }

    body.auth-needed #authCreate {
      border-color: #d99028;
      box-shadow: 0 0 0 3px rgba(217, 144, 40, .16);
      animation: sign-in-nudge 1.4s ease-in-out infinite;
    }

    @keyframes sign-in-nudge {
      0%, 100% { transform: translateY(0); }
      50% { transform: translateY(-1px); }
    }

    .button-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }

    .input-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 8px;
    }

    .status {
      margin-top: 10px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }

    .error { color: var(--danger); }
    .ok { color: var(--ok); }

    @media (max-width: 860px) {
      main { grid-template-columns: 1fr; }
      .toolbar { grid-template-columns: 1fr; }
      .solution-list { max-height: none; }
      .header-row { display: block; }
      .header-actions { justify-content: flex-start; margin-top: 10px; }
      .auth-chip { margin-top: 10px; max-width: 100%; }
      .input-row { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body class="auth-checking">
  <div class="app-shell" id="appShell">
    <header>
      <div class="header-row">
        <div>
          <h1>PAC Solution Exporter</h1>
          <div class="subtitle">List Power Platform solutions from PAC, select with checkboxes, and export unmanaged ZIP packages locally.</div>
        </div>
        <div class="header-actions">
          <a class="setup-link" href="/setup" target="_blank" rel="noopener">Setup guide</a>
          <div class="auth-chip" id="authChip">Checking PAC auth...</div>
        </div>
      </div>
    </header>

    <main>
    <section>
      <div class="toolbar">
        <div>
          <label for="environment">Environment URL or ID</label>
          <input id="environment" placeholder="Blank uses active PAC environment">
        </div>
        <div>
          <label for="includeSystem">System solutions</label>
          <button class="secondary" id="includeSystem" type="button" aria-pressed="false">Excluded</button>
        </div>
        <button id="loadSolutions" type="button">Load solutions</button>
      </div>
      <div class="solution-list">
        <div class="solution-progress-bar" aria-hidden="true"></div>
        <div class="solution-progress" id="solutionProgress" role="status" aria-live="polite">Loading solutions from PAC...</div>
        <div class="filters">
          <input class="filter-search" id="searchText" placeholder="Filter by unique or friendly name">
          <label class="filter-check"><input id="showUnmanaged" type="checkbox" checked> Unmanaged</label>
          <label class="filter-check"><input id="showManaged" type="checkbox" checked> Managed</label>
          <label class="filter-check"><input id="showSystem" type="checkbox"> System/reference</label>
        </div>
        <table>
          <thead>
            <tr>
              <th class="select"><input id="selectAllVisible" type="checkbox" title="Select all visible solutions"></th>
              <th style="width:50px">No.</th>
              <th>Solution</th>
              <th style="width:105px">Version</th>
              <th style="width:140px">Last updated</th>
              <th style="width:135px">Updated by</th>
              <th style="width:120px">Type</th>
            </tr>
          </thead>
          <tbody id="solutionsBody">
            <tr><td colspan="7" class="hint">Load solutions to begin.</td></tr>
          </tbody>
        </table>
      </div>
    </section>

    <aside>
      <div class="side-block">
        <label>PAC authentication</label>
        <div class="button-row">
          <button class="secondary" id="authList" type="button">Auth list</button>
          <button class="secondary" id="authCreate" type="button">Sign in</button>
        </div>
        <div class="hint">Sign in uses the environment box above when provided and may open a Microsoft login window.</div>
        <div class="signin-callout" id="signinCallout">No usable PAC auth profile was found. Use Sign in, complete Microsoft login, then click Auth list to confirm.</div>
        <div class="mini-log" id="authOutput">PAC auth status has not been checked.</div>
      </div>

      <div class="side-block">
        <label>Selected solutions</label>
        <div class="hint">Check rows in the table, or use the header checkbox to select every currently visible filtered row.</div>
        <div class="summary" id="selectedSummary"></div>
      </div>

      <div class="side-block">
        <label for="outputDir">Export folder</label>
        <div class="input-row">
          <input id="outputDir" placeholder="Blank creates ./exports/<timestamp>">
          <button class="secondary" id="pickOutputDir" type="button">Browse</button>
        </div>
        <div class="hint">The folder path is resolved on this machine, where the Node app is running.</div>
      </div>

      <div class="side-block">
        <button id="exportSelected" type="button">Export selected</button>
        <button class="secondary" id="clearSelection" type="button">Clear</button>
        <div class="status" id="status">Ready.</div>
      </div>

      <div class="side-block">
        <label>Export log</label>
        <div class="log" id="log"></div>
      </div>
    </aside>
    </main>
  </div>

  <div class="auth-gate" id="authGate" role="status" aria-live="polite">
    <div class="auth-gate-panel">
      <div class="auth-gate-title">Checking PAC authentication</div>
      <div class="auth-gate-copy">The exporter is verifying local Power Platform CLI auth before enabling the workspace.</div>
      <div class="progress-bar" aria-hidden="true"></div>
    </div>
  </div>

  <script>
    let solutions = [];
    let selectedSolutionNames = new Set();
    let includeSystemSolutions = false;
    let activeJobId = '';
    let pollTimer = 0;
    let authState = { authenticated: false, user: '' };

    const $ = id => document.getElementById(id);

    function setStatus(text, kind) {
      $('status').className = 'status ' + (kind || '');
      $('status').textContent = text;
    }

    function setLog(lines) {
      const log = $('log');
      log.textContent = Array.isArray(lines) ? lines.join('\\n') : String(lines || '');
      log.scrollTop = log.scrollHeight;
    }

    function setAuthOutput(text) {
      const output = $('authOutput');
      output.textContent = text || '';
      output.scrollTop = output.scrollHeight;
    }

    function setSolutionsLoading(isLoading) {
      const button = $('loadSolutions');
      const list = document.querySelector('.solution-list');
      list.classList.toggle('loading', isLoading);
      button.classList.toggle('loading', isLoading);
      button.disabled = isLoading;
      button.textContent = isLoading ? 'Loading...' : 'Load solutions';
      button.setAttribute('aria-busy', String(isLoading));
    }

    function debounce(fn, delay) {
      let timer = 0;
      return (...args) => {
        window.clearTimeout(timer);
        timer = window.setTimeout(() => fn(...args), delay);
      };
    }

    function updateAuthState(data) {
      authState = {
        authenticated: Boolean(data && data.authenticated),
        user: data && data.user ? String(data.user) : '',
      };
      const chip = $('authChip');
      chip.classList.toggle('signed-in', authState.authenticated);
      chip.textContent = authState.authenticated
        ? 'Signed in: ' + (authState.user || 'PAC profile found')
        : 'Not signed in';
      document.body.classList.remove('auth-checking');
      document.body.classList.toggle('auth-needed', !authState.authenticated);
      $('authGate').hidden = true;
      $('authCreate').disabled = authState.authenticated;
      if (!authState.authenticated) {
        window.requestAnimationFrame(() => $('authCreate').focus({ preventScroll: false }));
      }
    }

    function visibleSolutions() {
      const search = $('searchText').value.trim().toLowerCase();
      const showUnmanaged = $('showUnmanaged').checked;
      const showManaged = $('showManaged').checked;
      const showSystem = $('showSystem').checked;

      return solutions
        .map((solution, index) => ({ solution, index }))
        .filter(item => {
          const solution = item.solution;
          if (solution.system && !showSystem) return false;
          if (solution.managed && !showManaged) return false;
          if (!solution.managed && !showUnmanaged) return false;
          if (!search) return true;
          return (String(solution.uniqueName || '') + ' ' + String(solution.friendlyName || '')).toLowerCase().includes(search);
        });
    }

    function selectedSolutions() {
      return solutions.filter(solution => selectedSolutionNames.has(solution.uniqueName));
    }

    function updateSelectAllState() {
      const checkbox = $('selectAllVisible');
      const visible = visibleSolutions();
      const selectedVisible = visible.filter(item => selectedSolutionNames.has(item.solution.uniqueName));
      checkbox.checked = visible.length > 0 && selectedVisible.length === visible.length;
      checkbox.indeterminate = selectedVisible.length > 0 && selectedVisible.length < visible.length;
    }

    function renderSelection() {
      const summary = $('selectedSummary');
      const selected = selectedSolutions();
      summary.innerHTML = '';
      selected.forEach(solution => {
        const pill = document.createElement('span');
        pill.className = 'pill unmanaged';
        pill.textContent = solution.uniqueName;
        summary.appendChild(pill);
      });
      if (!selected.length) {
        const pill = document.createElement('span');
        pill.className = 'pill';
        pill.textContent = 'No solutions selected';
        summary.appendChild(pill);
      }
      updateSelectAllState();
    }

    function renderSolutions() {
      const body = $('solutionsBody');
      body.innerHTML = '';
      const visible = visibleSolutions();
      if (!solutions.length) {
        body.innerHTML = '<tr><td colspan="7" class="hint">No solutions loaded.</td></tr>';
        renderSelection();
        return;
      }
      if (!visible.length) {
        body.innerHTML = '<tr><td colspan="7" class="hint">No solutions match the current filters.</td></tr>';
        renderSelection();
        return;
      }
      const fragment = document.createDocumentFragment();
      visible.forEach(item => {
        const solution = item.solution;
        const index = item.index;
        const row = document.createElement('tr');
        if (selectedSolutionNames.has(solution.uniqueName)) row.classList.add('selected-row');
        row.innerHTML =
          '<td class="select"><input type="checkbox" class="row-check" title="Select solution"></td>' +
          '<td class="num">' + (index + 1) + '</td>' +
          '<td class="solution-cell"><div class="name"></div><div class="friendly"></div><div class="hover-context"><div><span>Unique name</span><strong class="ctx-unique"></strong></div><div><span>Friendly name</span><strong class="ctx-friendly"></strong></div></div></td>' +
          '<td></td>' +
          '<td></td>' +
          '<td></td>' +
          '<td class="managed"><span class="pill"></span></td>';
        const checkbox = row.children[0].querySelector('.row-check');
        checkbox.checked = selectedSolutionNames.has(solution.uniqueName);
        const solutionCell = row.children[2];
        const friendlyName = solution.friendlyName || '';
        solutionCell.title = 'Unique: ' + solution.uniqueName + (friendlyName ? '\\nFriendly: ' + friendlyName : '');
        solutionCell.querySelector('.name').textContent = solution.uniqueName;
        solutionCell.querySelector('.friendly').textContent = friendlyName;
        solutionCell.querySelector('.ctx-unique').textContent = solution.uniqueName;
        solutionCell.querySelector('.ctx-friendly').textContent = friendlyName || '-';
        row.children[3].textContent = solution.version || '';
        row.children[4].textContent = solution.modifiedOn || '-';
        row.children[5].textContent = solution.modifiedBy || '-';
        const pill = row.children[6].querySelector('.pill');
        pill.textContent = solution.system ? 'System' : (solution.managed ? 'Managed' : 'Unmanaged');
        if (solution.system) pill.classList.add('system');
        if (!solution.managed && !solution.system) pill.classList.add('unmanaged');

        function setSelected(checked) {
          if (checked) {
            selectedSolutionNames.add(solution.uniqueName);
          } else {
            selectedSolutionNames.delete(solution.uniqueName);
          }
          renderSolutions();
        }

        checkbox.addEventListener('click', event => {
          event.stopPropagation();
          setSelected(event.target.checked);
        });
        row.addEventListener('click', () => {
          setSelected(!selectedSolutionNames.has(solution.uniqueName));
        });
        fragment.appendChild(row);
      });
      body.appendChild(fragment);
      renderSelection();
    }

    async function postJson(url, body) {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await response.json();
      if (!response.ok || data.ok === false) {
        throw new Error(data.error || 'Request failed');
      }
      return data;
    }

    async function loadSolutions() {
      setStatus('Loading solutions from PAC...');
      setLog('');
      setSolutionsLoading(true);
      try {
        const data = await postJson('/api/solutions', {
          environment: $('environment').value,
          includeSystemSolutions,
        });
        solutions = data.solutions || [];
        const available = new Set(solutions.map(solution => solution.uniqueName));
        selectedSolutionNames = new Set([...selectedSolutionNames].filter(name => available.has(name)));
        renderSolutions();
        setStatus('Loaded ' + solutions.length + ' solutions.' + (data.metadataWarning ? ' Metadata warning: ' + data.metadataWarning : ''), data.metadataWarning ? 'error' : 'ok');
      } catch (error) {
        setStatus(error.message, 'error');
        setLog(error.message);
      } finally {
        setSolutionsLoading(false);
      }
    }

    async function authList(options) {
      const silent = options && options.silent;
      if (!silent) setAuthOutput('Checking PAC auth profiles...');
      $('authList').disabled = true;
      try {
        const data = await postJson('/api/auth/list', {});
        updateAuthState(data);
        setAuthOutput(data.output || (data.authenticated ? 'PAC auth profile found.' : 'No auth profiles returned.'));
      } catch (error) {
        updateAuthState({ authenticated: false, user: '' });
        setAuthOutput(error.message);
      } finally {
        $('authList').disabled = false;
      }
    }

    async function authCreate() {
      setAuthOutput('Starting PAC sign-in. Complete the Microsoft login window if one opens...');
      $('authCreate').disabled = true;
      try {
        const data = await postJson('/api/auth/create', {
          environment: $('environment').value,
        });
        updateAuthState(data);
        setAuthOutput(data.output || 'PAC sign-in completed.');
        await authList({ silent: true });
      } catch (error) {
        setAuthOutput(error.message);
      } finally {
        $('authCreate').disabled = authState.authenticated;
      }
    }

    async function pickOutputDir() {
      setStatus('Opening folder picker...');
      $('pickOutputDir').disabled = true;
      try {
        const data = await postJson('/api/folder/select', {
          initialPath: $('outputDir').value,
        });
        if (data.cancelled) {
          setStatus('Folder selection cancelled.');
          return;
        }
        if (data.path) {
          $('outputDir').value = data.path;
          setStatus('Export folder selected.', 'ok');
        }
      } catch (error) {
        setStatus(error.message, 'error');
      } finally {
        $('pickOutputDir').disabled = false;
      }
    }

    async function exportSelected() {
      const selected = selectedSolutions();
      if (!selected.length) {
        setStatus('Select at least one solution checkbox.', 'error');
        return;
      }
      setStatus('Starting export job...');
      setLog('');
      $('exportSelected').disabled = true;
      try {
        const data = await postJson('/api/export', {
          environment: $('environment').value,
          outputDir: $('outputDir').value,
          solutionNames: selected.map(solution => solution.uniqueName),
        });
        activeJobId = data.job.id;
        pollJob();
      } catch (error) {
        setStatus(error.message, 'error');
        $('exportSelected').disabled = false;
      }
    }

    async function pollJob() {
      if (!activeJobId) return;
      try {
        const response = await fetch('/api/jobs/' + encodeURIComponent(activeJobId));
        const data = await response.json();
        const job = data.job;
        setLog(job.logs || []);
        const label = job.current ? 'Current: ' + job.current + '. ' : '';
        setStatus(label + job.completed + '/' + job.total + ' completed, ' + job.failed + ' failed. Output: ' + (job.outputDir || 'pending'));
        if (['completed', 'completed_with_failures', 'failed'].includes(job.status)) {
          setStatus('Export job ' + job.status + '. Output: ' + job.outputDir, job.failed ? 'error' : 'ok');
          $('exportSelected').disabled = false;
          return;
        }
        pollTimer = window.setTimeout(pollJob, 1200);
      } catch (error) {
        setStatus(error.message, 'error');
        $('exportSelected').disabled = false;
      }
    }

    const renderSolutionsDebounced = debounce(renderSolutions, 80);

    $('includeSystem').addEventListener('click', () => {
      includeSystemSolutions = !includeSystemSolutions;
      $('includeSystem').textContent = includeSystemSolutions ? 'Included' : 'Excluded';
      $('includeSystem').setAttribute('aria-pressed', String(includeSystemSolutions));
    });
    $('authList').addEventListener('click', authList);
    $('authCreate').addEventListener('click', authCreate);
    $('loadSolutions').addEventListener('click', loadSolutions);
    $('pickOutputDir').addEventListener('click', pickOutputDir);
    $('exportSelected').addEventListener('click', exportSelected);
    $('clearSelection').addEventListener('click', () => {
      selectedSolutionNames.clear();
      renderSolutions();
    });
    $('selectAllVisible').addEventListener('change', event => {
      visibleSolutions().forEach(item => {
        if (event.target.checked) {
          selectedSolutionNames.add(item.solution.uniqueName);
        } else {
          selectedSolutionNames.delete(item.solution.uniqueName);
        }
      });
      renderSolutions();
    });
    $('searchText').addEventListener('input', renderSolutionsDebounced);
    $('showUnmanaged').addEventListener('change', renderSolutions);
    $('showManaged').addEventListener('change', renderSolutions);
    $('showSystem').addEventListener('change', renderSolutions);
    window.addEventListener('beforeunload', () => window.clearTimeout(pollTimer));
    authList({ silent: true });
    renderSolutions();
  </script>
</body>
</html>`;
}

function setupGuideHtml() {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Setup Guide - PAC Solution Exporter</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f4f6f8;
      --panel: #ffffff;
      --panel-2: #f8fafc;
      --text: #16202c;
      --muted: #66758a;
      --line: #dce3ec;
      --accent: #245bd8;
      --accent-2: #1747ad;
      --ok: #147a46;
      --shadow: 0 16px 40px rgba(21, 33, 51, .08);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Segoe UI, Arial, sans-serif;
      font-size: 14px;
      letter-spacing: 0;
    }
    header {
      background: #fff;
      border-bottom: 1px solid var(--line);
      padding: 16px 24px;
    }
    .header-row {
      max-width: 980px;
      margin: 0 auto;
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: center;
    }
    h1 { margin: 0; font-size: 22px; line-height: 1.2; }
    .subtitle { margin-top: 5px; color: var(--muted); font-size: 13px; }
    a.button {
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 12px;
      background: #fff;
      color: var(--text);
      font-weight: 700;
      text-decoration: none;
      white-space: nowrap;
    }
    a.button:hover { background: var(--panel-2); }
    main {
      max-width: 980px;
      margin: 0 auto;
      padding: 20px;
      display: grid;
      gap: 16px;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 18px;
    }
    h2 { margin: 0 0 10px; font-size: 16px; }
    p { margin: 8px 0; line-height: 1.55; }
    ul, ol { margin: 8px 0 0 20px; padding: 0; }
    li { margin: 6px 0; line-height: 1.45; }
    code {
      background: #eef3f8;
      border: 1px solid var(--line);
      border-radius: 4px;
      padding: 1px 4px;
      font-family: Consolas, monospace;
      font-size: 12px;
    }
    pre {
      overflow: auto;
      background: #111827;
      color: #d7e0ea;
      border-radius: 8px;
      padding: 12px;
      font: 12px Consolas, monospace;
      line-height: 1.5;
    }
    .grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
    }
    .mini {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #fbfcfe;
    }
    .mini strong { display: block; margin-bottom: 4px; }
    .ok { color: var(--ok); font-weight: 700; }
    .muted { color: var(--muted); }
    @media (max-width: 760px) {
      .header-row { display: block; }
      a.button { margin-top: 12px; }
      .grid { grid-template-columns: 1fr; }
      main { padding: 14px; }
    }
  </style>
</head>
<body>
  <header>
    <div class="header-row">
      <div>
        <h1>PAC Solution Exporter setup</h1>
        <div class="subtitle">Short first-time setup for Windows users.</div>
      </div>
      <a class="button" href="/">Back to app</a>
    </div>
  </header>

  <main>
    <section>
      <h2>What you need</h2>
      <div class="grid">
        <div class="mini"><strong>Node.js LTS</strong><span class="muted">Runs this local web app.</span></div>
        <div class="mini"><strong>Power Platform CLI</strong><span class="muted">Provides the <code>pac</code> commands used to list/export solutions.</span></div>
        <div class="mini"><strong>Power Platform access</strong><span class="muted">Your account must be able to read and export solutions from the target environment.</span></div>
        <div class="mini"><strong>This app folder</strong><span class="muted">Keep <code>server.js</code>, <code>package.json</code>, and the <code>scripts</code> folder together.</span></div>
      </div>
    </section>

    <section>
      <h2>Fast setup</h2>
      <p>Open PowerShell in the app folder and run:</p>
      <pre>Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
.\\scripts\\Install-Prerequisites.ps1 -InstallMissing
.\\scripts\\Create-Desktop-Shortcut.ps1
.\\Start-PAC-Solution-Exporter.cmd</pre>
      <p class="muted">The install script checks for Node.js and PAC CLI. If either is missing, it installs the missing component when <code>-InstallMissing</code> is used.</p>
    </section>

    <section>
      <h2>Manual install commands</h2>
      <p>If your company blocks scripts, install the components manually:</p>
      <pre>winget install -e --id OpenJS.NodeJS.LTS
Invoke-WebRequest https://aka.ms/PowerAppsCLI -OutFile "$env:TEMP\\powerapps-cli.msi"
msiexec.exe /i "$env:TEMP\\powerapps-cli.msi" /passive</pre>
      <p>Close and reopen PowerShell, then verify:</p>
      <pre>node -v
npm -v
pac</pre>
    </section>

    <section>
      <h2>First sign in</h2>
      <p>Use either the app's <strong>Sign in</strong> button or run PAC directly:</p>
      <pre>pac auth create --environment "&lt;environment-url-or-id&gt;"
pac auth list</pre>
      <p class="muted">After sign-in, reload the app. The top-right auth badge should show your username.</p>
    </section>

    <section>
      <h2>Daily use</h2>
      <ol>
        <li>Double-click <code>PAC Solution Exporter</code> on your desktop, or run <code>Start-PAC-Solution-Exporter.cmd</code>.</li>
        <li>Enter the environment URL or ID.</li>
        <li>Click <strong>Load solutions</strong>, select solutions, choose an export folder, then click <strong>Export selected</strong>.</li>
      </ol>
      <p class="ok">The app only lists and exports. It does not import, delete, publish, or add Dataverse components.</p>
    </section>

    <section>
      <h2>Official references</h2>
      <ul>
        <li><a href="https://nodejs.org/en/download" target="_blank" rel="noopener">Node.js downloads</a></li>
        <li><a href="https://learn.microsoft.com/power-platform/developer/cli/introduction#install-microsoft-power-platform-cli" target="_blank" rel="noopener">Power Platform CLI install options</a></li>
        <li><a href="https://learn.microsoft.com/power-platform/developer/howto/install-cli-msi" target="_blank" rel="noopener">Power Platform CLI Windows MSI</a></li>
      </ul>
    </section>
  </main>
</body>
</html>`;
}

async function route(req, res) {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

  if (req.method === 'GET' && url.pathname === '/') {
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
    res.end(indexHtml());
    return;
  }

  if (req.method === 'GET' && url.pathname === '/setup') {
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
    res.end(setupGuideHtml());
    return;
  }

  if (req.method === 'GET' && url.pathname === '/api/health') {
    sendJson(res, 200, { ok: true, port: PORT });
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/auth/list') {
    try {
      const result = await listAuthProfiles();
      sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      sendJson(res, 400, { ok: false, error: error.message });
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/auth/create') {
    try {
      const result = await createAuthProfile(await readBody(req));
      sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      sendJson(res, 400, { ok: false, error: error.message });
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/folder/select') {
    try {
      const result = await selectExportFolder(await readBody(req));
      sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      sendJson(res, 400, { ok: false, error: error.message });
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/solutions') {
    try {
      sendJson(res, 200, await listSolutions(await readBody(req)));
    } catch (error) {
      sendJson(res, 400, { ok: false, error: error.message });
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/export') {
    try {
      const result = createExportJob(await readBody(req));
      sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      sendJson(res, 400, { ok: false, error: error.message });
    }
    return;
  }

  const jobMatch = url.pathname.match(/^\/api\/jobs\/([^/]+)$/);
  if (req.method === 'GET' && jobMatch) {
    const job = jobs.get(decodeURIComponent(jobMatch[1]));
    if (!job) {
      sendJson(res, 404, { ok: false, error: 'Job not found.' });
      return;
    }
    sendJson(res, 200, { ok: true, job });
    return;
  }

  sendJson(res, 404, { ok: false, error: 'Not found.' });
}

const server = http.createServer((req, res) => {
  route(req, res).catch(error => sendJson(res, 500, { ok: false, error: error.message }));
});

server.listen(PORT, HOST, () => {
  console.log(`PAC Solution Exporter running at http://${HOST}:${PORT}`);
});
