/**
 * AIの町医者 司令塔 - Google Apps Script
 *
 * 機能:
 * 1. リポジトリデータのCRUD操作
 * 2. 日次でGitHub APIから pushed_at を再取得して更新
 * 3. 実行ログの記録・表示
 *
 * 設定手順:
 * 1. このコードをGASエディタに貼り付け
 * 2. GITHUB_TOKEN にGitHubのPersonal Access Tokenを設定（任意）
 * 3. トリガー設定: refreshAllRepos を毎日1回（深夜帯）に設定
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 設定
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const SHEET_NAME_REPOS = 'Repos';
const SHEET_NAME_DATA = 'Data';
const SHEET_NAME_LOG = 'ExecutionLog';
const GITHUB_TOKEN = ''; // 任意: ghp_xxxx 形式のトークン（プライベートリポジトリ用）

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Web App エントリーポイント
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function doGet(e) {
  const params = e.parameter;

  // payload パラメータがある場合（POST代替）
  if (params.payload) {
    try {
      const body = JSON.parse(params.payload);
      return handleRequest(body);
    } catch (err) {
      return jsonResponse({ status: 'error', message: 'Invalid payload: ' + err.message });
    }
  }

  // 通常のGETリクエスト
  const action = params.action;

  if (action === 'ping') {
    return jsonResponse({ status: 'ok', message: 'pong', timestamp: new Date().toISOString() });
  }

  if (action === 'get_repos') {
    return getRepos();
  }

  if (action === 'get_log') {
    return getExecutionLog();
  }

  // キーバリューストレージ
  if (params.key) {
    return getValue(params.key);
  }

  return jsonResponse({ status: 'error', message: 'Unknown action' });
}

function handleRequest(body) {
  const action = body.action;

  // キーバリュー保存
  if (body.key && body.value !== undefined) {
    return setValue(body.key, body.value);
  }

  if (action === 'add_repos') {
    return addRepos(body.repos || []);
  }

  if (action === 'update_repo') {
    return updateRepo(body);
  }

  if (action === 'delete_repo') {
    return deleteRepo(body.id);
  }

  if (action === 'refresh_repos') {
    return refreshAllReposManual();
  }

  return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// リポジトリ CRUD
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 全リポジトリ取得
 */
function getRepos() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'ok', repos: [] });
    }

    const headers = data[0];
    const repos = [];
    const now = new Date();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const repo = {};
      for (let j = 0; j < headers.length; j++) {
        repo[headers[j]] = row[j];
      }

      // days_since_push を動的に計算
      if (repo.pushed_at) {
        const pushedDate = new Date(repo.pushed_at);
        repo.days_since_push = Math.floor((now - pushedDate) / (1000 * 60 * 60 * 24));
      } else if (repo.updated_at) {
        const updatedDate = new Date(repo.updated_at);
        repo.days_since_push = Math.floor((now - updatedDate) / (1000 * 60 * 60 * 24));
      }

      repos.push(repo);
    }

    return jsonResponse({ status: 'ok', repos: repos });
  } catch (err) {
    logExecution('getRepos', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * リポジトリ追加
 */
function addRepos(repos) {
  if (!repos || !repos.length) {
    return jsonResponse({ status: 'error', message: 'No repos provided' });
  }

  try {
    const sheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const now = new Date();
    let added = 0;

    // 既存IDを取得
    const existingIds = new Set();
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      const idCol = data[0].indexOf('id');
      if (idCol >= 0) {
        for (let i = 1; i < data.length; i++) {
          existingIds.add(String(data[i][idCol]));
        }
      }
    }

    // ヘッダー確認・作成
    const headers = ['id', 'repo_name', 'name', 'description', 'url', 'homepage', 'language',
                     'stargazers_count', 'forks_count', 'size', 'score', 'isGem', 'gemReason',
                     'group', 'priority', 'status', 'memo',
                     'pushed_at', 'added_at', 'updated_at', 'days_since_push'];

    if (data.length === 0 || data[0].length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 新規リポジトリを追加
    for (const repo of repos) {
      if (existingIds.has(String(repo.id))) continue;

      const row = headers.map(h => {
        if (h === 'added_at') return now;
        if (h === 'updated_at') return repo.pushed_at ? new Date(repo.pushed_at) : now;
        if (h === 'pushed_at') return repo.pushed_at ? new Date(repo.pushed_at) : '';
        if (h === 'repo_name') return repo.name || '';
        if (h === 'days_since_push') return ''; // 動的計算のため空
        return repo[h] !== undefined ? repo[h] : '';
      });

      sheet.appendRow(row);
      added++;
    }

    logExecution('addRepos', 'success', `Added ${added} repos`);
    return jsonResponse({ status: 'ok', added: added });
  } catch (err) {
    logExecution('addRepos', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * リポジトリ更新
 */
function updateRepo(params) {
  const id = params.id;
  if (!id) {
    return jsonResponse({ status: 'error', message: 'No id provided' });
  }

  try {
    const sheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'error', message: 'Repo not found' });
    }

    const headers = data[0];
    const idCol = headers.indexOf('id');
    if (idCol < 0) {
      return jsonResponse({ status: 'error', message: 'id column not found' });
    }

    // 対象行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        targetRow = i + 1; // シートは1-indexed
        break;
      }
    }

    if (targetRow < 0) {
      return jsonResponse({ status: 'error', message: 'Repo not found' });
    }

    // 更新するフィールド
    const updateFields = ['group', 'priority', 'status', 'memo'];
    for (const field of updateFields) {
      if (params[field] !== undefined) {
        const col = headers.indexOf(field);
        if (col >= 0) {
          sheet.getRange(targetRow, col + 1).setValue(params[field]);
        }
      }
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    logExecution('updateRepo', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * リポジトリ削除
 */
function deleteRepo(id) {
  if (!id) {
    return jsonResponse({ status: 'error', message: 'No id provided' });
  }

  try {
    const sheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'error', message: 'Repo not found' });
    }

    const headers = data[0];
    const idCol = headers.indexOf('id');
    if (idCol < 0) {
      return jsonResponse({ status: 'error', message: 'id column not found' });
    }

    // 対象行を検索して削除
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        sheet.deleteRow(i + 1);
        return jsonResponse({ status: 'ok' });
      }
    }

    return jsonResponse({ status: 'error', message: 'Repo not found' });
  } catch (err) {
    logExecution('deleteRepo', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// GitHub API からの日次更新（トリガー用）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 全リポジトリの pushed_at を GitHub API から再取得
 * トリガーで毎日実行する
 */
function refreshAllRepos() {
  const startTime = new Date();
  let success = 0;
  let failed = 0;
  let skipped = 0;
  const errors = [];

  try {
    const sheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      logExecution('refreshAllRepos', 'success', 'No repos to refresh');
      return;
    }

    const headers = data[0];
    const idCol = headers.indexOf('id');
    const repoNameCol = headers.indexOf('repo_name');
    const urlCol = headers.indexOf('url');
    const pushedAtCol = headers.indexOf('pushed_at');
    const updatedAtCol = headers.indexOf('updated_at');

    if (pushedAtCol < 0) {
      logExecution('refreshAllRepos', 'error', 'pushed_at column not found');
      return;
    }

    // 各リポジトリを処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const repoUrl = row[urlCol] || '';
      const repoName = row[repoNameCol] || '';

      // GitHub URLからowner/repoを抽出
      const match = repoUrl.match(/github\.com\/([^\/]+)\/([^\/]+)/);
      if (!match) {
        skipped++;
        continue;
      }

      const owner = match[1];
      const repo = match[2].replace(/\.git$/, '');

      try {
        // GitHub API呼び出し
        const repoData = fetchGitHubRepo(owner, repo);

        if (repoData && repoData.pushed_at) {
          const pushedAt = new Date(repoData.pushed_at);
          const rowNum = i + 1;

          // pushed_at を更新
          sheet.getRange(rowNum, pushedAtCol + 1).setValue(pushedAt);

          // updated_at も pushed_at と同じに（GitHub基準）
          if (updatedAtCol >= 0) {
            sheet.getRange(rowNum, updatedAtCol + 1).setValue(pushedAt);
          }

          success++;
        } else {
          failed++;
          errors.push(`${repoName}: No pushed_at in response`);
        }

        // レート制限対策: 少し待つ
        Utilities.sleep(100);

      } catch (err) {
        failed++;
        errors.push(`${repoName}: ${err.message}`);
      }
    }

    const duration = Math.round((new Date() - startTime) / 1000);
    const summary = `Success: ${success}, Failed: ${failed}, Skipped: ${skipped}, Duration: ${duration}s`;
    logExecution('refreshAllRepos', failed > 0 ? 'partial' : 'success', summary);

    if (errors.length > 0) {
      logExecution('refreshAllRepos_errors', 'info', errors.slice(0, 10).join('; '));
    }

  } catch (err) {
    logExecution('refreshAllRepos', 'error', err.message);
  }
}

/**
 * 手動実行用（Web APIから呼び出し）
 */
function refreshAllReposManual() {
  refreshAllRepos();
  return jsonResponse({ status: 'ok', message: 'Refresh started. Check execution log.' });
}

/**
 * GitHub API からリポジトリ情報を取得
 */
function fetchGitHubRepo(owner, repo) {
  const url = `https://api.github.com/repos/${owner}/${repo}`;
  const options = {
    method: 'GET',
    headers: {
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'AInoMachiisha-Commander'
    },
    muteHttpExceptions: true
  };

  if (GITHUB_TOKEN) {
    options.headers['Authorization'] = 'Bearer ' + GITHUB_TOKEN;
  }

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200) {
    return JSON.parse(response.getContentText());
  } else if (code === 404) {
    throw new Error('Repo not found');
  } else if (code === 403) {
    throw new Error('Rate limit exceeded');
  } else {
    throw new Error(`HTTP ${code}`);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// キーバリューストレージ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getValue(key) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAME_DATA);
    const data = sheet.getDataRange().getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        return jsonResponse({ status: 'ok', value: data[i][1] });
      }
    }

    return jsonResponse({ status: 'ok', value: null });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function setValue(key, value) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAME_DATA);
    const data = sheet.getDataRange().getValues();

    // 既存キーを検索
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(typeof value === 'object' ? JSON.stringify(value) : value);
        return jsonResponse({ status: 'ok' });
      }
    }

    // 新規追加
    sheet.appendRow([key, typeof value === 'object' ? JSON.stringify(value) : value]);
    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 実行ログ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function logExecution(action, status, message) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAME_LOG);

    // ヘッダーがなければ作成
    const data = sheet.getDataRange().getValues();
    if (data.length === 0 || data[0][0] !== 'timestamp') {
      sheet.getRange(1, 1, 1, 4).setValues([['timestamp', 'action', 'status', 'message']]);
    }

    // ログ追加（最新が上に来るように2行目に挿入）
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 4).setValues([[new Date(), action, status, message]]);

    // 100件を超えたら古いログを削除
    const rowCount = sheet.getLastRow();
    if (rowCount > 101) {
      sheet.deleteRows(102, rowCount - 101);
    }
  } catch (err) {
    console.error('logExecution error:', err);
  }
}

function getExecutionLog() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAME_LOG);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return jsonResponse({ status: 'ok', logs: [] });
    }

    const logs = [];
    for (let i = 1; i < Math.min(data.length, 21); i++) { // 最新20件
      logs.push({
        timestamp: data[i][0],
        action: data[i][1],
        status: data[i][2],
        message: data[i][3]
      });
    }

    return jsonResponse({ status: 'ok', logs: logs });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ユーティリティ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// トリガー設定ヘルパー
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 日次トリガーを設定（手動で1回実行）
 */
function setupDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'refreshAllRepos') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 新しいトリガーを設定（毎日午前3時）
  ScriptApp.newTrigger('refreshAllRepos')
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .create();

  logExecution('setupDailyTrigger', 'success', 'Daily trigger set for 3:00 AM');
  console.log('Daily trigger created successfully');
}

/**
 * トリガー状態を確認
 */
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const info = triggers.map(t => ({
    function: t.getHandlerFunction(),
    type: t.getEventType().toString()
  }));
  console.log('Current triggers:', JSON.stringify(info, null, 2));
  return info;
}
