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
const SHEET_NAME_REPOS = 'repo_ideas';  // 旧シート（互換用）
const SHEET_NAME_CASES = 'cases';       // 新統合シート
const SHEET_NAME_CLIENTS = 'clients';   // クライアントマスタ
const SHEET_NAME_DATA = 'Data';
const SHEET_NAME_LOG = 'ExecutionLog';
const GITHUB_TOKEN = ''; // 任意: ghp_xxxx 形式のトークン（プライベートリポジトリ用）

// cases シートのヘッダー定義
const CASES_HEADERS = [
  'id', 'stage', 'source', 'repo_url', 'name', 'description', 'language', 'stars',
  'github_pushed_at', 'client_id', 'contract_type', 'monthly_amount', 'contract_start',
  'contract_end', 'auto_renew', 'contract_url', 'billing_status', 'invoice_number',
  'invoice_date', 'paid_date', 'memo', 'created_at', 'updated_at'
];

// clients シートのヘッダー定義
const CLIENTS_HEADERS = [
  'client_id', 'name', 'address', 'contact_name', 'contact_email',
  'billing_address', 'billing_method', 'notes', 'created_at', 'updated_at'
];

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

  // === 新統合API ===
  if (action === 'get_cases') {
    return getCases(params.stage);
  }

  if (action === 'get_clients') {
    return getClients();
  }

  // キーバリューストレージ
  if (params.key) {
    return getValue(params.key);
  }

  return jsonResponse({ status: 'error', message: 'Unknown action' });
}

/**
 * POST リクエストハンドラ
 * Content-Type: text/plain でCORSプリフライト回避
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return handleRequest(body);
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'doPost error: ' + err.message });
  }
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

  if (action === 'sync_repos') {
    return syncAllReposManual();
  }

  // === 新統合API ===
  if (action === 'create_case') {
    return createCase(body);
  }

  if (action === 'update_case') {
    return updateCase(body);
  }

  if (action === 'promote_case') {
    return promoteCase(body.id, body.data || {});
  }

  if (action === 'create_client') {
    return createClient(body);
  }

  if (action === 'update_client') {
    return updateClient(body);
  }

  if (action === 'import_repos_to_cases') {
    return importReposToCases();
  }

  // === トリガー管理 ===
  if (action === 'setup_trigger') {
    setupDailyTrigger();
    return jsonResponse({ status: 'ok', message: 'Daily trigger set for 3:00 AM' });
  }

  if (action === 'check_trigger') {
    const triggers = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'refreshAllRepos');
    return jsonResponse({ status: 'ok', exists: triggers.length > 0, count: triggers.length });
  }

  return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// リポジトリ CRUD
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 全リポジトリ取得
 * シート構造: id, repo_name, description, language, url, homepage, stars, size_kb,
 *            days_since_push, score, gem_reason, group, priority, status, memo, added_at, updated_at
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

      // updated_at は GitHub の pushed_at の意味なので、両方の名前で返す
      if (repo.updated_at) {
        repo.pushed_at = repo.updated_at;  // エイリアス追加
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
 * 既存シート構造: id, repo_name, description, language, url, homepage, stars, size_kb,
 *                days_since_push, score, gem_reason, group, priority, status, memo, added_at, updated_at
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
    const existingHeaders = data.length > 0 ? data[0] : [];

    if (data.length > 1) {
      const idCol = existingHeaders.indexOf('id');
      if (idCol >= 0) {
        for (let i = 1; i < data.length; i++) {
          existingIds.add(String(data[i][idCol]));
        }
      }
    }

    // 既存のヘッダーを使用（シートが空の場合のみ新規作成）
    const defaultHeaders = ['id', 'repo_name', 'description', 'language', 'url', 'homepage',
                            'stars', 'size_kb', 'days_since_push', 'score', 'gem_reason',
                            'group', 'priority', 'status', 'memo', 'added_at', 'updated_at'];
    const headers = existingHeaders.length > 0 ? existingHeaders : defaultHeaders;

    if (data.length === 0 || data[0].length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 新規リポジトリを追加
    for (const repo of repos) {
      if (existingIds.has(String(repo.id))) continue;

      const pushedAt = repo.pushed_at ? new Date(repo.pushed_at) : null;
      const daysSince = pushedAt ? Math.floor((now - pushedAt) / (1000 * 60 * 60 * 24)) : '';

      const row = headers.map(h => {
        if (h === 'added_at') return now;
        if (h === 'updated_at') return pushedAt || now;
        if (h === 'days_since_push') return daysSince;
        if (h === 'repo_name') return repo.name || '';
        if (h === 'stars') return repo.stargazers_count || 0;
        if (h === 'size_kb') return repo.size || 0;
        if (h === 'gem_reason') return repo.gemReason || '';
        if (h === 'url') return repo.html_url || repo.url || '';
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
 * 全リポジトリの updated_at と days_since_push を GitHub API から再取得
 * トリガーで毎日実行する
 *
 * シート構造: id, repo_name, description, language, url, homepage, stars, size_kb,
 *            days_since_push, score, gem_reason, group, priority, status, memo, added_at, updated_at
 */
function refreshAllRepos() {
  const startTime = new Date();
  const now = new Date();
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
    const repoNameCol = headers.indexOf('repo_name');
    const urlCol = headers.indexOf('url');
    const updatedAtCol = headers.indexOf('updated_at');
    const daysSinceCol = headers.indexOf('days_since_push');
    const starsCol = headers.indexOf('stars');
    const sizeCol = headers.indexOf('size_kb');

    if (updatedAtCol < 0) {
      logExecution('refreshAllRepos', 'error', 'updated_at column not found');
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
          const daysSince = Math.floor((now - pushedAt) / (1000 * 60 * 60 * 24));
          const rowNum = i + 1;

          // updated_at を GitHub の pushed_at で更新
          sheet.getRange(rowNum, updatedAtCol + 1).setValue(pushedAt);

          // days_since_push を更新
          if (daysSinceCol >= 0) {
            sheet.getRange(rowNum, daysSinceCol + 1).setValue(daysSince);
          }

          // stars も更新（変わっている可能性）
          if (starsCol >= 0 && repoData.stargazers_count !== undefined) {
            sheet.getRange(rowNum, starsCol + 1).setValue(repoData.stargazers_count);
          }

          // size も更新
          if (sizeCol >= 0 && repoData.size !== undefined) {
            sheet.getRange(rowNum, sizeCol + 1).setValue(repoData.size);
          }

          // score を再計算（日数経過で「埋もれ判定」が変わるため）
          const scoreCol = headers.indexOf('score');
          const gemReasonCol = headers.indexOf('gem_reason');
          if (scoreCol >= 0) {
            let sc = 0;
            let isGem = false;
            let reason = '';
            if (repoData.stargazers_count > 0) sc += repoData.stargazers_count * 5;
            if (repoData.forks_count > 0) sc += repoData.forks_count * 3;
            if (daysSince > 30 && daysSince < 365) sc += 20;
            if (daysSince < 30) sc += 10;
            if (repoData.description) sc += 15;
            if (repoData.homepage) sc += 20;
            if (repoData.size > 100) sc += 10;
            if (repoData.size > 1000) sc += 15;
            // 埋もれ宝判定: 実装済み（size>50）かつ半年〜2年放置
            if (!repoData.fork && daysSince > 180 && daysSince < 730 && repoData.size > 50) {
              isGem = true;
              reason = '実装済み長期放置';
            }
            // 収益化シグナル
            const name = (repoData.name + ' ' + (repoData.description || '')).toLowerCase();
            if (['stripe', 'payment', 'subscription', 'saas'].some(k => name.includes(k))) {
              sc += 30;
              isGem = true;
              reason = reason || '収益化シグナル';
            }
            // 横展開可能
            if (['template', 'boilerplate', 'starter', 'generator'].some(k => name.includes(k))) {
              sc += 25;
              isGem = true;
              reason = reason || '横展開可能';
            }
            sheet.getRange(rowNum, scoreCol + 1).setValue(Math.max(0, sc));
            if (gemReasonCol >= 0 && reason) {
              sheet.getRange(rowNum, gemReasonCol + 1).setValue(reason);
            }
          }

          success++;
        } else {
          failed++;
          errors.push(`${repoName}: No pushed_at in response`);
        }

        // レート制限対策: 少し待つ（171件 × 100ms = 約17秒）
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

/**
 * GitHub全リポジトリを取得（ページネーション対応）
 */
function fetchAllGitHubRepos(username) {
  const allRepos = [];
  let page = 1;
  const perPage = 100;

  while (true) {
    const url = `https://api.github.com/users/${username}/repos?per_page=${perPage}&page=${page}&sort=pushed`;
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

    if (code !== 200) {
      throw new Error(`GitHub API error: HTTP ${code}`);
    }

    const repos = JSON.parse(response.getContentText());
    if (!repos || repos.length === 0) break;

    allRepos.push(...repos);
    if (repos.length < perPage) break;

    page++;
    Utilities.sleep(100); // レート制限対策
  }

  return allRepos;
}

/**
 * GitHub全リポジトリをシートに同期（不足分を追加）
 * 既存データは保持、新規のみ追加
 */
function syncAllRepos() {
  const GITHUB_USERNAME = 'hideo-t'; // ユーザー名
  const startTime = new Date();
  const now = new Date();

  try {
    logExecution('syncAllRepos', 'info', 'Starting sync...');

    // 1. シートの既存IDを取得
    const sheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const data = sheet.getDataRange().getValues();
    const existingIds = new Set();

    if (data.length > 1) {
      const headers = data[0];
      const idCol = headers.indexOf('id');
      if (idCol >= 0) {
        for (let i = 1; i < data.length; i++) {
          existingIds.add(String(data[i][idCol]));
        }
      }
    }

    const existingCount = existingIds.size;
    logExecution('syncAllRepos', 'info', `Existing repos in sheet: ${existingCount}`);

    // 2. GitHubから全リポジトリ取得
    const allRepos = fetchAllGitHubRepos(GITHUB_USERNAME);
    logExecution('syncAllRepos', 'info', `GitHub repos fetched: ${allRepos.length}`);

    // 3. 不足分を特定
    const missingRepos = allRepos.filter(r => !existingIds.has(String(r.id)));
    logExecution('syncAllRepos', 'info', `Missing repos to add: ${missingRepos.length}`);

    if (missingRepos.length === 0) {
      logExecution('syncAllRepos', 'success', 'All repos already synced. No new repos to add.');
      return;
    }

    // 4. ヘッダー取得
    const headers = data[0];

    // 5. 不足分を追加
    let added = 0;
    for (const repo of missingRepos) {
      const pushedAt = repo.pushed_at ? new Date(repo.pushed_at) : null;
      const daysSince = pushedAt ? Math.floor((now - pushedAt) / (1000 * 60 * 60 * 24)) : '';

      // スコア計算（簡易版）
      let score = 0;
      if (repo.stargazers_count > 0) score += repo.stargazers_count * 5;
      if (repo.forks_count > 0) score += repo.forks_count * 3;
      if (repo.description) score += 15;
      if (repo.homepage) score += 20;
      if (repo.size > 100) score += 10;
      if (repo.size > 1000) score += 15;

      const row = headers.map(h => {
        if (h === 'id') return repo.id;
        if (h === 'repo_name') return repo.name || '';
        if (h === 'description') return repo.description || '';
        if (h === 'language') return repo.language || '';
        if (h === 'url') return repo.html_url || '';
        if (h === 'homepage') return repo.homepage || '';
        if (h === 'stars') return repo.stargazers_count || 0;
        if (h === 'size_kb') return repo.size || 0;
        if (h === 'days_since_push') return daysSince;
        if (h === 'score') return score;
        if (h === 'gem_reason') return '';
        if (h === 'group') return '未分類';
        if (h === 'priority') return 'mid';
        if (h === 'status') return 'idea';
        if (h === 'memo') return '';
        if (h === 'added_at') return now;
        if (h === 'updated_at') return pushedAt || now;
        return '';
      });

      sheet.appendRow(row);
      added++;
    }

    const duration = Math.round((new Date() - startTime) / 1000);
    const summary = `Added ${added} new repos. Total: ${existingCount + added}. Duration: ${duration}s`;
    logExecution('syncAllRepos', 'success', summary);

  } catch (err) {
    logExecution('syncAllRepos', 'error', err.message);
    throw err;
  }
}

/**
 * 同期を手動実行（Web APIから呼び出し）
 */
function syncAllReposManual() {
  syncAllRepos();
  return jsonResponse({ status: 'ok', message: 'Sync completed. Check execution log.' });
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
// 統合パイプライン: Cases CRUD
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * cases シートを初期化
 */
function initCasesSheet() {
  const sheet = getOrCreateSheet(SHEET_NAME_CASES);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0 || data[0][0] !== 'id') {
    sheet.getRange(1, 1, 1, CASES_HEADERS.length).setValues([CASES_HEADERS]);
  }
  return sheet;
}

/**
 * 全cases取得（stageでフィルタ可能）
 */
function getCases(stage) {
  try {
    const sheet = initCasesSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'ok', cases: [] });
    }

    const headers = data[0];
    const cases = [];
    const now = new Date();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const item = {};
      for (let j = 0; j < headers.length; j++) {
        item[headers[j]] = row[j];
      }

      // stageフィルタ
      if (stage && item.stage !== stage) continue;

      // days_since_push を動的に計算
      if (item.github_pushed_at) {
        const pushedDate = new Date(item.github_pushed_at);
        item.days_since_push = Math.floor((now - pushedDate) / (1000 * 60 * 60 * 24));
      }

      cases.push(item);
    }

    return jsonResponse({ status: 'ok', cases: cases });
  } catch (err) {
    logExecution('getCases', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * case作成
 */
function createCase(params) {
  try {
    const sheet = initCasesSheet();
    const now = new Date();

    // 新規ID生成
    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    if (data.length > 1) {
      const idCol = data[0].indexOf('id');
      for (let i = 1; i < data.length; i++) {
        const id = parseInt(data[i][idCol]) || 0;
        if (id > maxId) maxId = id;
      }
    }
    const newId = maxId + 1;

    const row = CASES_HEADERS.map(h => {
      if (h === 'id') return newId;
      if (h === 'stage') return params.stage || 'idea';
      if (h === 'source') return params.source || 'manual';
      if (h === 'created_at') return now;
      if (h === 'updated_at') return now;
      return params[h] !== undefined ? params[h] : '';
    });

    sheet.appendRow(row);
    logExecution('createCase', 'success', 'Created case #' + newId);
    return jsonResponse({ status: 'ok', id: newId });
  } catch (err) {
    logExecution('createCase', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * case更新
 */
function updateCase(params) {
  const id = params.id;
  if (!id) return jsonResponse({ status: 'error', message: 'No id provided' });

  try {
    const sheet = initCasesSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'error', message: 'Case not found' });
    }

    const headers = data[0];
    const idCol = headers.indexOf('id');
    const updatedAtCol = headers.indexOf('updated_at');

    // 対象行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow < 0) {
      return jsonResponse({ status: 'error', message: 'Case not found' });
    }

    // 更新
    for (const [key, value] of Object.entries(params)) {
      if (key === 'id' || key === 'action') continue;
      const col = headers.indexOf(key);
      if (col >= 0) {
        sheet.getRange(targetRow, col + 1).setValue(value);
      }
    }

    // updated_at を更新
    if (updatedAtCol >= 0) {
      sheet.getRange(targetRow, updatedAtCol + 1).setValue(new Date());
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    logExecution('updateCase', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * stage昇格
 * idea → project → contract → invoice
 */
function promoteCase(id, data) {
  if (!id) return jsonResponse({ status: 'error', message: 'No id provided' });

  try {
    const sheet = initCasesSheet();
    const sheetData = sheet.getDataRange().getValues();
    if (sheetData.length <= 1) {
      return jsonResponse({ status: 'error', message: 'Case not found' });
    }

    const headers = sheetData[0];
    const idCol = headers.indexOf('id');
    const stageCol = headers.indexOf('stage');
    const updatedAtCol = headers.indexOf('updated_at');

    // 対象行を検索
    let targetRow = -1;
    let currentStage = '';
    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][idCol]) === String(id)) {
        targetRow = i + 1;
        currentStage = sheetData[i][stageCol];
        break;
      }
    }

    if (targetRow < 0) {
      return jsonResponse({ status: 'error', message: 'Case not found' });
    }

    // 次のstageを決定
    const stageOrder = ['idea', 'project', 'contract', 'invoice'];
    const currentIdx = stageOrder.indexOf(currentStage);
    if (currentIdx < 0 || currentIdx >= stageOrder.length - 1) {
      return jsonResponse({ status: 'error', message: 'Cannot promote from ' + currentStage });
    }
    const nextStage = stageOrder[currentIdx + 1];

    // contract昇格時はclient_id必須
    if (nextStage === 'contract' && !data.client_id) {
      return jsonResponse({ status: 'error', message: 'client_id is required for contract' });
    }

    // invoice昇格時は請求番号を自動採番
    if (nextStage === 'invoice') {
      const now = new Date();
      const ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
      // 同月の請求件数を取得
      const invoiceNumCol = headers.indexOf('invoice_number');
      let seqNum = 1;
      for (let i = 1; i < sheetData.length; i++) {
        const inv = sheetData[i][invoiceNumCol];
        if (inv && String(inv).startsWith(ym)) {
          const num = parseInt(String(inv).split('-')[1]) || 0;
          if (num >= seqNum) seqNum = num + 1;
        }
      }
      data.invoice_number = ym + '-' + String(seqNum).padStart(3, '0');
      data.invoice_date = now;
      data.billing_status = 'unbilled';
    }

    // contract昇格時のデフォルト値
    if (nextStage === 'contract') {
      if (!data.contract_start) data.contract_start = new Date();
      if (data.auto_renew === undefined) data.auto_renew = false;
    }

    // stage更新
    sheet.getRange(targetRow, stageCol + 1).setValue(nextStage);

    // 追加データを更新
    for (const [key, value] of Object.entries(data)) {
      const col = headers.indexOf(key);
      if (col >= 0) {
        sheet.getRange(targetRow, col + 1).setValue(value);
      }
    }

    // updated_at
    if (updatedAtCol >= 0) {
      sheet.getRange(targetRow, updatedAtCol + 1).setValue(new Date());
    }

    logExecution('promoteCase', 'success', 'Promoted #' + id + ' to ' + nextStage);
    return jsonResponse({ status: 'ok', newStage: nextStage });
  } catch (err) {
    logExecution('promoteCase', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// クライアントマスタ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * clients シートを初期化
 */
function initClientsSheet() {
  const sheet = getOrCreateSheet(SHEET_NAME_CLIENTS);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0 || data[0][0] !== 'client_id') {
    sheet.getRange(1, 1, 1, CLIENTS_HEADERS.length).setValues([CLIENTS_HEADERS]);
  }
  return sheet;
}

/**
 * 全クライアント取得
 */
function getClients() {
  try {
    const sheet = initClientsSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'ok', clients: [] });
    }

    const headers = data[0];
    const clients = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const client = {};
      for (let j = 0; j < headers.length; j++) {
        client[headers[j]] = row[j];
      }
      clients.push(client);
    }

    return jsonResponse({ status: 'ok', clients: clients });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * クライアント作成
 */
function createClient(params) {
  try {
    const sheet = initClientsSheet();
    const now = new Date();

    // 新規ID生成
    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const id = parseInt(data[i][0]) || 0;
        if (id > maxId) maxId = id;
      }
    }
    const newId = maxId + 1;

    const row = CLIENTS_HEADERS.map(h => {
      if (h === 'client_id') return newId;
      if (h === 'created_at') return now;
      if (h === 'updated_at') return now;
      return params[h] !== undefined ? params[h] : '';
    });

    sheet.appendRow(row);
    logExecution('createClient', 'success', 'Created client #' + newId + ': ' + (params.name || ''));
    return jsonResponse({ status: 'ok', client_id: newId });
  } catch (err) {
    logExecution('createClient', 'error', err.message);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * クライアント更新
 */
function updateClient(params) {
  const id = params.client_id;
  if (!id) return jsonResponse({ status: 'error', message: 'No client_id provided' });

  try {
    const sheet = initClientsSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ status: 'error', message: 'Client not found' });
    }

    const headers = data[0];

    // 対象行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow < 0) {
      return jsonResponse({ status: 'error', message: 'Client not found' });
    }

    // 更新
    for (const [key, value] of Object.entries(params)) {
      if (key === 'client_id' || key === 'action') continue;
      const col = headers.indexOf(key);
      if (col >= 0) {
        sheet.getRange(targetRow, col + 1).setValue(value);
      }
    }

    // updated_at
    const updatedAtCol = headers.indexOf('updated_at');
    if (updatedAtCol >= 0) {
      sheet.getRange(targetRow, updatedAtCol + 1).setValue(new Date());
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

/**
 * 既存のrepo_ideasシートからcasesシートにインポート
 */
function importReposToCases() {
  try {
    const repoSheet = getOrCreateSheet(SHEET_NAME_REPOS);
    const caseSheet = initCasesSheet();
    const now = new Date();

    const repoData = repoSheet.getDataRange().getValues();
    if (repoData.length <= 1) {
      return jsonResponse({ status: 'ok', imported: 0, message: 'No repos to import' });
    }

    const repoHeaders = repoData[0];
    const caseData = caseSheet.getDataRange().getValues();

    // 既存のcases IDを取得
    const existingRepoUrls = new Set();
    if (caseData.length > 1) {
      const urlCol = caseData[0].indexOf('repo_url');
      for (let i = 1; i < caseData.length; i++) {
        if (caseData[i][urlCol]) existingRepoUrls.add(caseData[i][urlCol]);
      }
    }

    let imported = 0;
    for (let i = 1; i < repoData.length; i++) {
      const repo = {};
      for (let j = 0; j < repoHeaders.length; j++) {
        repo[repoHeaders[j]] = repoData[i][j];
      }

      // 既にインポート済みならスキップ
      if (existingRepoUrls.has(repo.url)) continue;

      // casesシートに追加
      const row = CASES_HEADERS.map(h => {
        if (h === 'id') return repo.id;
        if (h === 'stage') return 'idea';
        if (h === 'source') return 'github';
        if (h === 'repo_url') return repo.url || '';
        if (h === 'name') return repo.repo_name || '';
        if (h === 'description') return repo.description || '';
        if (h === 'language') return repo.language || '';
        if (h === 'stars') return repo.stars || 0;
        if (h === 'github_pushed_at') return repo.updated_at || '';
        if (h === 'memo') return repo.memo || '';
        if (h === 'created_at') return repo.added_at || now;
        if (h === 'updated_at') return repo.updated_at || now;
        return '';
      });

      caseSheet.appendRow(row);
      imported++;
    }

    logExecution('importReposToCases', 'success', 'Imported ' + imported + ' repos');
    return jsonResponse({ status: 'ok', imported: imported });
  } catch (err) {
    logExecution('importReposToCases', 'error', err.message);
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
