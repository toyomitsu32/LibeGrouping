// ===== はしご酒 グルーピングシステム - Google Apps Script =====
// 参加者データをGoogle Spreadsheetから読み込み、グルーピングアルゴリズムを実行
// Gemini APIでタグ抽出とカード生成を行い、結果をWebアプリで提供

// ===== CONFIG & MENU =====

// 定数定義
const SHEET_PARTICIPANTS = '参加者';
const SHEET_SETTINGS = '設定';
const SHEET_TEAMS = 'チーム名';
const SHEET_RESULTS = '結果';
const DATA_START_ROW = 14;

// メニューの初期化
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 はしご酒グルーピング')
    .addItem('① グルーピング実行', 'runGrouping')
    .addSeparator()
    .addItem('② 第1部のカード生成', 'runCardGenerationPart1')
    .addItem('③ 第2部のカード生成', 'runCardGenerationPart2')
    .addItem('④ 第3部のカード生成', 'runCardGenerationPart3')
    .addItem('⑤ 例外チームのカード生成', 'runCardGenerationPart4')
    .addSeparator()
    .addItem('🔄 手動調整をWebアプリに反映', 'syncResultsFromSheet')
    .addSeparator()
    .addItem('🌐 WebアプリURLを表示', 'showWebAppUrl')
    .addItem('⚙️ 設定シートを開く', 'openSettingsSheet')
    .addToUi();
}

function openSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

// ===== DATA READING =====

/**
 * 設定シートから設定値を読み込む
 * @returns {Object} 設定オブジェクト
 */
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    throw new Error('設定シート（設定）が見つかりません');
  }

  const data = sheet.getRange('A1:B12').getValues();
  const settings = {};

  const mapping = {
    1: 'geminiApiKey',
    2: 'part1Theme',
    3: 'part2Theme',
    4: 'part3Theme',
    5: 'exceptionCategoryName',
    6: 'maxGroupSize',
    7: 'minGroupSize',
    8: 'eventName',
    9: 'part1Time',
    10: 'part2Time',
    11: 'part3Time',
    12: 'part4Time'
  };

  for (const [row, key] of Object.entries(mapping)) {
    const rowIdx = parseInt(row) - 1;
    if (data[rowIdx] && data[rowIdx][1] != null && data[rowIdx][1] !== '') {
      const val = data[rowIdx][1];

      // 時間設定（part1Time, part2Time, part3Time）のフォーマット整形（時間のみ）
      if (val instanceof Date && (key === 'part1Time' || key === 'part2Time' || key === 'part3Time' || key === 'part4Time')) {
        const h = val.getHours().toString().padStart(2, '0');
        const m = val.getMinutes().toString().padStart(2, '0');
        settings[key] = `${h}:${m}`;
      } else {
        settings[key] = String(val).trim();
      }
    }
  }

  // デフォルト値の設定
  settings.maxGroupSize = parseInt(settings.maxGroupSize) || 4;
  settings.minGroupSize = parseInt(settings.minGroupSize) || 3;
  settings.exceptionCategoryName = settings.exceptionCategoryName || '子連れ';

  Logger.log('Settings loaded: ' + JSON.stringify(settings));
  return settings;
}

/**
 * 参加者データを読み込む
 * @returns {Array<Object>} 参加者オブジェクトの配列
 */
function getParticipants() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PARTICIPANTS);
  if (!sheet) {
    throw new Error('参加者シート（参加者）が見つかりません');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    return [];
  }

  // 設定から例外カテゴリー名を取得
  const settings = getSettings();
  const targetLabel = (settings.exceptionCategoryName || '子連れ').toLowerCase();

  // 例外チーム列を動的に探すためのヘッダー検索（行1〜行3を検索）
  const headersRow1 = sheet.getRange(1, 1, 1, 30).getValues()[0];
  const headersRow2 = sheet.getRange(2, 1, 1, 30).getValues()[0];
  let oViceIdx = -1;
  // 列6（G列）以降を検索
  for (let c = 6; c < 30; c++) {
    const h1 = String(headersRow1[c]).toLowerCase();
    const h2 = String(headersRow2[c]).toLowerCase();
    if (h1.includes('ovice') || h1.includes(targetLabel) || h1.includes('参加') ||
      h2.includes('ovice') || h2.includes(targetLabel) || h2.includes('参加')) {
      oViceIdx = c;
      break;
    }
  }
  // 見つけられなかった場合、かつデフォルトが「子連れ」の場合は「子連れ」という単語でもう一度試す
  if (oViceIdx === -1 && targetLabel !== '子連れ') {
    for (let c = 6; c < 30; c++) {
      if (String(headersRow1[c]).includes('子連れ') || String(headersRow2[c]).includes('子連れ')) {
        oViceIdx = c;
        break;
      }
    }
  }

  // それでも見つからなかった場合、H列（index 7）をフォールバックとして使用
  if (oViceIdx === -1) {
    oViceIdx = 7; // H列
    Logger.log('列の自動検出ができなかったため、H列（8列目）をフォールバックとして使用します');
  } else {
    Logger.log('例外カテゴリー列の検索結果: 列' + (oViceIdx + 1) + '（' + String.fromCharCode(65 + oViceIdx) + '列）');
  }

  const maxCol = 40;
  const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, maxCol);
  const data = range.getValues();
  const allHeaders = [
    sheet.getRange(1, 1, 1, maxCol).getValues()[0],
    sheet.getRange(2, 1, 1, maxCol).getValues()[0],
    sheet.getRange(3, 1, 1, maxCol).getValues()[0]
  ];

  const participants = [];

  // 列インデックスの特定
  // デフォルト: B(1)=実名, L(11)=画像URL, M(12)=プロフURL, N(13)=プロフ本文
  let accountIdx = 1;
  let iconUrlIdx = 11;
  let profileUrlIdx = 12;
  let profileIdx = 13;
  let categoryIdx = -1;

  Logger.log('--- Column Detection Start ---');
  for (let c = 0; c < maxCol; c++) {
    const h1 = String(allHeaders[0][c] || '').trim();
    const h2 = String(allHeaders[1][c] || '').trim();
    const h3 = String(allHeaders[2][c] || '').trim();
    const combined = (h1 + " " + h2 + " " + h3).toLowerCase();

    if (combined.trim() === "") continue;

    // どの列が何であるか判定
    // ニックネーム列（系列やカテゴリー、K列という言葉が入っていれば除外）
    if ((combined.includes('ニックネーム') || combined.includes('表示名') || combined.includes('アカウント')) &&
      !combined.includes('系列') && !combined.includes('カテゴリー') && !combined.includes('k列')) {
      accountIdx = c;
    }
    // カテゴリー列
    if (combined.includes('系列') || combined.includes('カテゴリー')) {
      categoryIdx = c;
    }
    // アイコン画像URL列
    if (combined.includes('画像url') || combined.includes('アイコンurl') || (combined.includes('画像') && combined.includes('url')) || combined.includes('アイコン')) {
      // ただしプロフィールURLと書いてある場合は避ける
      if (!combined.includes('プロフィールurl') && !combined.includes('プロフurl')) {
        iconUrlIdx = c;
      }
    }
    // プロフィールURL列
    if ((combined.includes('プロフィールurl') || combined.includes('プロフurl')) && !combined.includes('画像')) {
      profileUrlIdx = c;
    }
    // 自己紹介/プロフィール本文
    if (combined.includes('自己紹介') || combined.includes('本文') || combined.includes('プロフィール') && !combined.includes('url')) {
      profileIdx = c;
    }

    Logger.log(`Column ${c}: "${combined.substring(0, 30)}..."`);
  }

  Logger.log(`[Resulting Mapping] AccountIdx:${accountIdx}, IconIdx:${iconUrlIdx}, ProfileURLIdx:${profileUrlIdx}, ProfileIdx:${profileIdx}, CategoryIdx:${categoryIdx}`);
  Logger.log('--- Column Detection End ---');

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const no = row[0];
    const name = String(row[1]).trim();
    if (!name) continue;

    const gender = String(row[2]).trim();
    const part1Cell = String(row[3]).trim();
    const part2Cell = String(row[4]).trim();
    const part3Cell = String(row[5]).trim();
    const part4Cell = oViceIdx !== -1 ? String(row[oViceIdx]).trim() : '';

    const isParticipating = (val) => {
      if (val === true) return true;
      const v = String(val).trim().toLowerCase();
      return v === 'true' || v.includes('○') || v.includes('〇') || v.includes('⚪') || v === '1' || v === 'yes' || v === '参加';
    };

    const part1 = isParticipating(part1Cell);
    const part2 = isParticipating(part2Cell);
    const part3 = isParticipating(part3Cell);
    const part4 = isParticipating(part4Cell);

    // 指定された列からデータを抽出
    const account = String(row[accountIdx] || '').trim();
    const profile = String(row[profileIdx] || '').trim();
    const rawIconUrl = String(row[iconUrlIdx] || '').trim();
    const rawProfileUrl = String(row[profileUrlIdx] || '').trim();

    // アイコンURLのバリデーション (httpで始まるもの)
    const iconUrl = rawIconUrl.toLowerCase().indexOf('http') === 0 ? rawIconUrl : '';
    const profileUrl = rawProfileUrl.toLowerCase().indexOf('http') === 0 ? rawProfileUrl : '';

    if (part1 || part2 || part3 || part4) {
      participants.push({
        no: no, name: name, gender: gender,
        part1: part1, part2: part2, part3: part3, part4: part4,
        account: account || name,
        profile: profile,
        iconUrl: iconUrl,
        profileUrl: profileUrl
      });
    }
  }

  Logger.log('Loaded ' + participants.length + ' participants');
  return participants;
}

/**
 * チーム名を読み込む
 * @returns {Object} {第1部: [...], 第2部: [...], 第3部: [...]}の形式
 */
function getTeamNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEAMS);
  if (!sheet) {
    return { '第1部': [], '第2部': [], '第3部': [] };
  }

  const data = sheet.getDataRange().getValues();
  const teams = { '第1部': [], '第2部': [], '第3部': [] };

  for (let i = 1; i < data.length; i++) { // ヘッダーをスキップ
    const part = String(data[i][0]).trim();
    const teamName = String(data[i][1]).trim();
    if (part && teamName && teams.hasOwnProperty(part)) {
      teams[part].push(teamName);
    }
  }

  Logger.log('Loaded team names: ' + JSON.stringify(teams));
  return teams;
}

// ===== GROUPING ALGORITHM =====

/**
 * メンバーをグループに分配する
 * @param {Array<string>} members - メンバー名の配列
 * @param {number} minSize - グループの最小人数
 * @param {number} maxSize - グループの最大人数
 * @returns {Array<Array<string>>} グループ化されたメンバーの配列
 */
function distributeIntoGroups(members, minSize, maxSize) {
  // メンバーをシャッフル
  const remaining = [...members];
  shuffleArray(remaining);

  // 均等なサイズになるようメンバーを分配
  // 必要な総グループ数を算出（最大人数で割った時の切り上げ値）
  const totalMembers = members.length;
  // グループ数のベースを計算（上限を超えないように最小限のグループを作る）
  let numGroups = Math.max(1, Math.ceil(totalMembers / maxSize));

  // しかし、もしそのグループ数だと「最低人数(minSize)」を満たせない(例: 10人で3グループだと3,3,4でOK。5人で2グループだと3,2でNG)場合は、
  // 許容できるならグループ数を減らす（ただしその分maxSizeを超える可能性は生じる）
  if (numGroups > 1 && numGroups * minSize > totalMembers) {
    // 例：5人の場合、num=2だと 2*3=6 > 5 となるので、num=1 にする
    numGroups = Math.floor(totalMembers / minSize);
    if (numGroups === 0) numGroups = 1;
  }

  // グループの枠を準備
  const groups = Array.from({ length: numGroups }, () => []);

  // 残りのメンバーを、人数の少ないグループから順に1人ずつ配分していく（ラウンドロビン式）
  for (const member of remaining) {
    const minLenGroup = groups.reduce((a, b) => a.length < b.length ? a : b);
    minLenGroup.push(member);
  }

  Logger.log('Distributed ' + members.length + ' members into ' + groups.length + ' groups');
  return groups;
}

/**
 * グルーピングを実行
 */
function runGrouping() {
  showToast('グルーピングを開始しています...', 'グルーピング');

  try {
    const settings = getSettings();
    const participants = getParticipants();
    const teamNames = getTeamNames();

    const cleanTheme = (t) => (t || '').replace(/\s?チーム$/, '');
    const exceptionLabel = settings.exceptionCategoryName || '子連れ';
    const parts = [
      { key: 'part1', label: '第1部', theme: cleanTheme(settings.part1Theme), singleGroup: false },
      { key: 'part2', label: '第2部', theme: cleanTheme(settings.part2Theme), singleGroup: false },
      { key: 'part3', label: '第3部', theme: cleanTheme(settings.part3Theme), singleGroup: false },
      { key: 'part4', label: exceptionLabel, theme: exceptionLabel, singleGroup: true }
    ];

    const allParticipantNames = participants.map(p => p.name);
    const maxSize = settings.maxGroupSize || 10;

    const groupingResult = {
      timestamp: new Date().toISOString(),
      part1: [],
      part2: [],
      part3: [],
      part4: []
    };

    // 各部ごとにグルーピングを実行
    for (const part of parts) {
      const partMembers = participants
        .filter(p => p[part.key])
        .map(p => p.name);

      if (partMembers.length === 0) continue;

      showToast(part.label + ' のグルーピング中...', 'グルーピング');

      if (part.singleGroup) {
        // oViceなど、単独の1チームにする場合
        groupingResult[part.key] = [{
          team_name: '子連れチーム',
          members: partMembers,
          summary: '',
          cards: []
        }];
        continue;
      }

      const groups = distributeIntoGroups(
        partMembers,
        settings.minGroupSize,
        settings.maxGroupSize
      );

      // チーム名を割り当て
      const availableTeamNames = [...teamNames[part.label]];

      const groupData = groups.map((members, idx) => ({
        team_name: availableTeamNames[idx] || (part.label + ' チーム ' + (idx + 1)),
        members: members,
        summary: '',
        cards: []
      }));

      groupingResult[part.key] = groupData;
    }

    // 結果をシステムシートに一時保存
    setSystemData('groupingResult', groupingResult);

    // アイコンデータを保存
    const iconsData = {};
    participants.forEach(p => {
      if (p.iconUrl) {
        iconsData[p.name] = p.iconUrl;
      }
    });
    setSystemData('iconsData', iconsData);

    // プロフィールURLデータを保存
    const profileUrlsData = {};
    participants.forEach(p => {
      if (p.profileUrl) {
        profileUrlsData[p.name] = p.profileUrl;
      }
    });
    setSystemData('profileUrlsData', profileUrlsData);

    // 古いカード生成結果をクリアし、新しいグルーピング結果で上書き
    setSystemData('cardResult', null);

    // 結果シートおよび WebApp 用データに保存
    saveAllResults();

    showToast('グルーピング完了！', '成功');
    Logger.log('Grouping result: ' + JSON.stringify(groupingResult));

  } catch (e) {
    Logger.log('Error in runGrouping: ' + e.toString());
    showToast('エラーが発生しました: ' + e.toString(), 'エラー');
  }
}

// ===== GEMINI API =====

/**
 * Gemini APIを呼び出す
 * @param {string} prompt - プロンプト
 * @param {string} apiKey - Gemini APIキー
 * @returns {string} APIレスポンスのテキスト
 */
function callGemini(prompt, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(apiKey);

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 30000
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log('Gemini API Error: ' + responseCode + ' - ' + responseText);
      throw new Error('Gemini API returned status ' + responseCode);
    }

    const result = JSON.parse(responseText);
    if (!result.candidates || result.candidates.length === 0) {
      throw new Error('No candidates in Gemini response');
    }

    const content = result.candidates[0].content.parts[0].text;
    return content;

  } catch (e) {
    Logger.log('Error calling Gemini API: ' + e.toString());
    throw e;
  }
}

/**
 * JSON文字列をパース（マークダウンコードフェンスを削除）
 * @param {string} text - パースするテキスト
 * @returns {Object} パースされたJSON
 */
function parseJsonSafely(text) {
  // マークダウンコードフェンスを削除
  let cleaned = text.replace(/^```json\n?/, '').replace(/\n?```$/, '');
  cleaned = cleaned.replace(/^```\n?/, '').replace(/\n?```$/, '');
  return JSON.parse(cleaned.trim());
}

// ===== TAG EXTRACTION =====
// タグ抽出処理はカード生成プロンプトに統合されたため廃止しました。

// ===== CARD GENERATION =====

/**
 * 各部ごとの呼び出し関数
 */
function runCardGenerationPart1() { runCardGeneration('part1', '第1部'); }
function runCardGenerationPart2() { runCardGeneration('part2', '第2部'); }
function runCardGenerationPart3() { runCardGeneration('part3', '第3部'); }
function runCardGenerationPart4() { runCardGeneration('part4', '子連れチーム'); }

/**
 * カード生成を実行（指定された部のみ処理するよう改修）
 * @param {string} targetPart - 処理対象の部（'part1', 'part2', ...）
 * @param {string} partLabel - Toast用ラベル
 */
function runCardGeneration(targetPart, partLabel) {
  showToast(partLabel + 'のカード生成を開始しています...', 'カード生成');

  try {
    const settings = getSettings();
    const props = PropertiesService.getScriptProperties();

    if (!settings.geminiApiKey) {
      throw new Error('Gemini APIキーが設定されていません');
    }

    // スプレッドシート側の手動調整を反映させるために同期を実行
    syncResultsFromSheet();

    const groupingResultStr = props.getProperty('groupingResult');
    if (!groupingResultStr) {
      throw new Error('グルーピング結果が見つかりません。先にグルーピングを実行してください');
    }

    const groupingResult = JSON.parse(groupingResultStr);

    // プロフィール情報を取得してマップ化する
    const allParticipants = getParticipants();
    const profileMap = {};
    for (const p of allParticipants) {
      profileMap[p.name] = p.profile || '情報なし';
    }

    // 既存の結果があれば読み込み、なければ現在のグルーピング結果をベースに初期化
    const existingCardResult = getSystemData('cardResult');
    const cardResult = existingCardResult ? existingCardResult : {
      timestamp: new Date().toISOString(),
      part1: groupingResult.part1 || [],
      part2: groupingResult.part2 || [],
      part3: groupingResult.part3 || [],
      part4: groupingResult.part4 || []
    };

    // タイムスタンプは最新に更新
    cardResult.timestamp = new Date().toISOString();

    let allGroups = groupingResult[targetPart] ? [...groupingResult[targetPart]] : [];

    if (allGroups.length === 0) {
      showToast(partLabel + 'のグルーピングデータがありません。', 'カード生成スキップ');
      return;
    }

    let totalGroups = allGroups.length;
    let processedGroups = 0;
    const batchSize = 4; // 1回で最大4チーム分を同時に回答させる

    for (let i = 0; i < allGroups.length; i += batchSize) {
      const batchedGroups = allGroups.slice(i, i + batchSize);

      showToast('カード生成: ' + Math.min(i + batchSize, totalGroups) + ' / ' + totalGroups, 'カード生成');

      // バッチに含まれるチームそれぞれの情報を組み立てる（プロフィール直書き）
      const teamsText = batchedGroups.map(group => {
        return `チーム名: ${group.team_name}
メンバーの自己紹介文:
${group.members.map(name => `■ ${name}さんの自己紹介:\n${profileMap[name]}`).join('\n\n')}`;
      }).join('\n\n======\n\n');

      const prompt = `以下の複数チームのメンバープロフィールを分析し、**各チームごと**に「楽しい共通点カード」を【概ね6枚】作成してください。

${teamsText}

以下のJSON形式で返してください（Markdownブロックはつけず、生JSONのみ出力）。「チーム名」をキーにしてそれぞれのカード結果を含めてください。
{
  "チーム名A": {
    "summary": "チームの盛り上がりを予感させる、総評（60文字〜100文字）",
    "cards": [
      {
        "category": "EXPERIENCE|HOBBY|BUSINESS|VALUES|OTHER",
        "title": "キャッチーで楽しい共通点のタイトル",
        "description": "クスッと笑えたり「おっ！」と思える楽しい解説文。どのメンバー同士が共通しているのか会話のネタになるように（名前入り）。",
        "members": ["該当メンバー名"]
      }
    ]
  },
  "チーム名B": {
    "summary": "...",
    "cards": [...]
  }
}

ルール:
- 出力キーはリクエストで渡した「チーム名」と一言一句一致させること
- 各カードのcategoryはEXPERIENCE, HOBBY, BUSINESS, VALUES, OTHERのいずれか
- membersには該当メンバーのフルネームをそのまま入れること
- 1チームあたり【6個】程度の共通点を作成すること
- **重要：年齢（同年代など）や性別、身体的特徴に関する共通点は除外すること**。趣味、経験、価値観、ビジネスなどの実用的な共通点にフォーカスのこと`;

      try {
        const response = callGemini(prompt, settings.geminiApiKey);
        const batchedData = parseJsonSafely(response);

        // バッチ処理された各グループに結果を書き戻す
        for (const group of batchedGroups) {
          const teamData = batchedData[group.team_name];
          if (teamData) {
            group.summary = teamData.summary || '';
            group.cards = teamData.cards || [];
            Logger.log('Cards generated for group: ' + group.team_name);
          } else {
            Logger.log('Could not find generation data for team: ' + group.team_name);
            group.summary = group.team_name;
            group.cards = [];
          }
        }
      } catch (e) {
        Logger.log('Error generating batched cards: ' + e.toString());
        // エラー時はフェイルセーフ
        for (const group of batchedGroups) {
          group.summary = group.team_name;
          group.cards = [];
        }
      }

      processedGroups += batchedGroups.length;

      // 無料枠対策: 4チーム一括処理ごとにAPI負担を軽減するウェイト
      if (processedGroups < totalGroups) {
        Utilities.sleep(12000);
      }
    }

    // 元の構造に生成したデータを書き戻す
    cardResult[targetPart] = groupingResult[targetPart] || [];

    // 生成結果をシステムシートに保存
    setSystemData('cardResult', cardResult);

    // 結果シートおよび WebApp 用データに最新状態を書き込む
    saveAllResults();

    showToast('カード生成完了！', '成功');
    Logger.log('Card generation completed');

  } catch (e) {
    Logger.log('Error in runCardGeneration: ' + e.toString());
    showToast('エラーが発生しました: ' + e.toString(), 'エラー');
  }
}

// 全工程を実行する関数は6分制限に引っかかるため廃止し、部ごとの個別実行に切り替えました。

/**
 * 全結果を結果シートに保存
 */
function saveAllResults() {
  try {
    const settings = getSettings();
    const props = PropertiesService.getScriptProperties();

    // 保存前にスプレッドシート側の手動調整内容を読み取って同期する
    syncResultsFromSheet();

    const groupingResult = getSystemData('cardResult') || getSystemData('groupingResult') || {};
    const iconsData = getSystemData('iconsData') || {};
    const profileUrlsData = getSystemData('profileUrlsData') || {};
    const accountsData = getAccountsMap();

    // WebAppで使用するフォーマットに整形
    const webAppData = {
      eventName: settings.eventName || 'はしご酒',
      parts: {
        part1: groupingResult.part1 || [],
        part2: groupingResult.part2 || [],
        part3: groupingResult.part3 || [],
        part4: groupingResult.part4 || []
      },
      partInfo: {
        part1: {
          label: '第1部',
          time: settings.part1Time || '16:50',
          theme: (settings.part1Theme || '').replace(/\s?チーム$/, '')
        },
        part2: {
          label: '第2部',
          time: settings.part2Time || '18:30',
          theme: (settings.part2Theme || '').replace(/\s?チーム$/, '')
        },
        part3: {
          label: '第3部',
          time: settings.part3Time || '20:00',
          theme: (settings.part3Theme || '').replace(/\s?チーム$/, '')
        },
        part4: {
          label: settings.exceptionCategoryName || '子連れ',
          time: settings.part4Time || '16:50',
          theme: settings.exceptionCategoryName || '子連れ'
        }
      },
      icons: iconsData,
      profileUrls: profileUrlsData,
      accounts: accountsData,
      timestamp: new Date().toISOString()
    };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) {
      throw new Error('結果シート（結果）が見つかりません');
    }

    // シート全体を一度クリアする
    sheet.clear();

    // JSONデータ（Webアプリ用）はA1, A2に可視化して配置（ご要望対応）
    sheet.getRange('A1').setValue('【システム用データ】以下の行のデータはWebアプリが読み込むものです').setBackground('#fef08a').setFontWeight('bold');

    const jsonStr = JSON.stringify(webAppData);
    // システムシートに「完全なデータ」を保存
    setSystemData('webAppFinalData', jsonStr);

    if (jsonStr.length < 50000) {
      sheet.getRange('A2').setValue(jsonStr).setFontColor('#6b7280');
      sheet.getRange('AA1').setValue(jsonStr);
    } else {
      // 50,000文字を超える場合はセルを空にする（警告も表示しない）
      sheet.getRange('A2').setValue("");
      sheet.getRange('AA1').setValue("");
    }

    // 結果一覧を表形式で書き出す（JSONの下、4行目から開始）
    let currentRow = 4;

    // タイトル
    sheet.getRange(currentRow, 1).setValue(`🎊 ${webAppData.eventName} グルーピング結果 🎊`).setFontWeight('bold').setFontSize(14);
    currentRow += 2;

    const partsKeys = ['part1', 'part2', 'part3', 'part4'];

    for (const key of partsKeys) {
      const partGroups = webAppData.parts[key];
      if (!partGroups || partGroups.length === 0) continue;

      const partInfo = webAppData.partInfo[key];

      // 各部のヘッダー
      sheet.getRange(currentRow, 1).setValue(`【${partInfo.label}】 ${partInfo.time} 〜 （テーマ：${partInfo.theme || 'なし'}）`).setFontWeight('bold').setBackground('#f3f4f6');
      currentRow++;

      // テーブルヘッダー（チーム名 + 人数 + メンバー1〜10）
      const maxMembers = 10;
      const headerRow = ['チーム名', '人数'];
      for (let m = 1; m <= maxMembers; m++) {
        headerRow.push('メンバー' + m);
      }
      const totalCols = headerRow.length;
      sheet.getRange(currentRow, 1, 1, totalCols).setValues([headerRow]).setFontWeight('bold').setBackground('#e5e7eb');
      currentRow++;

      // グループごとのデータ行
      const outputData = [];
      for (const group of partGroups) {
        const row = [
          group.team_name,
          `${group.members.length}名`
        ];
        // メンバーを1人1セルに展開
        for (let m = 0; m < maxMembers; m++) {
          row.push(group.members[m] || '');
        }
        outputData.push(row);
      }

      // まとめて書き込み
      if (outputData.length > 0) {
        sheet.getRange(currentRow, 1, outputData.length, totalCols).setValues(outputData);
        // 罫線を設定
        sheet.getRange(currentRow - 1, 1, outputData.length + 1, totalCols).setBorder(true, true, true, true, true, true);
        currentRow += outputData.length;
      }

      currentRow += 2; // 部と部の間に空白行
    }

    // 列幅の自動調整
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 60);
    for (let c = 3; c <= 12; c++) {
      sheet.setColumnWidth(c, 200);
    }

  } catch (e) {
    Logger.log('Error in saveAllResults: ' + e.toString());
    throw e;
  }
}

// ===== WEB APP =====

/**
 * スプレッドシートの「結果」シートから現在のグルーピング状況を読み取り、
 * スクリプトプロパティ（JSONデータ）を更新する。
 * これにより、手動でのメンバー入れ替えをWebアプリに反映させる。
 */
function syncResultsFromSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const props = PropertiesService.getScriptProperties();
    const settings = getSettings();

    // 現在のカード生成結果があれば取得（AI生成済みのサマリーを保持するため）
    const groupingResult = getSystemData('cardResult') || getSystemData('groupingResult') || {};

    // 各部のデータをクリアして再構築
    const exceptionLabel = settings.exceptionCategoryName || '子連れ';
    const parts = {
      '【第1部】': 'part1',
      '【第2部】': 'part2',
      '【第3部】': 'part3',
      [`【${exceptionLabel}】`]: 'part4'
    };

    let currentPartKey = null;
    const newPartsData = { part1: [], part2: [], part3: [], part4: [] };

    // シートを走査してチーム名とメンバーを抽出
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0]).trim();

      // 各部のセクション開始を検知
      for (const [label, key] of Object.entries(parts)) {
        if (firstCell.startsWith(label)) {
          currentPartKey = key;
          i++; // ヘッダー行（チーム名、人数...）をスキップ
          break;
        }
      }

      if (currentPartKey && firstCell && !Object.keys(parts).some(l => firstCell.startsWith(l)) && firstCell !== 'チーム名' && firstCell !== '総合判定' && !firstCell.startsWith('🎊')) {
        const teamName = firstCell;
        const members = [];
        // C列以降(index 2〜)にメンバー名が入っている
        for (let c = 2; c < row.length; c++) {
          const mName = String(row[c]).trim();
          if (mName) members.push(mName);
        }

        if (members.length > 0) {
          // 既存のAI生成結果（summary/cards）があれば引き継ぐ
          const existingTeam = (groupingResult[currentPartKey] || []).find(t => t.team_name === teamName);

          newPartsData[currentPartKey].push({
            team_name: teamName,
            members: members,
            summary: existingTeam ? existingTeam.summary : '',
            cards: existingTeam ? existingTeam.cards : []
          });
        }
      }
    }

    // 両方のプロパティを更新
    const updatedResult = {
      timestamp: new Date().toISOString(),
      ...newPartsData
    };

    // メンバー名とチーム名のみの情報を保存（軽量版）
    const lightweightResult = {
      timestamp: new Date().toISOString(),
      ...newPartsData
    };
    setSystemData('groupingResult', lightweightResult);

    // すでにカード生成済みのデータがある場合は、その構造も最新のメンバーに更新
    if (getSystemData('cardResult')) {
      // AIサマリーを保持しつつメンバーだけ更新する処理（updatedResultにはAIサマリーが含まれるためそれを使用）
      setSystemData('cardResult', updatedResult);
    }

    // 重要：スクリプトプロパティだけでなく、シート側のJSON（A2セル）も更新する
    saveAllResultsInternal(updatedResult);

    Logger.log('Synced results from sheet and updated JSON successfully');
    showToast('手動調整内容をWebアプリに反映しました。', '同期完了');
  } catch (e) {
    Logger.log('Error in syncResultsFromSheet: ' + e.toString());
    throw e;
  }
}

/**
 * シートへの再書き出しを行わずに、Webアプリ用JSONデータ（A2セル等）のみを更新する内部関数
 */
function saveAllResultsInternal(updatedResult) {
  const settings = getSettings();
  let iconsData = getSystemData('iconsData');
  let profileUrlsData = getSystemData('profileUrlsData');

  // アイコンデータがシステムシートにない場合は、参加者シートから復旧を試みる
  if (!iconsData || Object.keys(iconsData).length === 0) {
    Logger.log('Icons data missing in system sheet, rebuilding from participants sheet...');
    const participants = getParticipants();
    iconsData = {};
    profileUrlsData = {};
    participants.forEach(p => {
      if (p.iconUrl) iconsData[p.name] = p.iconUrl;
      if (p.profileUrl) profileUrlsData[p.name] = p.profileUrl;
    });
    // 復旧したデータを保存
    setSystemData('iconsData', iconsData);
    setSystemData('profileUrlsData', profileUrlsData);
  }

  const webAppData = {
    eventName: settings.eventName || 'はしご酒',
    parts: {
      part1: updatedResult.part1 || [],
      part2: updatedResult.part2 || [],
      part3: updatedResult.part3 || [],
      part4: updatedResult.part4 || []
    },
    partInfo: {
      part1: { label: '第1部', time: settings.part1Time || '16:50', theme: (settings.part1Theme || '').replace(/\s?チーム$/, '') },
      part2: { label: '第2部', time: settings.part2Time || '18:30', theme: (settings.part2Theme || '').replace(/\s?チーム$/, '') },
      part3: { label: '第3部', time: settings.part3Time || '20:00', theme: (settings.part3Theme || '').replace(/\s?チーム$/, '') },
      part4: {
        label: settings.exceptionCategoryName || '子連れ',
        time: settings.part4Time || '16:50',
        theme: settings.exceptionCategoryName || '子連れ'
      }
    },
    icons: iconsData,
    profileUrls: profileUrlsData,
    accounts: getAccountsMap(), // アカウント表示名のマップを追加
    timestamp: new Date().toISOString()
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  if (sheet) {
    const jsonStr = JSON.stringify(webAppData);
    // システムシートに「完全なデータ」を保存
    setSystemData('webAppFinalData', jsonStr);

    if (jsonStr.length < 50000) {
      sheet.getRange('A2').setValue(jsonStr);
      sheet.getRange('AA1').setValue(jsonStr);
    } else {
      // 50,000文字を超える場合はセルを空にする
      sheet.getRange('A2').setValue("");
      sheet.getRange('AA1').setValue("");
    }
  }
}

/**
 * Webアプリをサーブ
 */
function doGet() {
  try {
    const output = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('はしご酒 グルーピング結果')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return output;

  } catch (e) {
    Logger.log('Error in doGet: ' + e.toString());
    return HtmlService.createHtmlOutput('<p>Error: ' + e.toString() + '</p>');
  }
}

/**
 * WebAppデータを取得（クライアントサイドから呼び出し）
 */
function getWebAppData() {
  Logger.log('getWebAppData called');
  try {
    const props = PropertiesService.getScriptProperties();

    // 1. まずシステムシート（完全なデータ）を確認
    const finalData = getSystemData('webAppFinalData');
    if (finalData) {
      Logger.log('Found webAppFinalData in system sheet. Size: ' + (typeof finalData === 'string' ? finalData.length : JSON.stringify(finalData).length));
      return finalData;
    }
    Logger.log('webAppFinalData not found in system sheet');

    // 2. プロパティにない場合、または警告のみの場合はシートからフォールバック
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) {
      return { eventName: 'はしご酒', parts: { part1: [], part2: [], part3: [] }, error: '結果シートが見つかりません' };
    }

    const dataJson = sheet.getRange('A2').getValue();
    if (dataJson && String(dataJson).length > 10) {
      return JSON.parse(dataJson);
    }

    // 3. それでもデータが正しく取得できない場合の最終フォールバック
    // ここで iconsData や profileUrlsData を使って最小限のレスポンスを組み立てることも可能
    return {
      eventName: 'はしご酒',
      parts: { part1: [], part2: [], part3: [], part4: [] },
      error: 'データが巨大すぎるか、まだ生成されていません。スプレッドシートのメニューから「グルーピング実行」を再度お試しください。'
    };

  } catch (e) {
    Logger.log('Error in getWebAppData: ' + e.toString());
    return {
      eventName: 'はしご酒',
      parts: { part1: [], part2: [], part3: [], part4: [] },
      partInfo: {
        part1: { label: '第1部', time: '', theme: '' },
        part2: { label: '第2部', time: '', theme: '' },
        part3: { label: '第3部', time: '', theme: '' },
        part4: { label: '例外', time: '', theme: '' }
      },
      error: 'エラーが発生しました: ' + e.toString()
    };
  }
}

/**
 * WebアプリURLを表示
 */
function showWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutput(
      '<p><strong>WebアプリURL:</strong></p>' +
      '<input type="text" value="' + url + '" style="width: 100%; padding: 10px;" />' +
      '<p style="margin-top: 10px; font-size: 12px; color: #666;">上記URLをコピーしてブラウザで開いてください。</p>'
    ).setWidth(500).setHeight(150);
    ui.showModalDialog(html, 'WebアプリURL');
  } catch (e) {
    Logger.log('Error in showWebAppUrl: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + e.toString());
  }
}

// ===== UTILITIES =====

/**
 * 配列をシャッフル（Fisher-Yates）
 * @param {Array} arr - シャッフルする配列
 */
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

/**
 * トーストメッセージを表示
 * @param {string} message - メッセージ
 * @param {string} title - タイトル（オプション）
 */
function showToast(message, title) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'はしご酒グルーピング', 5);
}

/**
 * 参加者の名前とアカウント表示名のマッピングを取得
 */
function getAccountsMap() {
  const participants = getParticipants();
  const map = {};
  participants.forEach(p => {
    map[p.name] = p.account || p.name;
  });
  return map;
}
// ===== SYSTEM DATA STORAGE (System Sheet) =====

const SYSTEM_SHEET_NAME = '_システムデータ_';

/**
 * 隠しシート（システムデータ）にデータを保存する
 * @param {string} key - 保存キー
 * @param {any} value - 保存データ
 */
function setSystemData(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SYSTEM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_SHEET_NAME);
    sheet.hideSheet();
  }

  const jsonStr = (value === null) ? "" : (typeof value === 'string' ? value : JSON.stringify(value));
  const CHUNK_SIZE = 45000; // セルの5万文字制限に対し余裕を持たせる
  const chunks = [];

  if (jsonStr) {
    for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
      chunks.push([jsonStr.substring(i, i + CHUNK_SIZE)]);
    }
  }

  // キーを探す
  const keys = sheet.getRange("A:A").getValues();
  let rowIndex = -1;
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setValue(key);
  }

  // 既存のデータ行をクリア
  const currentCount = parseInt(sheet.getRange(rowIndex, 2).getValue() || "0");
  if (currentCount > 0) {
    sheet.getRange(rowIndex, 3, 1, currentCount).clearContent();
  }

  // 新しいデータを横に並べて保存
  if (chunks.length > 0) {
    const output = chunks.map(c => c[0]);
    sheet.getRange(rowIndex, 3, 1, output.length).setValues([output]);
    sheet.getRange(rowIndex, 2).setValue(output.length);
  } else {
    sheet.getRange(rowIndex, 2).setValue(0);
  }
}

/**
 * 隠しシート（システムデータ）からデータを取得する
 * @param {string} key - 取得キー
 * @returns {any} 復元されたデータ（オブジェクト）
 */
function getSystemData(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_SHEET_NAME);
  if (!sheet) return null;

  const keys = sheet.getRange("A:A").getValues();
  let rowIndex = -1;
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return null;

  const count = parseInt(sheet.getRange(rowIndex, 2).getValue() || "0");
  if (count === 0) return null;

  const chunks = sheet.getRange(rowIndex, 3, 1, count).getValues()[0];
  const jsonStr = chunks.join("");

  try {
    return JSON.parse(jsonStr);
  } catch (e) {
    Logger.log("Error parsing system data for " + key + ": " + e.toString());
    return null;
  }
}
