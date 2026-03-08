// ===== はしご酒 グルーピングシステム - Google Apps Script =====
// 参加者データをGoogle Spreadsheetから読み込み、グルーピングアルゴリズムを実行
// Gemini APIでタグ抽出とカード生成を行い、結果をWebアプリで提供

// ===== CONFIG & MENU =====

// 定数定義
const SHEET_PARTICIPANTS = '参加者';
const SHEET_PAIRS = 'ペア指定';
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
    .addItem('⑤ 子連れチームのカード生成', 'runCardGenerationPart4')
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

  const data = sheet.getRange('A1:B11').getValues();
  const settings = {};

  const mapping = {
    1: 'geminiApiKey',
    2: 'part1Theme',
    3: 'part2Theme',
    4: 'part3Theme',
    5: 'maxGroupSize',
    6: 'minGroupSize',
    7: 'eventName',
    8: 'part1Time',
    9: 'part2Time',
    10: 'part3Time',
    11: 'part4Time'
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

  // 子連れチーム列を動的に探すためのヘッダー検索（行1〜行3を検索）
  const headersRow1 = sheet.getRange(1, 1, 1, 30).getValues()[0];
  const headersRow2 = sheet.getRange(2, 1, 1, 30).getValues()[0];
  let oViceIdx = -1;
  // 列6（G列）以降を検索
  for (let c = 6; c < 30; c++) {
    const h1 = String(headersRow1[c]).toLowerCase();
    const h2 = String(headersRow2[c]).toLowerCase();
    if (h1.includes('ovice') || h1.includes('子連れ') || h1.includes('参加') ||
      h2.includes('ovice') || h2.includes('子連れ') || h2.includes('参加')) {
      oViceIdx = c;
      break;
    }
  }
  // 見つからなかった場合、H列（index 7）をフォールバックとして使用
  if (oViceIdx === -1) {
    oViceIdx = 7; // H列
    Logger.log('子連れ列: 自動検出できなかったため、H列（8列目）をフォールバックとして使用します');
  } else {
    Logger.log('子連れ列の検索結果: 列' + (oViceIdx + 1) + '（' + String.fromCharCode(65 + oViceIdx) + '列）');
  }

  const maxCol = Math.max(16, oViceIdx + 1); // P列(16列目)まで読む
  const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, maxCol);
  const data = range.getValues();

  const participants = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const no = row[0];
    const name = String(row[1]).trim();
    const gender = String(row[2]).trim();
    const part1Cell = String(row[3]).trim();
    const part2Cell = String(row[4]).trim();
    const part3Cell = String(row[5]).trim();
    const part4Cell = oViceIdx !== -1 ? String(row[oViceIdx]).trim() : '';

    // 参加判定（特殊文字「⚪︎」や通常の「○」「〇」などを包括して検知）
    const isParticipating = (val) => {
      if (val === true) return true;
      const v = String(val).trim().toLowerCase();
      return v === 'true' || v.includes('○') || v.includes('〇') || v.includes('⚪') || v === '1' || v === 'yes' || v === '参加';
    };

    const part1 = isParticipating(part1Cell);
    const part2 = isParticipating(part2Cell);
    const part3 = isParticipating(part3Cell);
    const part4 = isParticipating(part4Cell);
    const account = String(row[12]).trim();
    const profile = String(row[13]).trim();
    const iconUrl = row[14] ? String(row[14]).trim() : ''; // O列（15列目）
    const profileUrl = row[15] ? String(row[15]).trim() : ''; // P列（16列目）

    // 名前が空の場合はスキップ
    if (!name) {
      continue;
    }

    // 少なくとも1つの部に参加している場合のみ追加
    if (part1 || part2 || part3 || part4) {
      // デバッグ用ログ：何が入っているか確認
      Logger.log(`Row: ${i}, Name: ${name}, Part1: ${part1Cell}(${part1}), Part2: ${part2Cell}(${part2}), Part3: ${part3Cell}(${part3}), Part4: ${part4Cell}(${part4})`);

      participants.push({
        no: no,
        name: name,
        gender: gender,
        part1: part1,
        part2: part2,
        part3: part3,
        part4: part4,
        account: account,
        profile: profile,
        iconUrl: iconUrl,
        profileUrl: profileUrl
      });
    }
  }

  Logger.log('===========================');
  Logger.log('Loaded ' + participants.length + ' participants');
  return participants;
}

/**
 * ペア指定を読み込む
 * @returns {Array<Array<string>>} [member1, member2]のペアの配列
 */
function getPairConstraints() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PAIRS);
  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const rowGroups = [];

  for (let i = 1; i < data.length; i++) {
    const rowMembers = data[i]
      .map(cell => String(cell).trim())
      .filter(name => name.length > 0);

    if (rowMembers.length >= 2) {
      rowGroups.push(rowMembers);
      Logger.log('  ペア指定 行' + (i + 1) + ': [' + rowMembers.join(', ') + ']');
    }
  }

  Logger.log('Loaded ' + rowGroups.length + ' group constraints from ペア指定');
  return rowGroups;
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
 * 行ごとのグループ制約を処理する（先の行を優先、同一メンバーは最初の行に所属）
 * @param {Array<Array<string>>} rowGroups - 行単位のグループ制約
 * @returns {Array<Array<string>>} 重複排除されたグループの配列
 */
/**
 * ペアグループをマージする（メンバーが重複する場合は1つの大きなグループに統合）
 * @param {Array<Array<string>>} pairGroups - ペアグループの配列
 * @returns {Array<Array<string>>} 統合後のグループ配列
 */
function mergePairConstraints(pairGroups) {
  if (pairGroups.length === 0) return [];

  const groups = pairGroups.map(g => new Set(g));
  let merged;

  do {
    merged = false;
    for (let i = 0; i < groups.length; i++) {
      for (let j = i + 1; j < groups.length; j++) {
        // 共通メンバーがいるかチェック
        const hasCommon = [...groups[i]].some(m => groups[j].has(m));
        if (hasCommon) {
          // i番目のグループにj番目を統合し、j番目を削除
          groups[j].forEach(m => groups[i].add(m));
          groups.splice(j, 1);
          merged = true;
          break;
        }
      }
      if (merged) break;
    }
  } while (merged);

  const result = groups.map(s => [...s]);
  Logger.log('Merged into ' + result.length + ' distinct groups');
  return result;
}

/**
 * 名前を正規化する（空白・不可視文字の除去、Unicode正規化）
 * @param {string} name - 正規化する名前
 * @returns {string} 正規化後の名前
 */
function normalizeName(name) {
  return String(name)
    .replace(/[\s\u00A0\u3000\u200B\uFEFF]+/g, '') // 全角スペース、NBSP、ゼロ幅スペース等を除去
    .normalize('NFC') // Unicode正規化
    .trim();
}

/**
 * ペア指定シートの名前を参加者名に名寄せする
 * 完全一致 → 正規化一致 → 部分一致 の優先順で解決
 * @param {Array<Array<string>>} pairGroups - ペアグループの配列
 * @param {Array<string>} participantNames - 参加者名の配列
 * @returns {Array<Array<string>>} 参加者名に解決されたペアグループの配列
 */
function resolvePairNames(pairGroups, participantNames) {
  // 正規化済みマップを作成
  const normalizedMap = new Map();
  for (const name of participantNames) {
    normalizedMap.set(normalizeName(name), name);
  }

  const resolved = [];

  for (const group of pairGroups) {
    const resolvedGroup = [];
    const unresolvedMembers = [];

    for (const pairName of group) {
      // 1. 完全一致
      if (participantNames.includes(pairName)) {
        resolvedGroup.push(pairName);
        continue;
      }

      // 2. 正規化一致
      const normalized = normalizeName(pairName);
      if (normalizedMap.has(normalized)) {
        const resolved = normalizedMap.get(normalized);
        Logger.log('✅ 名寄せ成功（正規化一致）: "' + pairName + '" → "' + resolved + '"');
        resolvedGroup.push(resolved);
        continue;
      }

      // 3. 部分一致（ペア指定名が参加者名に含まれる、または参加者名がペア指定名に含まれる）
      let found = false;
      for (const participantName of participantNames) {
        if (participantName.includes(pairName) || pairName.includes(participantName)) {
          Logger.log('✅ 名寄せ成功（部分一致）: "' + pairName + '" → "' + participantName + '"');
          resolvedGroup.push(participantName);
          found = true;
          break;
        }
        // 正規化後の部分一致も試す
        const normPair = normalizeName(pairName);
        const normParticipant = normalizeName(participantName);
        if (normParticipant.includes(normPair) || normPair.includes(normParticipant)) {
          Logger.log('✅ 名寄せ成功（正規化部分一致）: "' + pairName + '" → "' + participantName + '"');
          resolvedGroup.push(participantName);
          found = true;
          break;
        }
      }
      if (found) continue;

      // 解決できなかった
      unresolvedMembers.push(pairName);
      Logger.log('⚠️ 名寄せ失敗: "' + pairName + '" は参加者リストに見つかりませんでした');
    }

    if (unresolvedMembers.length > 0) {
      Logger.log('⚠️ ペアグループ [' + group.join(', ') + '] のうち [' + unresolvedMembers.join(', ') + '] が参加者と一致しません');
    }

    if (resolvedGroup.length >= 2) {
      resolved.push(resolvedGroup);
    } else {
      Logger.log('⚠️ ペアグループ [' + group.join(', ') + '] は有効メンバーが2人未満のため無視されます');
    }
  }

  Logger.log('名寄せ完了: ' + resolved.length + ' / ' + pairGroups.length + ' グループが有効');
  resolved.forEach((g, i) => Logger.log('  解決済みグループ ' + (i + 1) + ': [' + g.join(', ') + ']'));
  return resolved;
}

/**
 * ペア指定の検証を実行する
 * 「いずれか1つの部でOKなら総合OK」の仕様で判定
 * @param {Object} groupingResult - グルーピング結果 {part1: [...], part2: [...], ...}
 * @param {Array<Array<string>>} mergedPairs - マージ済みペアグループ
 * @param {Array<Object>} parts - 各部の定義 [{key, label, ...}, ...]
 * @param {Array<Object>} participants - 参加者リスト
 * @returns {Array<Object>} ペアごとの検証結果
 */
/**
 * ペア指定の検証を実行する
 * 「いずれか1つの部でOKなら総合OK」の仕様で判定
 * @param {Object} groupingResult - グルーピング結果
 * @param {Array<Object>} pairTargets - 名寄せ済みの個別のペア要望リスト
 * @param {Array<Object>} parts - 各部の定義
 * @param {Array<Object>} participants - 参加者リスト
 * @returns {Array<Object>} ペア要望ごとの検証結果
 */
function verifyPairConstraints(groupingResult, pairTargets, parts, participants) {
  const results = [];

  for (const target of pairTargets) {
    const pairMembers = target.originalGroup.join(', ');
    const resolvedGroup = target.resolvedGroup;
    const partDetails = [];
    let satisfiedPart = null;
    let satisfiedTeam = null;

    for (const part of parts) {
      if (part.singleGroup) continue;

      const partGroups = groupingResult[part.key] || [];
      if (partGroups.length === 0) continue;

      // 各メンバーの配置先を調べる
      const memberTeamMap = {};
      for (const member of resolvedGroup) {
        for (const team of partGroups) {
          if (team.members.includes(member)) {
            memberTeamMap[member] = team.team_name;
            break;
          }
        }
      }

      const teams = [...new Set(Object.values(memberTeamMap))];
      const assignedCount = Object.keys(memberTeamMap).length;
      const allSameTeam = teams.length === 1 && assignedCount === resolvedGroup.length;

      if (allSameTeam) {
        partDetails.push({ part: part.label, status: '✅', team: teams[0] });
        if (!satisfiedPart) {
          satisfiedPart = part.label;
          satisfiedTeam = teams[0];
        }
      } else {
        const detail = resolvedGroup.map(m => memberTeamMap[m] || '不参加/不在').join(',');
        partDetails.push({ part: part.label, status: '❌', team: '分散(' + detail + ')' });
      }
    }

    const overallOk = satisfiedPart !== null;
    // 詳細な状況（部:ステータス(チーム詳細)）を構築
    const partSummary = partDetails.map(d => {
      let text = d.part + ':' + d.status;
      if (d.status === '❌' && d.team.startsWith('分散')) {
        text += '(' + d.team.replace('分散(', '') + ')';
      } else if (d.status === '✅') {
        text += '(' + d.team + ')';
      }
      return text;
    }).join(' / ');

    results.push({
      pairMembers: pairMembers,
      overallStatus: overallOk ? '✅ OK' : '❌ NG',
      satisfiedPart: satisfiedPart || '—',
      satisfiedTeam: satisfiedTeam || '—',
      partSummary: partSummary
    });
  }

  return results;
}

/**
 * メンバーをグループに分配する
 * @param {Array<string>} members - メンバー名の配列
 * @param {Array<Array<string>>} mergedPairs - マージされたペアグループ
 * @param {number} minSize - グループの最小人数
 * @param {number} maxSize - グループの最大人数
 * @returns {Array<Array<string>>} グループ化されたメンバーの配列
 */
function distributeIntoGroups(members, mergedPairs, minSize, maxSize) {
  // ペア制約の考慮（上限を超えない範囲で事前にグループ化）
  let initialGroups = [];
  const usedMembers = new Set();

  const skippedPairs = [];
  const memberSet = new Set(members);

  for (const pairGroup of mergedPairs) {
    // ペアメンバーが実際のmembersリストに存在するかチェック
    const validMembers = pairGroup.filter(m => memberSet.has(m));
    const invalidMembers = pairGroup.filter(m => !memberSet.has(m));

    if (invalidMembers.length > 0) {
      Logger.log('⚠️ ペアグループ [' + pairGroup.join(', ') + '] のうち [' + invalidMembers.join(', ') + '] はこのパートの参加者ではありません');
    }

    if (validMembers.length >= 2 && validMembers.length <= maxSize) {
      initialGroups.push([...validMembers]);
      validMembers.forEach(m => usedMembers.add(m));
      Logger.log('✅ ペア制約を適用: [' + validMembers.join(', ') + ']');
    } else if (validMembers.length > maxSize) {
      // 最大人数(maxSize)を超えてしまう統合ペアは、人数制限遵守を優先してスキップする
      const names = validMembers.join(', ');
      Logger.log('❌ 人数制限遵守のためペア制約をスキップ（' + validMembers.length + '人 > 最大' + maxSize + '人）: ' + names);
      skippedPairs.push(validMembers);
    } else if (validMembers.length < 2) {
      Logger.log('⚠️ ペアグループ [' + pairGroup.join(', ') + '] は有効メンバーが2人未満のため無視');
    }
  }

  if (skippedPairs.length > 0) {
    const msgs = skippedPairs.map(g => g.join(' & ') + '（' + g.length + '人）');
    showToast('⚠️ 以下のペア指定はグループ上限(' + maxSize + '人)を超えるため反映できませんでした:\n' + msgs.join('\n'), '警告', 15);
  }

  // 残りのメンバーを取得してシャッフル
  const remaining = members.filter(m => !usedMembers.has(m));
  shuffleArray(remaining);

  // ペアで構成された初期グループに残りのメンバーを追加して均す
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

  // まず初期グループ（ペア）を、作成した枠に配置していく
  let groupIndex = 0;
  for (const initGrp of initialGroups) {
    if (groupIndex < numGroups) {
      groups[groupIndex].push(...initGrp);
      groupIndex++;
    } else {
      // もしペアの数が計算上のグループ数を超えていたら、人数の少ない枠に入れる
      const minLenGroup = groups.reduce((a, b) => a.length < b.length ? a : b);
      minLenGroup.push(...initGrp);
    }
  }

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
    const pairs = getPairConstraints();
    const teamNames = getTeamNames();

    const cleanTheme = (t) => (t || '').replace(/\s?チーム$/, '');
    const parts = [
      { key: 'part1', label: '第1部', theme: cleanTheme(settings.part1Theme), singleGroup: false },
      { key: 'part2', label: '第2部', theme: cleanTheme(settings.part2Theme), singleGroup: false },
      { key: 'part3', label: '第3部', theme: cleanTheme(settings.part3Theme), singleGroup: false },
      { key: 'part4', label: '子連れ', theme: '子連れ', singleGroup: true }
    ];

    const allParticipantNames = participants.map(p => p.name);
    const maxSize = settings.maxGroupSize || 10;

    // 各部ごとに、既に割り当てられた「マージされたペアグループ（Set）」を管理し、シミュレーションに利用する
    const partAssignments = {
      part1: [], // Array<Set<string>>
      part2: [],
      part3: []
    };

    // 各行（要望）をシートの上の行から順に処理（優先順位）
    const pairTargets = pairs.map((group, rowIdx) => {
      const resolved = resolvePairNames([group], allParticipantNames);
      if (resolved.length === 0) return null;
      const resolvedGroup = resolved[0];

      // このペアのメンバー全員が参加している部をリストアップ
      const availablePartKeys = parts.filter(p => {
        if (p.singleGroup) return false;
        return resolvedGroup.every(name => {
          const participant = participants.find(partici => partici.name === name);
          return participant && participant[p.key];
        });
      }).map(p => p.key);

      if (availablePartKeys.length === 0) return null;

      // 各利用可能な部について、このペアを追加した時の「マージ後の最大人数」を計算し、スコアリングする
      const scores = availablePartKeys.map(partKey => {
        const currentGroups = partAssignments[partKey];
        // 既存のグループ（Set）のうち、今回のペアと「連鎖的に繋がる（共通メンバーを持つ）」ものだけを抽出して合算する
        let connectedMembers = new Set(resolvedGroup);
        let changed = true;
        while (changed) {
          changed = false;
          for (const s of currentGroups) {
            const hasCommon = [...s].some(m => connectedMembers.has(m));
            if (hasCommon) {
              const oldSize = connectedMembers.size;
              s.forEach(m => connectedMembers.add(m));
              if (connectedMembers.size > oldSize) {
                changed = true;
              }
            }
          }
        }

        return {
          partKey: partKey,
          projectedSize: connectedMembers.size,
          isOver: connectedMembers.size > maxSize
        };
      });

      // スコアでソート: 1. maxSizeを超えない 2. 超えない中での最小人数（分散優先） 3. ランダム
      shuffleArray(scores);
      scores.sort((a, b) => {
        if (a.isOver !== b.isOver) return a.isOver ? 1 : -1;
        return a.projectedSize - b.projectedSize;
      });

      const best = scores[0];
      const targetPart = parts.find(p => p.key === best.partKey);

      // 割り当てを確定させ、追跡用データを正確にマージ更新（連結成分を1つのSetにまとめる）
      const currentGroups = partAssignments[best.partKey];
      let newMergedSet = new Set(resolvedGroup);
      let otherGroups = [];
      for (const s of currentGroups) {
        if ([...s].some(m => newMergedSet.has(m))) {
          s.forEach(m => newMergedSet.add(m));
        } else {
          otherGroups.push(s);
        }
      }
      otherGroups.push(newMergedSet);
      partAssignments[best.partKey] = otherGroups;

      const isReallyOver = best.isOver;
      Logger.log('🎯 要望' + (rowIdx + 1) + ' 決定: [' + resolvedGroup.join(',') + '] -> ' + targetPart.label + (isReallyOver ? ' (⚠️人数制限超過注意: ' + best.projectedSize + ')' : ' (OK)'));

      return {
        originalGroup: group,
        resolvedGroup: resolvedGroup,
        targetPartKey: targetPart.key,
        targetPartLabel: targetPart.label
      };
    }).filter(t => t !== null);

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

      // この部をターゲットにしているペア要望を抽出し、メンバーが重複する場合は統合する
      const targetingPairs = pairTargets
        .filter(t => t.targetPartKey === part.key)
        .map(t => t.resolvedGroup);

      const relevantPairs = mergePairConstraints(targetingPairs);

      // デバッグ
      Logger.log('--- ' + part.label + ' ペア適用状況 ---');
      Logger.log('  この部をターゲットにする要望数: ' + targetingPairs.length);
      Logger.log('  統合後のペア数: ' + relevantPairs.length);
      relevantPairs.forEach((g, i) => Logger.log('  部内ペア ' + (i + 1) + ': [' + g.join(', ') + ']'));

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

      // 共通グルーピング処理
      const groups = distributeIntoGroups(
        partMembers,
        relevantPairs,
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

    // 結果をスクリプトプロパティに一時保存
    const props = PropertiesService.getScriptProperties();
    props.setProperty('groupingResult', JSON.stringify(groupingResult));

    // アイコンデータを保存
    const iconsData = {};
    participants.forEach(p => {
      if (p.iconUrl) {
        iconsData[p.name] = p.iconUrl;
      }
    });
    props.setProperty('iconsData', JSON.stringify(iconsData));

    // プロフィールURLデータを保存
    const profileUrlsData = {};
    participants.forEach(p => {
      if (p.profileUrl) {
        profileUrlsData[p.name] = p.profileUrl;
      }
    });
    props.setProperty('profileUrlsData', JSON.stringify(profileUrlsData));

    // 古いカード生成結果をクリアし、新しいグルーピング結果で上書き
    props.deleteProperty('cardResult');

    // ペア指定の検証を実行
    showToast('ペア指定の検証中...', 'グルーピング');
    const verifyResults = verifyPairConstraints(groupingResult, pairTargets, parts, participants);
    props.setProperty('pairVerifyResults', JSON.stringify(verifyResults));
    Logger.log('ペア検証結果: ' + JSON.stringify(verifyResults));

    // 検証結果のサマリー（いずれか1つの部でOKならOK）
    const okCount = verifyResults.filter(r => r.overallStatus === '✅ OK').length;
    const ngCount = verifyResults.filter(r => r.overallStatus === '❌ NG').length;
    if (ngCount > 0) {
      showToast('⚠️ ペア指定 ' + ngCount + '件 がどの部でも満たされていません（' + okCount + '件 OK）', '警告', 10);
    }

    // 結果シートおよび WebApp 用データに保存
    saveAllResults();

    showToast('グルーピング完了！（ペア検証: ' + okCount + '/' + (okCount + ngCount) + ' OK）', '成功');
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
    const existingCardResultStr = props.getProperty('cardResult');
    const cardResult = existingCardResultStr ? JSON.parse(existingCardResultStr) : {
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

    // 結果をスクリプトプロパティに保存
    props.setProperty('cardResult', JSON.stringify(cardResult));

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

    const groupingResultStr = props.getProperty('cardResult') || props.getProperty('groupingResult') || '{}';
    const groupingResult = JSON.parse(groupingResultStr);
    const iconsDataStr = props.getProperty('iconsData') || '{}';
    const iconsData = JSON.parse(iconsDataStr);
    const profileUrlsDataStr = props.getProperty('profileUrlsData') || '{}';
    const profileUrlsData = JSON.parse(profileUrlsDataStr);

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
          label: '子連れ',
          time: settings.part4Time || '16:50',
          theme: '子連れ'
        }
      },
      icons: iconsData,
      profileUrls: profileUrlsData,
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
    sheet.getRange('A2').setValue(JSON.stringify(webAppData)).setFontColor('#6b7280'); // 少し文字色を薄くする

    // JSONデータはAA1にも引き続き念のため退避処理（元のgetWebAppDataが動作するように）
    sheet.getRange('AA1').setValue(JSON.stringify(webAppData));

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

    // ===== ペア指定 検証結果セクション =====
    const props2 = PropertiesService.getScriptProperties();
    const verifyStr = props2.getProperty('pairVerifyResults');
    if (verifyStr) {
      const verifyResults = JSON.parse(verifyStr);
      if (verifyResults.length > 0) {
        currentRow += 1;

        // セクションタイトル
        sheet.getRange(currentRow, 1).setValue('📋 ペア指定 検証結果（いずれか1つの部でOKなら✅）').setFontWeight('bold').setFontSize(12);
        currentRow++;

        // サマリー
        const okCount = verifyResults.filter(r => r.overallStatus === '✅ OK').length;
        const ngCount = verifyResults.filter(r => r.overallStatus === '❌ NG').length;
        const summaryText = 'ペア指定: ' + verifyResults.length + '件 ／ ✅ OK: ' + okCount + '件 ／ ❌ NG: ' + ngCount + '件';
        sheet.getRange(currentRow, 1).setValue(summaryText).setFontColor(ngCount > 0 ? '#DC2626' : '#16A34A').setFontWeight('bold');
        currentRow++;

        // テーブルヘッダー
        const verifyHeader = ['総合判定', 'OKの部', 'OKのチーム', 'ペア指定メンバー', '各部の状況'];
        const verifyCols = verifyHeader.length;
        sheet.getRange(currentRow, 1, 1, verifyCols).setValues([verifyHeader]).setFontWeight('bold').setBackground('#e5e7eb');
        currentRow++;

        // データ行
        const verifyData = verifyResults.map(r => [
          r.overallStatus,
          r.satisfiedPart,
          r.satisfiedTeam,
          r.pairMembers,
          r.partSummary
        ]);
        sheet.getRange(currentRow, 1, verifyData.length, verifyCols).setValues(verifyData);

        // 条件付き書式：OK行は緑背景、NG行は赤背景
        for (let vi = 0; vi < verifyData.length; vi++) {
          const rowNum = currentRow + vi;
          const bgColor = verifyResults[vi].overallStatus === '✅ OK' ? '#DCFCE7' : '#FEE2E2';
          sheet.getRange(rowNum, 1, 1, verifyCols).setBackground(bgColor);
        }

        // 罫線
        sheet.getRange(currentRow - 1, 1, verifyData.length + 1, verifyCols).setBorder(true, true, true, true, true, true);
        currentRow += verifyData.length + 1;
      }
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
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) {
      return {
        eventName: 'はしご酒',
        parts: { part1: [], part2: [], part3: [] },
        partInfo: {},
        tags: {},
        error: '結果シートが見つかりません'
      };
    }

    // A2にデータが配置されているため、A2を取得
    const dataJson = sheet.getRange('A2').getValue();
    if (!dataJson) {
      return {
        eventName: 'はしご酒',
        parts: { part1: [], part2: [], part3: [] },
        partInfo: {},
        tags: {},
        error: 'データがまだ生成されていません。メニューから「全工程を実行」を選択してください。'
      };
    }

    return JSON.parse(dataJson);

  } catch (e) {
    Logger.log('Error in getWebAppData: ' + e.toString());
    return {
      eventName: 'はしご酒',
      parts: { part1: [], part2: [], part3: [] },
      partInfo: {},
      tags: {},
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