// ===== スプレッドシート操作 (SheetService.js) =====

/**
 * 文字列のクリーンアップ（「チーム」の削除とトリミング）
 */
const clean = t => (t || '').toString().replace(/\s?チーム$/, '').trim();

/**
 * 設定シートから設定値を読み込む
 */
function getSettings() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
    if (!sheet) throw new Error('設定シートが見つかりません');

    // 行が挿入されてもずれないように、A列の文言（キーワード）で判定する
    const data = sheet.getDataRange().getValues();
    const settings = {};

    for (let i = 0; i < data.length; i++) {
        const label = String(data[i][0]).trim();
        const val = data[i][1];

        if (!label || val == null || val === '') continue;

        let key = null;
        if (label.includes('イベント名')) key = 'eventName';
        else if (label.includes('最大')) key = 'maxGroupSize';
        else if (label.includes('最小')) key = 'minGroupSize';
        else if ((label.includes('第1部') || label.includes('Part1')) && (label.includes('時間') || label.includes('開始'))) key = 'part1Time';
        else if ((label.includes('第2部') || label.includes('Part2')) && (label.includes('時間') || label.includes('開始'))) key = 'part2Time';
        else if ((label.includes('第3部') || label.includes('Part3')) && (label.includes('時間') || label.includes('開始'))) key = 'part3Time';
        else if ((label.includes('例外') || label.includes('子連れ') || label.includes('人数制限なし')) && (label.includes('時間') || label.includes('開始'))) key = 'part4Time';
        else if (label.includes('第1部') || label.includes('Part1')) key = 'part1Theme';
        else if (label.includes('第2部') || label.includes('Part2')) key = 'part2Theme';
        else if (label.includes('第3部') || label.includes('Part3')) key = 'part3Theme';
        else if (label.includes('例外') || label.includes('子連れ') || label.includes('人数制限なし')) key = 'exceptionCategoryName';

        if (key && settings[key] === undefined) {
            if (val instanceof Date && key.includes('Time')) {
                const h = val.getHours().toString().padStart(2, '0');
                const m = val.getMinutes().toString().padStart(2, '0');
                settings[key] = `${h}:${m}`;
            } else {
                settings[key] = String(val).trim();
            }
        }
    }

    settings.maxGroupSize = parseInt(settings.maxGroupSize) || 4;
    settings.minGroupSize = parseInt(settings.minGroupSize) || 3;
    settings.exceptionCategoryName = settings.exceptionCategoryName || '例外';
    return settings;
}

/**
 * 参加者データを読み込む（列の動的検出含む）
 */
function getParticipants() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PARTICIPANTS);
    if (!sheet) throw new Error('参加者シートが見つかりません');

    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return [];

    const settings = getSettings();
    const targetLabel = clean(settings.exceptionCategoryName).toLowerCase();

    const maxCol = 40;
    const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, maxCol);
    const data = range.getValues();
    const allHeaders = [
        sheet.getRange(1, 1, 1, maxCol).getValues()[0],
        sheet.getRange(2, 1, 1, maxCol).getValues()[0],
        sheet.getRange(3, 1, 1, maxCol).getValues()[0]
    ];

    let accountIdx = 1, iconUrlIdx = 11, profileUrlIdx = 12, profileIdx = 13, oViceIdx = 7;

    // 動的列検出
    for (let c = 0; c < maxCol; c++) {
        const combined = (allHeaders[0][c] + " " + allHeaders[1][c] + " " + allHeaders[2][c]).toLowerCase();
        if (combined.includes('nick') || combined.includes('ニックネーム') || combined.includes('表示名') || combined.includes('アカウント')) {
            if (!combined.includes('系列') && !combined.includes('カテゴリー') && !combined.includes('k列')) accountIdx = c;
        }
        if (combined.includes('画像url') || combined.includes('アイコン') || (combined.includes('画像') && combined.includes('url'))) {
            if (!combined.includes('プロフィール') && !combined.includes('プロフurl')) iconUrlIdx = c;
        }
        if ((combined.includes('プロフurl') || combined.includes('プロフィールurl')) && !combined.includes('画像')) profileUrlIdx = c;
        if (combined.includes('自己紹介') || combined.includes('本文') || (combined.includes('プロフィール') && !combined.includes('url'))) profileIdx = c;
        if (combined.includes(targetLabel)) {
            if (c >= 6) {
                oViceIdx = c;
                break; // 具体的なラベルが見つかったら確定
            }
        } else if (combined.includes('ovice') || combined.includes('参加')) {
            if (c >= 6) oViceIdx = c;
        }
    }

    const isParticipating = (val) => {
        const v = String(val).trim().toLowerCase();
        return v === 'true' || v.includes('○') || v.includes('〇') || v.includes('⚪') || v === '1' || v === 'yes' || v === '参加';
    };

    return data.map(row => {
        const name = String(row[1]).trim();
        if (!name) return null;
        const parts = [row[3], row[4], row[5], row[oViceIdx]].map(isParticipating);
        if (!parts.some(p => p)) return null;

        return {
            no: row[0], name: name, gender: String(row[2]).trim(),
            part1: parts[0], part2: parts[1], part3: parts[2], exception: parts[3],
            account: String(row[accountIdx] || '').trim() || name,
            profile: String(row[profileIdx] || '').trim(),
            iconUrl: String(row[iconUrlIdx] || '').trim(),
            profileUrl: String(row[profileUrlIdx] || '').trim()
        };
    }).filter(p => p !== null);
}

/**
 * チーム名を読み込む
 */
function getTeamNames() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SETTINGS);
    const settings = getSettings();

    const p4 = clean(settings.exceptionCategoryName) || '例外';

    const teams = { part1: [], part2: [], part3: [], exception: [] };

    const data = sheet.getDataRange().getValues();

    for (let i = 0; i < data.length; i++) {
        const rawA = String(data[i][0]).trim();
        const rawB = String(data[i][1]).trim();
        if (!rawA || !rawB || rawA.toLowerCase() === 'part') continue;

        // "第1部"や"Part1"などの設定自体を拾わないためのガード
        // 設定項目のラベル（名称、時間、人数など）が含まれている行は無視する
        if (rawA.includes('名称') || rawA.includes('時間') || rawA.includes('開始') || rawA.includes('人数') || rawA.includes('イベント')) continue;
        if (rawB.includes(':') || !isNaN(Number(rawB))) continue;

        // 数字のみ抽出
        const match = rawA.match(/\d+/);
        const num = match ? match[0] : null;

        let targetKey = null;
        if (num === '1' || rawA.toLowerCase().includes('part1') || rawA.includes('第1部')) targetKey = 'part1';
        else if (num === '2' || rawA.toLowerCase().includes('part2') || rawA.includes('第2部')) targetKey = 'part2';
        else if (num === '3' || rawA.toLowerCase().includes('part3') || rawA.includes('第3部')) targetKey = 'part3';
        else if (num === '4' || rawA.toLowerCase().includes('part4') || rawA.toLowerCase().includes('exception') ||
            rawA.includes('例外') || rawA.includes('子連れ') || rawA.includes('コズレ') || rawA.includes('人数制限なし')) targetKey = 'exception';

        if (targetKey) {
            teams[targetKey].push(rawB);
        }
    }

    // デバッグ用ログ: 読み取れたチーム名の数を出力
    Logger.log('TeamNames loaded: ' + Object.keys(teams).map(k => `${k}:${teams[k].length}`).join(', '));

    return teams;
}

/**
 * システム用データの保存・取得
 */
function setSystemData(key, value) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SYSTEM_SHEET_NAME);
    if (!sheet) {
        sheet = ss.insertSheet(SYSTEM_SHEET_NAME);
        sheet.hideSheet();
    }
    const jsonStr = (value === null) ? "" : (typeof value === 'string' ? value : JSON.stringify(value));
    const CHUNK_SIZE = 45000;
    const chunks = [];
    if (jsonStr) {
        for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) chunks.push([jsonStr.substring(i, i + CHUNK_SIZE)]);
    }

    const keys = sheet.getRange("A:A").getValues();
    let rowIndex = keys.findIndex(r => r[0] === key) + 1;
    if (rowIndex <= 0) {
        rowIndex = sheet.getLastRow() + 1;
        sheet.getRange(rowIndex, 1).setValue(key);
    }

    const currentCount = parseInt(sheet.getRange(rowIndex, 2).getValue() || "0");
    if (currentCount > 0) sheet.getRange(rowIndex, 3, 1, currentCount).clearContent();

    if (chunks.length > 0) {
        const output = chunks.map(c => c[0]);
        sheet.getRange(rowIndex, 3, 1, output.length).setValues([output]);
        sheet.getRange(rowIndex, 2).setValue(output.length);
    } else {
        sheet.getRange(rowIndex, 2).setValue(0);
    }
}

function getSystemData(key) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SYSTEM_SHEET_NAME);
    if (!sheet) return null;
    const keys = sheet.getRange("A:A").getValues();
    const rowIndex = keys.findIndex(r => r[0] === key) + 1;
    if (rowIndex <= 0) return null;

    const count = parseInt(sheet.getRange(rowIndex, 2).getValue() || "0");
    if (count === 0) return null;
    const data = sheet.getRange(rowIndex, 3, 1, count).getValues()[0].join('');
    try { return JSON.parse(data); } catch (e) { return data; }
}

/**
 * データの正規化 (Single Source of Truth) - 提案A
 */
function getNormalizedMappingData() {
    const participants = getParticipants();
    const icons = {};
    const profileUrls = {};
    const displayNames = {};

    participants.forEach(p => {
        // 表示名の確定
        displayNames[p.name] = p.account || p.name;

        // アイコンURLの確定（Fallback込み）
        let icon = p.iconUrl;
        if (!icon || !icon.toLowerCase().startsWith('http')) {
            icon = `https://ui-avatars.com/api/?name=${encodeURIComponent(displayNames[p.name])}&background=random`;
        }
        icons[p.name] = icon;

        // プロフィールURLの確定
        profileUrls[p.name] = (p.profileUrl && p.profileUrl.toLowerCase().startsWith('http')) ? p.profileUrl : '';
    });

    return { icons, profileUrls, displayNames };
}

/**
 * ユーザープロパティからAPIキーを取得
 */
function getUserApiKey() {
    return PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
}

/**
 * ユーザープロパティにAPIキーを保存
 */
function setUserApiKey(key) {
    if (key) {
        PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', key.trim());
    }
}

/**
 * シートの表形式データからグルーピング構造を解析する
 */
function parseSheetToGrouping(data, settings, groupingResult) {
    const p1 = clean(settings.part1Theme) || '第1部';
    const p2 = clean(settings.part2Theme) || '第2部';
    const p3 = clean(settings.part3Theme) || '第3部';
    const p4 = clean(settings.exceptionCategoryName) || '例外';

    const partsMapping = [
        { key: 'part1', patterns: [p1, '第1部', 'Part1', 'Part 1'] },
        { key: 'part2', patterns: [p2, '第2部', 'Part2', 'Part 2'] },
        { key: 'part3', patterns: [p3, '第3部', 'Part3', 'Part 3'] },
        { key: 'exception', patterns: [p4, '例外', 'Exception', 'Part4', 'Part 4', '子連れ', 'コズレ'] }
    ];

    let currentPartKey = null;
    const newPartsData = { part1: [], part2: [], part3: [], exception: [] };

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const firstCell = String(row[0]).trim();
        if (!firstCell) continue;

        // システムヘッダーや表頭をスキップ
        if (firstCell.startsWith('🎊') ||
            firstCell.includes('チーム名') || firstCell.includes('人数') ||
            firstCell.includes('区分') || firstCell.includes('総合判定') ||
            firstCell.includes('システム用データ')) {
            continue;
        }

        const members = [];
        // C列以降にメンバー名が入っている
        for (let c = 2; c < row.length; c++) {
            const mName = String(row[c]).trim();
            // "1名" などの人数セルや空セルを除外
            if (mName && !mName.endsWith('名')) members.push(mName);
        }

        if (members.length === 0) {
            // メンバーが1人もいない行はセクションヘッダー（各部の区切り）の可能性を探る
            for (const mapping of partsMapping) {
                const isMatch = mapping.patterns.some(p => {
                    if (!p) return false;
                    const cleanP = p.replace(/[【】]/g, '');
                    return firstCell.includes(p) || firstCell.includes(cleanP);
                });

                if (isMatch) {
                    currentPartKey = mapping.key;
                    break;
                }
            }
        } else {
            // メンバーが存在する行はチーム行として処理
            if (currentPartKey) {
                const teamName = firstCell;
                // インサイト情報を維持するために既存データを参照
                const existingTeams = groupingResult[currentPartKey] || [];
                const existingTeam = existingTeams.find(t => clean(t.team_name) === clean(teamName));

                newPartsData[currentPartKey].push({
                    team_name: teamName,
                    members: members,
                    summary: existingTeam ? existingTeam.summary : '',
                    cards: existingTeam ? existingTeam.cards : []
                });
            }
        }
    }
    return newPartsData;
}

