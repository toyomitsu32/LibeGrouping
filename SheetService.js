// ===== スプレッドシート操作 (SheetService.js) =====

/**
 * 設定シートから設定値を読み込む
 */
function getSettings() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
    if (!sheet) throw new Error('設定シートが見つかりません');

    const data = sheet.getRange('A1:B13').getValues();
    const settings = {};
    const mapping = {
        2: 'eventName',
        3: 'maxGroupSize',
        4: 'minGroupSize',
        5: 'part1Theme',
        6: 'part2Theme',
        7: 'part3Theme',
        8: 'exceptionCategoryName',
        // 9: 運営など（予備）
        10: 'part1Time',
        11: 'part2Time',
        12: 'part3Time',
        13: 'part4Time'
    };

    for (const [row, key] of Object.entries(mapping)) {
        const val = data[parseInt(row) - 1][1];
        if (val != null && val !== '') {
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
    settings.exceptionCategoryName = settings.exceptionCategoryName || '子連れ';
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
    const targetLabel = settings.exceptionCategoryName.toLowerCase();

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
        if (combined.includes('ovice') || combined.includes(targetLabel) || combined.includes('参加')) {
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
            part1: parts[0], part2: parts[1], part3: parts[2], part4: parts[3],
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEAMS);
    const data = sheet ? sheet.getDataRange().getValues() : [];
    const teams = { '第1部': [], '第2部': [], '第3部': [] };
    for (let i = 1; i < data.length; i++) {
        const part = String(data[i][0]).trim();
        const name = String(data[i][1]).trim();
        if (part && name && teams[part]) teams[part].push(name);
    }
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
    const exceptionLabel = settings.exceptionCategoryName || '子連れ';
    const partsLabelToKey = {
        '【第1部】': 'part1', '【第2部】': 'part2', '【第3部】': 'part3',
        [`【${exceptionLabel}】`]: 'part4'
    };

    let currentPartKey = null;
    const newPartsData = { part1: [], part2: [], part3: [], part4: [] };

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const firstCell = String(row[0]).trim();

        // 各部のセクション開始を検知
        for (const [label, key] of Object.entries(partsLabelToKey)) {
            if (firstCell.startsWith(label)) {
                currentPartKey = key;
                i++; // ヘッダー行をスキップ
                break;
            }
        }

        if (currentPartKey && firstCell && !partsLabelToKey[firstCell] &&
            firstCell !== 'チーム名' && firstCell !== '人数' && firstCell !== '総合判定' && !firstCell.startsWith('🎊')) {

            const teamName = firstCell;
            const members = [];
            // C列以降にメンバー名が入っている
            for (let c = 2; c < row.length; c++) {
                const mName = String(row[c]).trim();
                if (mName) members.push(mName);
            }

            if (members.length > 0) {
                // 既存のAI生成結果があれば引き継ぐ
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
    return newPartsData;
}
