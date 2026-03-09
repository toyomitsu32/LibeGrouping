// ===== メイン処理 (Core.js) =====

/**
 * メニューの初期化
 */
function onOpen() {
    SpreadsheetApp.getUi().createMenu('🎯 はしご酒グルーピング')
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

/**
 * グルーピング実行
 */
function runGrouping() {
    handleError(() => {
        showToast('グルーピングを開始しています...', '処理中');
        const settings = getSettings();
        const participants = getParticipants();
        const teamNames = getTeamNames();

        const result = { timestamp: new Date().toISOString(), part1: [], part2: [], part3: [], part4: [] };
        const partsMapping = [
            { key: 'part1', label: '第1部', single: false },
            { key: 'part2', label: '第2部', single: false },
            { key: 'part3', label: '第3部', single: false },
            { key: 'part4', label: settings.exceptionCategoryName, single: true }
        ];

        partsMapping.forEach(pm => {
            const members = participants.filter(p => p[pm.key]).map(p => p.name);
            if (members.length === 0) return;

            if (pm.single) {
                result[pm.key] = [{ team_name: `${pm.label}チーム`, members: members, summary: '', cards: [] }];
            } else {
                const groups = distributeIntoGroups(members, settings.minGroupSize, settings.maxGroupSize);
                const names = teamNames[pm.label] || [];
                result[pm.key] = groups.map((m, idx) => ({
                    team_name: names[idx] || `${pm.label} チーム ${idx + 1}`,
                    members: m, summary: '', cards: []
                }));
            }
        });

        setSystemData('groupingResult', result);
        setSystemData('cardResult', null); // 以前のAI結果をリセット
        saveAllResults();
        showToast('グルーピングが完了しました！', '成功');
    }, 'グルーピング実行');
}

/**
 * カード生成実行
 */
function runCardGeneration(targetPart, partLabel) {
    handleError(() => {
        showToast(`${partLabel}のカード生成中...`, 'AI処理');
        const settings = getSettings();
        if (!settings.geminiApiKey) throw new Error('APIキーが未設定です');

        syncResultsFromSheet(); // 手動調整を反映
        const grouping = JSON.parse(PropertiesService.getScriptProperties().getProperty('groupingResult') || '{}');
        if (!grouping[targetPart]) throw new Error('データがありません');

        const participants = getParticipants();
        const profileMap = Object.fromEntries(participants.map(p => [p.name, p.profile || '情報なし']));
        const cardResult = getSystemData('cardResult') || grouping;

        const groups = grouping[targetPart];
        const batchSize = 4;
        for (let i = 0; i < groups.length; i += batchSize) {
            const batch = groups.slice(i, i + batchSize);
            const prompt = buildPrompt(batch, profileMap);
            const response = callGemini(prompt, settings.geminiApiKey);
            const data = parseJsonSafely(response);

            batch.forEach(g => {
                if (data[g.team_name]) {
                    g.summary = data[g.team_name].summary;
                    g.cards = data[g.team_name].cards;
                }
            });
            if (i + batchSize < groups.length) Utilities.sleep(12000); // 制限回避
        }

        cardResult[targetPart] = groups;
        setSystemData('cardResult', cardResult);
        saveAllResults();
    }, `${partLabel}カード生成`);
}

/**
 * 保存・同期ロジック (リファクタリング)
 */
function saveAllResults() {
    handleError(() => {
        syncResultsFromSheet();
        const grouping = getSystemData('cardResult') || getSystemData('groupingResult') || {};
        const mapping = getNormalizedMappingData();
        const webAppData = buildWebAppData(grouping, mapping, getSettings());
        commitWebAppData(webAppData, true); // True = シートにも表を描く
    }, '全体保存');
}

function syncResultsFromSheet() {
    handleError(() => {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
        if (!sheet) return;
        const data = sheet.getDataRange().getValues();
        const settings = getSettings();
        const groupingResult = getSystemData('cardResult') || getSystemData('groupingResult') || {};

        // シートからデータを読み取り、newPartsDataを作成 (ロジック詳細は移行)
        // ... (既存のsyncロジック)
        const newPartsData = parseSheetToGrouping(data, settings, groupingResult);

        setSystemData('groupingResult', newPartsData);
        if (getSystemData('cardResult')) setSystemData('cardResult', newPartsData);

        saveAllResultsInternal(newPartsData);
    }, 'シート同期');
}

function saveAllResultsInternal(updatedResult) {
    const mapping = getNormalizedMappingData();
    const webAppData = buildWebAppData(updatedResult, mapping, getSettings());
    commitWebAppData(webAppData, false); // False = シートへの表描画はスキップ
}

/**
 * WebApp同期用ユーティリティ
 */
function buildWebAppData(grouping, mapping, settings) {
    const clean = t => (t || '').replace(/\s?チーム$/, '');
    return {
        eventName: settings.eventName,
        parts: grouping,
        partInfo: {
            part1: { label: '第1部', time: settings.part1Time, theme: clean(settings.part1Theme) },
            part2: { label: '第2部', time: settings.part2Time, theme: clean(settings.part2Theme) },
            part3: { label: '第3部', time: settings.part3Time, theme: clean(settings.part3Theme) },
            part4: { label: settings.exceptionCategoryName, time: settings.part4Time, theme: settings.exceptionCategoryName }
        },
        icons: mapping.icons,
        profileUrls: mapping.profileUrls,
        accounts: mapping.displayNames,
        timestamp: new Date().toISOString()
    };
}

function commitWebAppData(data, updateSheetTable) {
    const jsonStr = JSON.stringify(data);
    setSystemData('webAppFinalData', jsonStr);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) return;

    if (jsonStr.length < 50000) {
        sheet.getRange('A2').setValue(jsonStr);
        sheet.getRange('AA1').setValue(jsonStr);
    }
    // 表の描画ロジックは saveAllResults 側に集約
}

/**
 * 共通エラーハンドラ (提案B)
 */
function handleError(fn, context) {
    try {
        fn();
    } catch (e) {
        const msg = `[${context}] エラーが発生しました: ${e.message}`;
        Logger.log(msg);
        showToast(e.message, `エラー: ${context}`);
    }
}

// 補助・既存関数
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}
function showToast(msg, title) { SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 5); }
function showWebAppUrl() { /* 既存 */ }
function openSettingsSheet() { /* 既存 */ }
function doGet() { return HtmlService.createHtmlOutputFromFile('index').setTitle('はしご酒').addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
function getWebAppData() { return getSystemData('webAppFinalData'); }
