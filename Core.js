// ===== メイン処理 (Core.js) =====

/**
 * メニューの初期化
 */
function onOpen() {
    SpreadsheetApp.getUi().createMenu('🎯 はしご酒グルーピング')
        .addItem('① グルーピング実行', 'runGrouping')
        .addSeparator()
        .addItem('② 第1部のプロフィールから共通の話題を見つける', 'runCardGenerationPart1')
        .addItem('③ 第2部のプロフィールから共通の話題を見つける', 'runCardGenerationPart2')
        .addItem('④ 第3部のプロフィールから共通の話題を見つける', 'runCardGenerationPart3')
        .addItem('⑤ 例外チームのプロフィールから共通の話題を見つける', 'runCardGenerationPart4')
        .addSeparator()
        .addItem('🔄 手動調整をWebアプリに反映', 'syncResultsFromSheet')
        .addSeparator()
        .addItem('🌐 WebアプリURLを表示', 'showWebAppUrl')
        .addItem('⚙️ 設定シートを開く', 'openSettingsSheet')
        .addItem('🔧 設定シートの項目を整理する', 'reformatSettingsSheet')
        .addSeparator()
        .addItem('🔑 APIキーを設定・再設定する', 'promptForApiKey')
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
        saveAllResults(true); // True = 同期スキップ（計算したての結果をそのまま保存）
        showToast('グルーピングが完了しました！', '成功');
    }, 'グルーピング実行');
}

/**
 * メンバーをグループに分配する
 */
function distributeIntoGroups(members, minSize, maxSize) {
    const remaining = [...members];
    shuffleArray(remaining);
    const totalMembers = members.length;
    let numGroups = Math.max(1, Math.ceil(totalMembers / maxSize));

    if (numGroups > 1 && numGroups * minSize > totalMembers) {
        numGroups = Math.floor(totalMembers / minSize);
        if (numGroups === 0) numGroups = 1;
    }

    const groups = Array.from({ length: numGroups }, () => []);
    for (const member of remaining) {
        const minLenGroup = groups.reduce((a, b) => a.length < b.length ? a : b);
        minLenGroup.push(member);
    }
    return groups;
}

/**
 * カード生成実行の各部ラッパー
 */
function runCardGenerationPart1() { runCardGeneration('part1', '第1部'); }
function runCardGenerationPart2() { runCardGeneration('part2', '第2部'); }
function runCardGenerationPart3() { runCardGeneration('part3', '第3部'); }
function runCardGenerationPart4() { runCardGeneration('part4', '例外チーム'); }

/**
 * カード生成実行本体
 */
function runCardGeneration(targetPart, partLabel) {
    handleError(() => {
        let apiKey = getUserApiKey();
        if (!apiKey) {
            apiKey = promptForApiKey();
            if (!apiKey) return; // キャンセルされた場合
        }
        showToast(`${partLabel}の共通の話題を分析中...`, 'AI処理');
        const settings = getSettings();

        syncResultsFromSheet(); // 手動調整を反映
        const groupingStr = getSystemData('groupingResult');
        const grouping = typeof groupingStr === 'string' ? JSON.parse(groupingStr) : groupingStr;
        if (!grouping || !grouping[targetPart]) throw new Error('データがありません。先にグルーピングを実行してください。');

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
        saveAllResults(true); // 更新した内容をそのまま保存
        showToast(`${partLabel}の共通の話題が見つかりました！`, '成功');
    }, `${partLabel}分析`);
}

/**
 * 保存・同期ロジック
 */
function saveAllResults(skipSync = false) {
    handleError(() => {
        if (!skipSync) {
            syncResultsFromSheet();
        }
        const grouping = getSystemData('cardResult') || getSystemData('groupingResult') || {};
        const mapping = getNormalizedMappingData();
        const webAppData = buildWebAppData(grouping, mapping, getSettings());

        // WebApp用JSONを保存
        commitWebAppData(webAppData);

        // ヒューマンリーダブルな表の描画
        paintResultsToSheet(webAppData);
    }, '全体保存');
}

function syncResultsFromSheet() {
    handleError(() => {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
        if (!sheet) return;
        const data = sheet.getDataRange().getValues();
        const settings = getSettings();
        const groupingResult = getSystemData('cardResult') || getSystemData('groupingResult') || { part1: [], part2: [], part3: [], part4: [] };

        const newPartsData = parseSheetToGrouping(data, settings, groupingResult);

        setSystemData('groupingResult', newPartsData);
        if (getSystemData('cardResult')) setSystemData('cardResult', newPartsData);

        saveAllResultsInternal(newPartsData);
    }, 'シート同期');
}

function saveAllResultsInternal(updatedResult) {
    const mapping = getNormalizedMappingData();
    const webAppData = buildWebAppData(updatedResult, mapping, getSettings());
    commitWebAppData(webAppData);
}

/**
 * WebApp同期用ユーティリティ
 */
function buildWebAppData(grouping, mapping, settings) {
    const clean = t => (t || '').replace(/\s?チーム$/, '');
    return {
        eventName: settings.eventName,
        parts: {
            part1: grouping.part1 || [],
            part2: grouping.part2 || [],
            part3: grouping.part3 || [],
            part4: grouping.part4 || []
        },
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

function commitWebAppData(data) {
    const jsonStr = JSON.stringify(data);
    setSystemData('webAppFinalData', jsonStr);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) return;

    if (jsonStr.length < 50000) {
        sheet.getRange('A2').setValue(jsonStr);
        sheet.getRange('AA1').setValue(jsonStr);
    } else {
        sheet.getRange('A2').setValue("");
        sheet.getRange('AA1').setValue("");
    }
}

/**
 * 結果をスプレッドシートに表形式で描画
 */
function paintResultsToSheet(webAppData) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
    if (!sheet) return;

    sheet.clear();
    sheet.getRange('A1').setValue('【システム用データ】以下の行のデータはWebアプリが読み込むものです').setBackground('#fef08a').setFontWeight('bold');

    const jsonStr = JSON.stringify(webAppData);
    if (jsonStr.length < 50000) {
        sheet.getRange('A2').setValue(jsonStr).setFontColor('#6b7280');
        sheet.getRange('AA1').setValue(jsonStr);
    }

    let currentRow = 4;
    sheet.getRange(currentRow, 1).setValue(`🎊 ${webAppData.eventName} グルーピング結果 🎊`).setFontWeight('bold').setFontSize(14);
    currentRow += 2;

    const partsKeys = ['part1', 'part2', 'part3', 'part4'];
    for (const key of partsKeys) {
        const partGroups = webAppData.parts[key];
        if (!partGroups || partGroups.length === 0) continue;
        const partInfo = webAppData.partInfo[key];

        sheet.getRange(currentRow, 1).setValue(`【${partInfo.label}】 ${partInfo.time} 〜 （テーマ：${partInfo.theme || 'なし'}）`).setFontWeight('bold').setBackground('#f3f4f6');
        currentRow++;

        const maxMembers = 10;
        const headerRow = ['チーム名', '人数', ...Array.from({ length: maxMembers }, (_, i) => 'メンバー' + (i + 1))];
        sheet.getRange(currentRow, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold').setBackground('#e5e7eb');
        currentRow++;

        const outputData = partGroups.map(group => [
            group.team_name,
            `${group.members.length}名`,
            ...Array.from({ length: maxMembers }, (_, m) => group.members[m] || '')
        ]);

        if (outputData.length > 0) {
            sheet.getRange(currentRow, 1, outputData.length, headerRow.length).setValues(outputData);
            sheet.getRange(currentRow - 1, 1, outputData.length + 1, headerRow.length).setBorder(true, true, true, true, true, true);
            currentRow += outputData.length + 2;
        }
    }

    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 60);
    for (let c = 3; c <= 12; c++) sheet.setColumnWidth(c, 200);
}

/**
 * 共通エラーハンドラ
 */
function handleError(fn, context) {
    try {
        fn();
    } catch (e) {
        const msg = `[${context}] エラーが発生しました: ${e.message}`;
        Logger.log(msg + "\nStack: " + e.stack);
        showToast(e.message, `エラー: ${context}`);
    }
}

// 共通ユーティリティ
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}
function showToast(msg, title) { SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 5); }
function showWebAppUrl() {
    const url = ScriptApp.getService().getUrl();
    const html = `<p>WebアプリのURLはこちらです：</p><p><a href="${url}" target="_blank">${url}</a></p>`;
    const output = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(output, 'WebアプリURL');
}
function openSettingsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (sheet) ss.setActiveSheet(sheet);
}

/**
 * 設定シートの項目名（A列）を新しい構成に並べ替える
 */
function reformatSettingsSheet() {
    handleError(() => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEET_SETTINGS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_SETTINGS);
        }

        const labels = [
            ['【APIキー】', '※メニューから個別に設定してください'],
            ['イベント名', ''],
            ['最大グループ人数', '4'],
            ['最小グループ人数', '3'],
            ['第1部のテーマ', ''],
            ['第2部のテーマ', ''],
            ['第3部のテーマ', ''],
            ['例外カテゴリー名', '子連れ'],
            ['運営など', ''],
            ['第1部 開始時間', '18:00'],
            ['第2部 開始時間', '19:30'],
            ['第3部 開始時間', '21:00'],
            ['例外部 開始時間', '18:00']
        ];

        sheet.getRange(1, 1, labels.length, 2).setValues(labels);
        sheet.setColumnWidth(1, 200);
        sheet.setColumnWidth(2, 300);
        sheet.getRange('A1:A13').setBackground('#f3f4f6').setFontWeight('bold');
        sheet.getRange('B1').setFontColor('#9ca3af'); // APIキーの注意書きを薄く

        ss.setActiveSheet(sheet);
        showToast('設定シートの項目を整理しました。値を入力してください。', '完了');
    }, '設定シート整理');
}

/**
 * APIキーの入力を求めるダイアログを表示
 */
function promptForApiKey() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Gemini APIキーの設定',
        'あなたのGemini APIキーを入力してください。\nこのキーはあなた自身のGoogleアカウント内（ユーザープロパティ）に安全に保存され、他のユーザーには見えません。',
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.OK) {
        const key = response.getResponseText().trim();
        if (key) {
            setUserApiKey(key);
            showToast('APIキーを保存しました。', '設定完了');
            return key;
        } else {
            ui.alert('キーが空です。設定をキャンセルしました。');
        }
    }
    return null;
}
function doGet() { return HtmlService.createHtmlOutputFromFile('index').setTitle('はしご酒').addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
function getWebAppData() { return getSystemData('webAppFinalData'); }
