// ===== 外部API通信 (ApiService.js) =====

/**
 * Gemini APIを呼び出す
 */
function callGemini(prompt, apiKey) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(apiKey);

    const payload = {
        contents: [{ parts: [{ text: prompt }] }]
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
            throw new Error(`Gemini API Error: ${responseCode} - ${responseText}`);
        }

        const result = JSON.parse(responseText);
        if (!result.candidates || result.candidates.length === 0) {
            throw new Error('No candidates in Gemini response');
        }

        return result.candidates[0].content.parts[0].text;
    } catch (e) {
        throw new Error(`callGemini failed: ${e.message}`);
    }
}

/**
 * JSON文字列を安全にパース
 */
function parseJsonSafely(text) {
    try {
        let cleaned = text.replace(/^```json\n?/, '').replace(/\n?```$/, '');
        cleaned = cleaned.replace(/^```\n?/, '').replace(/\n?```$/, '');
        return JSON.parse(cleaned.trim());
    } catch (e) {
        Logger.log(`JSON Parse Error: ${e.message}\nRaw Text: ${text}`);
        throw new Error('AIからのレスポンスを解析できませんでした。');
    }
}

/**
 * カード生成プロンプトを構築 (単一チーム集中分析版)
 */
function buildPrompt(group, profileMap) {
    const profilesText = group.members.map(name => `■ ${name}さんの自己紹介:\n${profileMap[name]}`).join('\n\n');

    return `あなたはコミュニティ分析のプロフェッショナルです。
「リベシティ（Libecity）」というコミュニティのメンバーのプロフィールデータを分析し、メンバー間の「共通点」や「意外な共通点」、「強力な繋がり」を見つけ出してください。

【対象チーム名】: ${group.team_name}

【分析対象者のプロフィール】:
${profilesText}

あなたの仕事は、これらのプロフィールを多角的に分析し、会話のきっかけになるインサイトを抽出することです。
特に以下の要素に注目してください：
1. 家族構成や境遇
2. ビジネスのフェーズや業種
3. 共通の趣味
4. 共通のリベ大での活動（お金の勉強フェス、USJオフ会など）
5. 共通のイチオシのリベ大コンテンツ
6. 共通のストレングスファインダーの傾向
7. 共通の経歴・スキル
8. 共通の体験（過去の失敗、成功体験など）

表面的な共通点だけでなく、"意外な"組み合わせや、話が弾みそうなトピックを優先してください。

【重要ルール】
- **必ず15個〜22個程度のインサイト**を出力してください。
- 些細な共通点（例：好きな食べ物、出身地域、使用ツール、性格の傾向など）でも積極的に抽出してください。
- 特定の1人のプロフィールが何度も登場するように、組み合わせを工夫してください。
- 年齢（同年代など）や性別、身体的特徴に関する共通点は除外してください。

出力は以下のJSON形式のみで返してください（Markdownブロックは不要）。

{
  "summary": "チームの盛り上がりを予感させる、総評（60文字〜100文字）",
  "cards": [
    {
      "category": "EXPERIENCE|HOBBY|BUSINESS|VALUES|OTHER",
      "title": "キャッチーで楽しい共通点のタイトル",
      "description": "クスッと笑えたり「おっ！」と思える楽しい解説文。どのメンバー同士が共通しているのか会話のネタになるように具体的に（名前入り）。",
      "members": ["該当メンバー名1", "該当メンバー名2"]
    }
  ]
}`;
}
