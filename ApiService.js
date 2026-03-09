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
 * カード生成プロンプトを構築
 */
function buildPrompt(batchedGroups, profileMap) {
    const teamsText = batchedGroups.map(group => {
        return `チーム名: ${group.team_name}
メンバーの自己紹介文:
${group.members.map(name => `■ ${name}さんの自己紹介:\n${profileMap[name]}`).join('\n\n')}`;
    }).join('\n\n======\n\n');

    return `以下の複数チームのメンバープロフィールを分析し、**各チームごと**に「楽しい共通点カード」を【概ね6枚】作成してください。

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
}
