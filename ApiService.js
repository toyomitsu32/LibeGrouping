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
