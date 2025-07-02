const express = require('express');
const cors = require('cors');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const PORT = process.env.PORT || 3000;

// CORS設定
// Electronのレンダラープロセスからのアクセスを許可
// 開発時はlocalhost、本番時はfile://プロトコルからのアクセスを考慮
app.use(cors({
  origin: ['http://localhost:3000', 'file://'], // 必要に応じてElectronアプリのオリジンを追加
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json());

// Gemini APIキーの取得
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;

if (!GEMINI_API_KEY) {
  console.error('GEMINI_API_KEY environment variable is not set.');
  process.exit(1);
}

const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);

// /generate-email エンドポイント
app.post('/generate-email', async (req, res) => {
  const { subject, to, purpose, bodyPoints, tone, manner } = req.body;

  if (!subject || !to || !purpose || !bodyPoints || !tone || !manner) {
    return res.status(400).json({ error: 'Missing required fields.' });
  }

  try {
    const model = genAI.getGenerativeModel({ model: "gemini-pro" });

    const prompt = `以下の情報に基づき、${tone}なトーンで${manner}のビジネスメールのドラフトを作成してください。` +
      `\n\n件名：${subject}` +
      `\n宛先：${to}` +
      `\n目的：${purpose}` +
      `\n本文に盛り込む要点：${bodyPoints.join('、')}\n\n` +
      `これらを考慮して具体的なビジネスメールを記述してください。個人情報や具体的な会社名は使用せず、一般的な表現で記述してください。署名は不要です。`;

    const result = await model.generateContent(prompt);
    const response = await result.response;
    const text = response.text();

    res.json({ emailBody: text });
  } catch (error) {
    console.error('Error generating email:', error);
    res.status(500).json({ error: 'Failed to generate email.', details: error.message });
  }
});

// /transform-text エンドポイント
app.post('/transform-text', async (req, res) => {
  const { text, type } = req.body;

  if (!text || !type) {
    return res.status(400).json({ error: 'Missing required fields.' });
  }

  if (type !== 'summary' && type !== 'bullet_points') {
    return res.status(400).json({ error: 'Invalid transformation type. Must be "summary" or "bullet_points".' });
  }

  try {
    const model = genAI.getGenerativeModel({ model: "gemini-pro" });
    let prompt = '';

    if (type === 'summary') {
      prompt = `以下の文章を簡潔に要約してください。\n\n${text}`;
    } else if (type === 'bullet_points') {
      prompt = `以下の文章を読み、重要な点を箇条書きでまとめてください。\n\n${text}`;
    }

    const result = await model.generateContent(prompt);
    const response = await result.response;
    const transformedText = response.text();

    res.json({ transformedText });
  } catch (error) {
    console.error('Error transforming text:', error);
    res.status(500).json({ error: 'Failed to transform text.', details: error.message });
  }
});

// サーバー起動
const server = app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

// Electronメインプロセスからのシャットダウンイベントを処理
process.on('SIGTERM', () => {
  console.log('SIGTERM signal received: closing HTTP server');
  server.close(() => {
    console.log('HTTP server closed.');
    process.exit(0);
  });
});

process.on('SIGINT', () => {
  console.log('SIGINT signal received: closing HTTP server');
  server.close(() => {
    console.log('HTTP server closed.');
    process.exit(0);
  });
});