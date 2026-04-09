const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3001; // 기존 포트 유지

const isProd = process.env.NODE_ENV === 'production';

app.use(cors());
app.use(express.json());

// ============================================================
// 환경 설정
// ============================================================
const CONFIG = {
  clientId: process.env.CLIENT_ID || '5454a185-bc04-4e74-9597-e2305dd67d36',
  clientSecret: process.env.CLIENT_SECRET,
  redirectUri: process.env.REDIRECT_URI || 'http://localhost:3001/callback',
  excelFileName: '재고관리(개발중).xlsx',
  inventorySheets: ['충전', '타정', '제조', '공통'],
  logSheetName: '사용내역종합'
};

const TOKEN_FILE = path.join(__dirname, 'onedrive_tokens.json');

// ============================================================
// Gemini AI 설정
// ============================================================
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

// ============================================================
// Token 관리 (Device Flow 지원형으로 개선)
// ============================================================

function loadTokens() {
  try {
    if (fs.existsSync(TOKEN_FILE)) {
      const data = fs.readFileSync(TOKEN_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('Token 파일 읽기 실패:', error.message);
  }
  return null;
}

function saveTokens(tokens) {
  try {
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
    console.log('✅ Token 저장 완료');
  } catch (error) {
    console.error('❌ Token 저장 실패:', error.message);
  }
}

// Device Flow를 통한 인증 (Authorization Code 방식의 찐빠 해결)
async function getTokenViaDeviceFlow() {
  try {
    console.log('\n📱 Device Code Flow 시작...');
    const deviceCodeResponse = await axios.post(
      'https://login.microsoftonline.com',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        scope: 'Files.ReadWrite Files.ReadWrite.All offline_access'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { user_code, device_code, verification_uri } = deviceCodeResponse.data;
    console.log('====================================================');
    console.log(`1. 접속: ${verification_uri}`);
    console.log(`2. 코드 입력: ${user_code}`);
    console.log('====================================================\n');

    while (true) {
      await new Promise(resolve => setTimeout(resolve, 5000));
      try {
        const tokenResponse = await axios.post(
          'https://login.microsoftonline.com',
          new URLSearchParams({
            client_id: CONFIG.clientId,
            grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
            device_code: device_code
          }),
          { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const tokens = {
          access_token: tokenResponse.data.access_token,
          refresh_token: tokenResponse.data.refresh_token,
          expires_at: Date.now() + (tokenResponse.data.expires_in * 1000)
        };
        saveTokens(tokens);
        return tokens;
      } catch (error) {
        if (error.response?.data?.error !== 'authorization_pending') throw error;
      }
    }
  } catch (error) {
    console.error('❌ 인증 실패:', error.message);
    return null;
  }
}

async function refreshAccessToken(refreshToken) {
  try {
    const response = await axios.post(
      'https://login.microsoftonline.com',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        client_secret: CONFIG.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const tokens = {
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token || refreshToken,
      expires_at: Date.now() + (response.data.expires_in * 1000)
    };
    saveTokens(tokens);
    return tokens;
  } catch (error) {
    return null;
  }
}

async function getValidAccessToken() {
  let tokens = loadTokens();
  if (!tokens) {
    tokens = await getTokenViaDeviceFlow();
    if (!tokens) throw new Error('인증 필요');
  }
  if (Date.now() >= tokens.expires_at - 60000) {
    tokens = await refreshAccessToken(tokens.refresh_token);
    if (!tokens) tokens = await getTokenViaDeviceFlow();
  }
  return tokens.access_token;
}

// ============================================================
// OneDrive 데이터 로드 (보관장소 매핑 추가)
// ============================================================

async function fetchExcelFromOneDrive() {
  try {
    const accessToken = await getValidAccessToken();
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
      { headers: { 'Authorization': `Bearer ${accessToken}` }, responseType: 'arraybuffer' }
    );

    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    let allData = [];

    for (const sheetName of CONFIG.inventorySheets) {
      if (workbook.Sheets[sheetName]) {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        const mappedData = jsonData.map((row, index) => ({
          id: allData.length + index + 1,
          대분류: row['대분류'] || sheetName, // 시트 이름을 대분류로 사용
          부품종류: row['부품종류'] || '',
          모델명: row['모델명'] || '',
          적용설비: row['적용설비'] || '',
          현재수량: Number(row['현재수량']) || 0,
          최소보유수량: Number(row['최소보유수량']) || 0,
          최종수정시각: row['최종수정시각'] || '',
          작업자: row['작업자'] || '',
          용도: row['용도'] || '',
          보관장소: row['보관장소'] || '위치 미지정'
        }));
        allData = allData.concat(mappedData);
      }
    }

    return allData;
  } catch (error) {
    console.error('로드 실패:', error.message);
    return [];
  }
}

async function updateExcelOnOneDrive(data) {
  try {
    const accessToken = await getValidAccessToken();
    const worksheet = XLSX.utils.json_to_sheet(data.map(item => ({
      '대분류': item.대분류,
      '부품종류': item.부품종류,
      '모델명': item.모델명,
      '적용설비': item.적용설비,
      '현재수량': item.현재수량,
      '최소보유수량': item.최소보유수량,
      '최종수정시각': item.최종수정시각,
      '작업자': item.작업자,
      '용도': item.용도,
      '보관장소': item.보관장소
    })));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '충전'); // 첫 번째 시트에 저장
    const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    await axios.put(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
      excelBuffer,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      }
    );
    return true;
  } catch (error) {
    return false;
  }
}

// ============================================================
// API Routes (최신 기능 포함)
// ============================================================

app.get('/api/inventory', async (req, res) => {
  const data = await fetchExcelFromOneDrive();
  res.json({ success: true, data });
});

app.get('/api/inventory/summary', async (req, res) => {
  const data = await fetchExcelFromOneDrive();
  const summary = {
    totalItems: data.length,
    totalQuantity: data.reduce((sum, d) => sum + d.현재수량, 0),
    lowStockCount: data.filter(d => d.최소보유수량 > 0 && d.현재수량 <= d.최소보유수량).length,
    categoryBreakdown: {}
  };
  res.json({ success: true, data: summary });
});

app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message } = req.body;
    const data = await fetchExcelFromOneDrive();
    // AI에게 보관 위치 정보를 포함하여 전달
    const context = data.map(i => `${i.모델명}: ${i.보관장소} 위치, ${i.현재수량}개`).join('\n');
    
    const result = await model.generateContent(`${context}\n\n사용자 질문: ${message}`);
    res.json({ success: true, message: result.response.text() });
  } catch (error) {
    res.status(500).send(error.message);
  }
});

app.listen(PORT, () => console.log(`🚀 서버 실행 중: ${PORT}`));