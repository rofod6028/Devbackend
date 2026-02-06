const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

// ============================================================
// 환경 설정
// ============================================================
const CONFIG = {
  clientId: '5454a185-bc04-4e74-9597-e2305dd67d36',
  clientSecret: 'Se98Q~SelMSaSB.Euko66Qqcny7wgcpuWy10ZbB0',
  redirectUri: 'http://localhost:3001/callback',
  authCode: 'M.C522_BAY.2.U.1599003d-a632-75b7-8461-d84093d4f45a',
  excelFileName: '재고관리.xlsx',
  sheetName: '재고관리'
};

// Refresh Token 저장 파일
const TOKEN_FILE = path.join(__dirname, 'onedrive_tokens.json');

// ============================================================
// Gemini AI 설정
// ============================================================
const genAI = new GoogleGenerativeAI('AIzaSyDulQlx2CxbO5foZIFyghq25UpQhrod-Qw');
const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });

// ============================================================
// Token 관리
// ============================================================

// Token 파일 읽기
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

// Token 파일 저장
function saveTokens(tokens) {
  try {
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
    console.log('✅ Token 저장 완료');
  } catch (error) {
    console.error('❌ Token 저장 실패:', error.message);
  }
}

// Authorization Code로 최초 Token 발급
async function getInitialTokens() {
  try {
    console.log('🔑 최초 Token 발급 시도...');
    
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        client_secret: CONFIG.clientSecret,
        code: CONFIG.authCode,
        redirect_uri: CONFIG.redirectUri,
        grant_type: 'authorization_code'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      }
    );

    const tokens = {
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token,
      expires_at: Date.now() + (response.data.expires_in * 1000)
    };

    saveTokens(tokens);
    console.log('✅ 최초 Token 발급 성공!');
    return tokens;

  } catch (error) {
    console.error('❌ 최초 Token 발급 실패:', error.response?.data || error.message);
    return null;
  }
}

// Refresh Token으로 Access Token 갱신
async function refreshAccessToken(refreshToken) {
  try {
    console.log('🔄 Access Token 갱신 중...');
    
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        client_secret: CONFIG.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      }
    );

    const tokens = {
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token || refreshToken,
      expires_at: Date.now() + (response.data.expires_in * 1000)
    };

    saveTokens(tokens);
    console.log('✅ Access Token 갱신 성공!');
    return tokens;

  } catch (error) {
    console.error('❌ Token 갱신 실패:', error.response?.data || error.message);
    return null;
  }
}

// 유효한 Access Token 가져오기
async function getValidAccessToken() {
  let tokens = loadTokens();

  // 저장된 Token이 없으면 최초 발급
  if (!tokens) {
    tokens = await getInitialTokens();
    if (!tokens) {
      throw new Error('Token 발급 실패. Authorization Code를 다시 발급받아야 합니다.');
    }
  }

  // Token이 만료되었으면 갱신
  if (Date.now() >= tokens.expires_at - 60000) { // 1분 여유
    tokens = await refreshAccessToken(tokens.refresh_token);
    if (!tokens) {
      throw new Error('Token 갱신 실패. Authorization Code를 다시 발급받아야 합니다.');
    }
  }

  return tokens.access_token;
}

// ============================================================
// OneDrive 엑셀 파일 읽기
// ============================================================

let cachedData = null;
let lastFetchTime = null;
const CACHE_DURATION = 60 * 1000; // 1분 캐시

async function fetchExcelFromOneDrive() {
  const now = Date.now();
  if (cachedData && lastFetchTime && (now - lastFetchTime) < CACHE_DURATION) {
    console.log('📦 캐시된 데이터 사용');
    return cachedData;
  }

  try {
    const accessToken = await getValidAccessToken();

    console.log(`📥 OneDrive에서 "${CONFIG.excelFileName}" 다운로드 중...`);

    // OneDrive 파일 다운로드
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        responseType: 'arraybuffer'
      }
    );

    // 엑셀 파싱
    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    const worksheet = workbook.Sheets[CONFIG.sheetName];

    if (!worksheet) {
      console.error(`❌ 시트 "${CONFIG.sheetName}"를 찾을 수 없습니다.`);
      console.log(`사용 가능한 시트: ${workbook.SheetNames.join(', ')}`);
      return getDummyData();
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    // 필드 매핑
    const mappedData = jsonData.map((row, index) => ({
      id: index + 1,
      부품종류: row['부품종류'] || '',
      모델명: row['모델명'] || '',
      적용설비: row['적용설비'] || '',
      현재수량: Number(row['현재수량']) || 0,
      최소보유수량: Number(row['최소보유수량']) || 0,
      최종수정시각: row['최종수정시각'] || ''
    }));

    cachedData = mappedData;
    lastFetchTime = now;

    console.log(`✅ OneDrive 엑셀 데이터 로드 완료: ${mappedData.length}건`);
    return mappedData;

  } catch (error) {
    console.error('❌ OneDrive 엑셀 읽기 실패:', error.response?.data || error.message);
    console.log('⚠️ 테스트용 더미 데이터 반환 중...');
    return getDummyData();
  }
}

// ============================================================
// OneDrive 엑셀 파일 쓰기 (수정)
// ============================================================

async function updateExcelOnOneDrive(data) {
  try {
    const accessToken = await getValidAccessToken();

    console.log(`📤 OneDrive에 "${CONFIG.excelFileName}" 업로드 중...`);

    // 데이터를 엑셀 형식으로 변환
    const worksheet = XLSX.utils.json_to_sheet(data.map(item => ({
      '부품종류': item.부품종류,
      '모델명': item.모델명,
      '적용설비': item.적용설비,
      '현재수량': item.현재수량,
      '최소보유수량': item.최소보유수량,
      '최종수정시각': item.최종수정시각
    })));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, CONFIG.sheetName);

    // 엑셀 파일을 버퍼로 변환
    const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // OneDrive에 업로드
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

    console.log('✅ OneDrive 엑셀 업데이트 완료!');
    
    // 캐시 갱신
    cachedData = data;
    lastFetchTime = Date.now();

    return true;

  } catch (error) {
    console.error('❌ OneDrive 엑셀 쓰기 실패:', error.response?.data || error.message);
    return false;
  }
}

// ============================================================
// 테스트용 더미 데이터
// ============================================================
function getDummyData() {
  return [
    { id: 1, 부품종류: '베어링', 모델명: 'SKF-6205', 적용설비: '펌프A', 현재수량: 15, 최소보유수량: 5, 최종수정시각: '2026-02-03 09:00' },
    { id: 2, 부품종류: '베어링', 모델명: 'SKF-6304', 적용설비: '펌프B', 현재수량: 3, 최소보유수량: 5, 최종수정시각: '2026-02-02 14:30' },
    { id: 3, 부품종류: '베어링', 모델명: 'NSK-7205', 적용설비: '모터C', 현재수량: 8, 최소보유수량: 3, 최종수정시각: '2026-02-01 11:00' },
    { id: 4, 부품종류: '오일필터', 모델명: 'MANN-W940', 적용설비: '컴프레서1', 현재수량: 20, 최소보유수량: 8, 최종수정시각: '2026-02-03 08:00' },
    { id: 5, 부품종류: '오일필터', 모델명: 'MANN-W1060', 적용설비: '컴프레서2', 현재수량: 4, 최소보유수량: 6, 최종수정시각: '2026-01-30 16:00' },
    { id: 6, 부품종류: '오일필터', 모델명: 'Donaldson-P551', 적용설비: '펌프D', 현재수량: 10, 최소보유수량: 4, 최종수정시각: '2026-02-02 10:15' },
    { id: 7, 부품종류: '벨트', 모델명: 'Gates-A68', 적용설비: '모터A', 현재수량: 6, 최소보유수량: 3, 최종수정시각: '2026-02-01 09:30' },
    { id: 8, 부품종류: '벨트', 모델명: 'Gates-B82', 적용설비: '모터B', 현재수량: 2, 최소보유수량: 4, 최종수정시각: '2026-01-28 13:00' },
    { id: 9, 부품종류: '패킹', 모델명: 'Teikoku-S1', 적용설비: '펌프A', 현재수량: 30, 최소보유수량: 10, 최종수정시각: '2026-02-03 07:45' },
    { id: 10, 부품종류: '패킹', 모델명: 'Teikoku-S2', 적용설비: '펌프B', 현재수량: 5, 최소보유수량: 8, 최종수정시각: '2026-01-25 11:20' },
    { id: 11, 부품종류: '볼트/너트', 모델명: 'M12-SUS304', 적용설비: '구조체1', 현재수량: 100, 최소보유수량: 30, 최종수정시각: '2026-02-02 15:00' },
    { id: 12, 부품종류: '볼트/너트', 모델명: 'M16-SUS316', 적용설비: '구조체2', 현재수량: 25, 최소보유수량: 20, 최종수정시각: '2026-02-01 08:00' },
    { id: 13, 부품종류: '감속기', 모델명: 'SEW-R57', 적용설비: '컨베이어1', 현재수량: 4, 최소보유수량: 2, 최종수정시각: '2026-01-29 14:00' },
    { id: 14, 부품종류: '감속기', 모델명: 'SEW-R67', 적용설비: '컨베이어2', 현재수량: 1, 최소보유수량: 2, 최종수정시각: '2026-01-20 09:00' },
  ];
}

// ============================================================
// API Routes
// ============================================================

// [GET] 전체 재고 데이터
app.get('/api/inventory', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    res.json({ success: true, data });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// [GET] 부품종류별 그룹화된 데이터
app.get('/api/inventory/categories', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const categories = {};

    data.forEach(item => {
      if (!categories[item.부품종류]) {
        categories[item.부품종류] = {
          name: item.부품종류,
          totalCount: 0,
          itemCount: 0,
          lowStockCount: 0,
          items: []
        };
      }
      categories[item.부품종류].items.push(item);
      categories[item.부품종류].totalCount += item.현재수량;
      categories[item.부품종류].itemCount += 1;
      if (item.현재수량 <= item.최소보유수량) {
        categories[item.부품종류].lowStockCount += 1;
      }
    });

    res.json({ success: true, data: Object.values(categories) });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// [GET] 특정 부품종류의 상세 리스트
app.get('/api/inventory/category/:categoryName', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const filtered = data.filter(item => item.부품종류 === req.params.categoryName);
    res.json({ success: true, data: filtered });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// [GET] 전체 사용량 요약
app.get('/api/inventory/summary', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();

    const summary = {
      totalItems: data.length,
      totalQuantity: data.reduce((sum, d) => sum + d.현재수량, 0),
      lowStockItems: data.filter(d => d.현재수량 <= d.최소보유수량),
      lowStockCount: data.filter(d => d.현재수량 <= d.최소보유수량).length,
      categoryBreakdown: {}
    };

    data.forEach(item => {
      if (!summary.categoryBreakdown[item.부품종류]) {
        summary.categoryBreakdown[item.부품종류] = { total: 0, count: 0, lowStock: 0 };
      }
      summary.categoryBreakdown[item.부품종류].total += item.현재수량;
      summary.categoryBreakdown[item.부품종류].count += 1;
      if (item.현재수량 <= item.최소보유수량) {
        summary.categoryBreakdown[item.부품종류].lowStock += 1;
      }
    });

    res.json({ success: true, data: summary });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// [POST] 재고 수량 업데이트 (예시)
app.post('/api/inventory/update', async (req, res) => {
  try {
    const { id, 현재수량 } = req.body;
    
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id === id);
    
    if (!item) {
      return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
    }

    // 수량 업데이트
    item.현재수량 = 현재수량;
    item.최종수정시각 = new Date().toLocaleString('ko-KR');

    // OneDrive에 저장
    const success = await updateExcelOnOneDrive(data);

    if (success) {
      res.json({ success: true, message: '업데이트 완료', data: item });
    } else {
      res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }

  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// ============================================================
// [POST] Gemini AI 채팅
// ============================================================
app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message, conversationHistory } = req.body;
    const inventoryData = await fetchExcelFromOneDrive();

    const inventoryContext = `
현재 스페어파츠 재고 상황:
${JSON.stringify(inventoryData, null, 2)}

규칙:
- 현재수량 ≤ 최소보유수량이면 → 재고 부족 상태
- 재고 부족 시 입고 권유
- 사용자가 입출고를 요청하면 구체적인 권유를 해주세요
- 질문에 대해 정확하고 간결하게 답변하세요
- 한국어로 답변해주세요
`;

    const contents = [];

    if (conversationHistory && conversationHistory.length > 0) {
      conversationHistory.forEach(msg => {
        contents.push({
          role: msg.role,
          parts: [{ text: msg.text }]
        });
      });
    }

    contents.push({
      role: 'user',
      parts: [{ text: `${inventoryContext}\n\n사용자 질문: ${message}` }]
    });

    const result = await model.generateContent({ contents });
    const responseText = result.response.text();

    res.json({
      success: true,
      message: responseText,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('AI Chat Error:', error);
    res.status(500).json({ success: false, message: 'AI 응답 중 오류가 발생했습니다.' });
  }
});

// ============================================================
// 서버 시작
// ============================================================
app.listen(PORT, () => {
  console.log(`\n🚀 백엔드 서버 실행 중: http://localhost:${PORT}`);
  console.log(`📋 API 엔드포인트:`);
  console.log(`   GET  /api/inventory          - 전체 재고`);
  console.log(`   GET  /api/inventory/categories - 종류별 그룹`);
  console.log(`   GET  /api/inventory/category/:name - 특정 종류 상세`);
  console.log(`   GET  /api/inventory/summary  - 전체 요약`);
  console.log(`   POST /api/inventory/update   - 재고 수량 업데이트`);
  console.log(`   POST /api/ai/chat            - AI 채팅`);
  console.log(`\n📁 OneDrive 엑셀 파일: ${CONFIG.excelFileName}`);
  console.log(`📊 시트명: ${CONFIG.sheetName}\n`);
});
