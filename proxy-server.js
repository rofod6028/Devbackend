const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

// ============================================================
// 환경 설정
// ============================================================
const CONFIG = {
  excelFileName: process.env.EXCEL_FILE_NAME || '재고관리(개발중).xlsx',
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  redirectUri: process.env.REDIRECT_URI || 'http://localhost:5000/callback',
  inventorySheets: ['충전', '타정', '제조', '공통'],
  logSheetName: '사용내역종합'
};

// 환경변수 로딩 상태 로깅
console.log('📋 환경변수 설정 상태:');
console.log(`   Excel File: ${CONFIG.excelFileName ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Client ID: ${CONFIG.clientId ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Gemini Key: ${process.env.GEMINI_API_KEY ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Refresh Token: ${process.env.REFRESH_TOKEN ? '✅ 설정됨' : '❌ 미설정'}`);

const TOKEN_FILE = path.join(__dirname, 'onedrive_tokens.json');
const LOG_FILE = path.resolve(__dirname, 'inventory_logs.json');

let memoryLogs = [];
let memoryTokens = null;

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

// ============================================================
// Token 관리
// ============================================================
function loadTokens() {
  if (memoryTokens) return memoryTokens;
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
  memoryTokens = tokens;
  if (!process.env.RENDER) {
    try {
      fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
      console.log('✅ Token 파일 저장 완료');
    } catch (error) {
      console.error('❌ Token 파일 저장 실패:', error.message);
    }
  }
}

async function refreshAccessToken(refreshToken) {
  try {
    console.log('🔄 Access Token 갱신 중...');
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: CONFIG.clientId,
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
    console.log('✅ Access Token 갱신 성공!');
    return tokens;
  } catch (error) {
    console.error('❌ Token 갱신 실패:', error.response?.data || error.message);
    return null;
  }
}

async function getValidAccessToken() {
  // 1. 환경변수 REFRESH_TOKEN 최우선 사용
  if (process.env.REFRESH_TOKEN) {
    try {
      console.log('🔑 환경변수 REFRESH_TOKEN으로 갱신 중...');
      const response = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        new URLSearchParams({
          client_id: CONFIG.clientId,
          refresh_token: process.env.REFRESH_TOKEN,
          grant_type: 'refresh_token'
        }),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );
      return response.data.access_token;
    } catch (err) {
      console.error('❌ 환경변수 토큰 갱신 실패. 로컬 인증으로 전환합니다.');
    }
  }

  // 2. 저장된 토큰 로드
  let tokens = loadTokens();

  // 3. 토큰 없으면 Device Flow
  if (!tokens) {
    console.log('⚠️ 저장된 토큰이 없습니다. Device Flow를 시작합니다.');
    tokens = await getTokenViaDeviceFlow();
    if (!tokens) throw new Error('인증에 실패했습니다.');
    return tokens.access_token;
  }

  // 4. 만료 시 갱신
  if (Date.now() >= tokens.expires_at - 60000) {
    console.log('🔄 토큰 만료됨. 갱신 중...');
    const refreshed = await refreshAccessToken(tokens.refresh_token);
    if (!refreshed) {
      tokens = await getTokenViaDeviceFlow();
      if (!tokens) throw new Error('재인증 실패');
      return tokens.access_token;
    }
    return refreshed.access_token;
  }

  return tokens.access_token;
}

// ============================================================
// 로그 관리
// ============================================================
function loadLogs() {
  try {
    if (fs.existsSync(LOG_FILE)) {
      const data = fs.readFileSync(LOG_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('❌ 로그 읽기 실패:', error.message);
  }
  return [];
}

function saveLogs(logs) {
  memoryLogs = logs;
  try {
    fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
  } catch (error) {
    console.error('❌ 로그 저장 실패:', error.message);
  }
}

function addLog(action, item, quantityChange, user = 'System') {
  const newLog = {
    id: uuidv4(),
    timestampKR: getKSTDate(),
    action,
    원본시트: item.원본시트 || '미분류',
    부품종류: item.부품종류,
    모델명: item.모델명,
    적용설비: item.적용설비,
    변경수량: quantityChange,
    변경전수량: item.현재수량 - quantityChange,
    변경후수량: item.현재수량,
    user
  };
  const logs = loadLogs();
  logs.unshift(newLog);
  if (logs.length > 1000) logs.splice(1000);
  saveLogs(logs);
  console.log(`📝 로그: ${action} - ${item.모델명} (${quantityChange > 0 ? '+' : ''}${quantityChange})`);
}

// ============================================================
// OneDrive 엑셀 읽기 (Graph API + OAuth 토큰 방식)
// ============================================================
let cachedData = null;
let lastFetchTime = null;
const CACHE_DURATION = 60 * 1000;

function invalidateCache() {
  cachedData = null;
  lastFetchTime = null;
}

async function fetchExcelFromOneDrive() {
  const now = Date.now();
  if (cachedData && lastFetchTime && (now - lastFetchTime) < CACHE_DURATION) {
    console.log('📦 캐시된 통합 데이터 사용');
    return cachedData;
  }

  try {
    const accessToken = await getValidAccessToken();
    console.log(`📥 OneDrive에서 "${CONFIG.excelFileName}" 다운로드 중...`);

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(CONFIG.excelFileName)}:/content`,
      {
        headers: { 'Authorization': `Bearer ${accessToken}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    let allMappedData = [];

    // 재고 시트 순회 (충전, 타정, 제조, 공통)
    CONFIG.inventorySheets.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) {
        console.warn(`⚠️ 시트 "${sheetName}"을 찾을 수 없습니다`);
        return;
      }

      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      console.log(`✅ "${sheetName}" 시트: ${jsonData.length}개 항목`);

      const mappedData = jsonData.map((row, index) => {
        const rowKeys = Object.keys(row);
        const foundKey = rowKeys.find(key => key.trim() === '보관장소');

        return {
          id: `${sheetName}_${index + 1}`,
          원본시트: sheetName,
          대분류: row['대분류'] || '미분류',
          부품종류: row['부품종류'] || '',
          모델명: row['모델명'] || '',
          적용설비: row['적용설비'] || '',
          현재수량: Number(row['현재수량']) || 0,
          최소보유수량: Number(row['최소보유수량']) || 0,
          최종수정시각: row['최종수정시각'] || '',
          작업자: row['작업자'] || '',
          용도: row['용도'] || '',
          보관장소: foundKey ? row[foundKey] : '위치 미지정'
        };
      });
      allMappedData = [...allMappedData, ...mappedData];
    });

    // 로그 시트 로드 (사용내역종합)
    const logWorksheet = workbook.Sheets[CONFIG.logSheetName];
    if (logWorksheet) {
      const logJson = XLSX.utils.sheet_to_json(logWorksheet);
      memoryLogs = logJson.reverse().slice(0, 1000);
      console.log(`📜 로그 시트 로드 완료: ${memoryLogs.length}건`);
    }

    cachedData = allMappedData;
    lastFetchTime = now;
    console.log(`✅ 데이터 로드 완료: 총 ${allMappedData.length}건`);
    return allMappedData;

  } catch (error) {
    console.error('❌ OneDrive 읽기 실패:', error.response?.data || error.message);
    return [];
  }
}

async function updateExcelOnOneDrive(data, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const accessToken = await getValidAccessToken();
      const workbook = XLSX.utils.book_new();

      // 재고 시트들 저장
      CONFIG.inventorySheets.forEach(sheetName => {
        const sheetData = data.filter(item => item.원본시트 === sheetName);
        const excelRows = sheetData.map(item => ({
          '대분류': item.대분류 || '미분류',
          '부품종류': item.부품종류 || '',
          '모델명': item.모델명 || '',
          '적용설비': item.적용설비 || '',
          '현재수량': Number(item.현재수량) || 0,
          '최소보유수량': Number(item.최소보유수량) || 0,
          '최종수정시각': item.최종수정시각 || '',
          '작업자': item.작업자 || '',
          '용도': item.용도 || '',
          '보관장소': item.보관장소 || '위치 미지정'
        }));
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(excelRows), sheetName);
      });

      // 로그 시트 저장
      const logRows = [...memoryLogs].reverse();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(logRows), CONFIG.logSheetName);

      const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

      await axios.put(
        `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(CONFIG.excelFileName)}:/content`,
        excelBuffer,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }
        }
      );

      console.log(`✅ OneDrive 업데이트 완료! (${CONFIG.excelFileName})`);
      invalidateCache();
      return true;

    } catch (error) {
      console.error(`❌ OneDrive 쓰기 실패 (${attempt}/${retries}): ${error.message}`);
      if (attempt < retries) {
        await new Promise(resolve => setTimeout(resolve, attempt * 2000));
        continue;
      }
      return false;
    }
  }
  return false;
}

const getKSTDate = () => {
  const curr = new Date();
  const utc = curr.getTime() + (curr.getTimezoneOffset() * 60 * 1000);
  const KR_TIME_DIFF = 9 * 60 * 60 * 1000;
  const kstDate = new Date(utc + KR_TIME_DIFF);
  return kstDate.toLocaleString('ko-KR');
};

// ============================================================
// API Routes
// ============================================================
app.get('/api/inventory', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    res.json({ success: true, data });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/categories', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const categories = {};
    data.forEach(item => {
      const mainCat = item.대분류 || '미분류';
      if (!categories[mainCat]) {
        categories[mainCat] = { name: mainCat, totalCount: 0, itemCount: 0, lowStockCount: 0, items: [] };
      }
      categories[mainCat].items.push(item);
      categories[mainCat].totalCount += item.현재수량;
      categories[mainCat].itemCount += 1;
      if (item.최소보유수량 > 0 && item.현재수량 <= item.최소보유수량) {
        categories[mainCat].lowStockCount += 1;
      }
    });
    res.json({ success: true, data: Object.values(categories) });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/category/:categoryName', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const filtered = data.filter(item => item.대분류 === req.params.categoryName);
    res.json({ success: true, data: filtered });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/summary', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const summary = {
      totalItems: data.length,
      totalQuantity: data.reduce((sum, d) => sum + d.현재수량, 0),
      lowStockItems: data.filter(d => d.최소보유수량 > 0 && d.현재수량 <= d.최소보유수량),
      lowStockCount: data.filter(d => d.최소보유수량 > 0 && d.현재수량 <= d.최소보유수량).length,
      categoryBreakdown: {}
    };
    data.forEach(item => {
      if (!summary.categoryBreakdown[item.부품종류]) {
        summary.categoryBreakdown[item.부품종류] = { total: 0, count: 0, lowStock: 0 };
      }
      summary.categoryBreakdown[item.부품종류].total += item.현재수량;
      summary.categoryBreakdown[item.부품종류].count += 1;
      if (item.현재수량 <= item.최소보유수량) summary.categoryBreakdown[item.부품종류].lowStock += 1;
    });
    res.json({ success: true, data: summary });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/update', async (req, res) => {
  try {
    const { id, 현재수량, action, user } = req.body;
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id == id);
    if (!item) return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      try {
        addLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual');
      } catch (logErr) {
        console.error('로그 기록 중 오류(무시됨):', logErr.message);
      }
      return res.status(200).json({ success: true, message: '업데이트 완료', data: item });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/manual-update', async (req, res) => {
  try {
    const { id, 현재수량, action, user } = req.body;
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id == id);

    if (!item) {
      console.error(`❌ 항목 찾기 실패: 요청된 ID=${id}, 데이터 첫항목 ID=${data[0]?.id}`);
      return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
    }

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();
    item.작업자 = user || 'Manual';

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      try {
        addLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual');
      } catch (logErr) {
        console.error('📝 로그 기록 오류(무시됨):', logErr.message);
      }
      return res.status(200).json({ success: true, message: '업데이트 완료', data: item });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    console.error('❌ manual-update 서버 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/alerts', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const alerts = data
      .filter(item => item.최소보유수량 > 0 && item.현재수량 <= item.최소보유수량)
      .map(item => ({
        id: item.id,
        부품종류: item.부품종류,
        모델명: item.모델명,
        적용설비: item.적용설비,
        현재수량: item.현재수량,
        최소보유수량: item.최소보유수량,
        부족수량: item.최소보유수량 - item.현재수량,
        긴급도: item.현재수량 === 0 ? 'critical' : 'warning'
      }))
      .sort((a, b) => {
        if (a.긴급도 === 'critical' && b.긴급도 !== 'critical') return -1;
        if (a.긴급도 !== 'critical' && b.긴급도 === 'critical') return 1;
        return b.부족수량 - a.부족수량;
      });
    res.json({ success: true, data: alerts, count: alerts.length });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/logs', (req, res) => {
  try {
    const logs = loadLogs();
    const limit = parseInt(req.query.limit) || 100;
    res.json({ success: true, data: logs.slice(0, limit) });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/search', async (req, res) => {
  try {
    const query = req.query.q ? String(req.query.q).toLowerCase() : '';
    const data = await fetchExcelFromOneDrive();

    if (!Array.isArray(data)) {
      return res.json({ success: true, data: [] });
    }

    const filtered = data.filter(item => {
      if (!item) return false;
      const model = String(item.모델명 || '').toLowerCase();
      const type = String(item.부품종류 || '').toLowerCase();
      const facility = String(item.적용설비 || '').toLowerCase();
      const mainCat = String(item.대분류 || '').toLowerCase();
      return model.includes(query) || type.includes(query) || facility.includes(query) || mainCat.includes(query);
    });

    res.json({ success: true, data: filtered });
  } catch (error) {
    console.error('❌ 검색 API 내부 에러:', error.stack);
    res.status(500).json({ success: false, message: '서버 내부 오류' });
  }
});

app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message, conversationHistory, user } = req.body;

    invalidateCache();
    let inventoryData = await fetchExcelFromOneDrive();

    const inventoryTable = inventoryData.map(item =>
      `- [${item.원본시트} / ${item.대분류}] ${item.모델명} | 현재: ${item.현재수량}개 | 위치: ${item.보관장소}`
    ).join('\n');

    const systemPrompt = `당신은 스마트 재고 관리 전문가입니다. 반드시 아래 [최신 재고 현황]을 근거로 답변하세요.\n\n[최신 재고 현황]\n${inventoryTable}\n\n[중요 지시]\n1. 답변 시 반드시 해당 부품이 속한 "원본시트(충전, 타정, 공통)" 정보를 확인하십시오.\n2. 입출고 처리 시 모델명뿐만 아니라 반드시 "원본시트" 이름을 JSON 명령에 포함해야 합니다.\n3. 수정 명령(INVENTORY_UPDATE) 형식에 "원본시트" 필드를 반드시 추가하십시오.\n4. 마크다운 코드 블록(\`\`\`json)은 절대 사용하지 말고 반드시 ~~~ 기호만 사용하세요.\n5. 상식을 뛰어넘는 요청(예: 한 번에 50개 이상 변동 등)을 할 경우 사용자에게 두 번 더 확인하십시오.\n\n[응답 형식 예시]\n친절한 설명 후 마지막에 아래 내용 추가:\n~~~INVENTORY_UPDATE\n{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1, "원본시트": "충전"}]}\n~~~`;

    const contents = [
      { role: 'user', parts: [{ text: systemPrompt }] },
      { role: 'model', parts: [{ text: '네, 시트별(충전/타정/공통) 실시간 재고 현황을 바탕으로 정확히 도와드리겠습니다!' }] }
    ];

    if (conversationHistory?.length > 0) {
      conversationHistory.forEach(msg => {
        contents.push({ role: msg.role === 'model' ? 'model' : 'user', parts: [{ text: msg.text }] });
      });
    }
    contents.push({ role: 'user', parts: [{ text: message }] });

    const result = await model.generateContent({ contents });
    let responseText = result.response.text();
    let inventoryUpdated = false;
    let updateResult = null;

    if (responseText.includes('~~~INVENTORY_UPDATE')) {
      try {
        const parts = responseText.split('~~~INVENTORY_UPDATE');
        let jsonPart = parts[1].split('~~~')[0].trim();
        jsonPart = jsonPart.replace(/```json|```/g, '');

        const updateData = JSON.parse(jsonPart);
        const { action, items } = updateData;

        for (const item of items) {
          const targetItem = inventoryData.find(d =>
            String(d.모델명 || '').replace(/\s+/g, '').toLowerCase() === String(item.모델명 || '').replace(/\s+/g, '').toLowerCase() &&
            d.원본시트 === item.원본시트
          );

          if (targetItem) {
            const changeQty = Number(item.수량) || 0;
            const finalChange = action === '출고' ? -changeQty : changeQty;
            targetItem.현재수량 = action === '출고' ? Math.max(0, targetItem.현재수량 - changeQty) : targetItem.현재수량 + changeQty;
            targetItem.최종수정시각 = getKSTDate();
            targetItem.작업자 = user || 'AI 어시스턴트';
            addLog(action, targetItem, finalChange, user || 'AI 어시스턴트');
          }
        }

        const success = await updateExcelOnOneDrive(inventoryData);
        if (success) {
          inventoryUpdated = true;
          updateResult = { success: true, action, items };
        }
      } catch (error) {
        console.error('❌ AI 명령 처리 오류:', error.message);
      }
    }

    res.json({ success: true, message: responseText, inventoryUpdated, updateResult, timestamp: new Date().toISOString() });
  } catch (error) {
    console.error('❌ AI 채팅 에러:', error.message);
    res.status(500).json({ success: false, message: 'AI 응답 오류' });
  }
});

// ============================================================
// Device Code Flow
// ============================================================
async function getTokenViaDeviceFlow() {
  try {
    console.log('\n📱 Device Code Flow 시작...\n');
    const deviceCodeResponse = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        scope: 'Files.ReadWrite Files.ReadWrite.All offline_access'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { user_code, device_code, verification_uri, expires_in, interval } = deviceCodeResponse.data;
    console.log('====================================================');
    console.log(`1. 브라우저에서 접속: ${verification_uri}`);
    console.log(`2. 코드 입력: ${user_code}`);
    console.log(`3. Microsoft 계정으로 로그인`);
    console.log('====================================================\n대기 중');

    const pollInterval = (interval || 5) * 1000;
    const maxAttempts = Math.floor(expires_in / (interval || 5));

    for (let i = 0; i < maxAttempts; i++) {
      await new Promise(resolve => setTimeout(resolve, pollInterval));
      try {
        const tokenResponse = await axios.post(
          'https://login.microsoftonline.com/common/oauth2/v2.0/token',
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
        console.log('\n✅ 인증 성공!');
        return tokens;
      } catch (error) {
        if (error.response?.data?.error === 'authorization_pending') {
          process.stdout.write('.');
        } else {
          throw error;
        }
      }
    }
    return null;
  } catch (error) {
    console.error('❌ Device Flow 실패:', error.response?.data || error.message);
    return null;
  }
}

app.listen(PORT, () => {
  console.log(`\n🚀 백엔드 서버 실행 중: http://localhost:${PORT}`);
  console.log(`📁 OneDrive 파일: ${CONFIG.excelFileName}`);

  if (process.env.REFRESH_TOKEN) {
    console.log('✅ REFRESH_TOKEN 환경변수 감지됨 - OneDrive 연동 준비 완료');
  } else {
    console.log('⚠️ REFRESH_TOKEN 없음 - 로컬에서 get-token.js를 먼저 실행하세요');
  }
});
