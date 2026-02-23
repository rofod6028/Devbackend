const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());

// ============================================================
// 환경 설정
// ============================================================
const CONFIG = {
  clientId: '5454a185-bc04-4e74-9597-e2305dd67d36',
  clientSecret: 'Se98Q~SelMSaSB.Euko66Qqcny7wgcpuWy10ZbB0',
  redirectUri: 'http://localhost:5000/callback',
  authCode: 'M.C522_BL2.2.U.4c467e99-93f8-bb64-6b65-9ab0e7b36087', // ⚠️ 만료되면 재발급 필요
  excelFileName: '재고관리.xlsx',
  sheetName: '재고관리'
};

// Refresh Token 저장 파일
const TOKEN_FILE = path.join(__dirname, 'onedrive_tokens.json');
const LOG_FILE = path.join(__dirname, 'inventory_logs.json');

// ============================================================
// Gemini AI 설정
// ============================================================
const genAI = new GoogleGenerativeAI('AIzaSyAkak8ZMrUHwGV01nPw69QCs1qnfwipZiA');
const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

// ============================================================
// Token 관리
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
  // Render 환경에서는 파일 저장 안 함 (서버 재시작하면 사라지기 때문)
  if (process.env.RENDER) {
    console.log('ℹ️ Render 환경: Token 파일 저장 생략');
    return;
  }
  try {
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
    console.log('✅ Token 저장 완료');
  } catch (error) {
    console.error('❌ Token 저장 실패:', error.message);
  }
}

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
    
    // ⚠️ Authorization Code 만료 안내
    if (error.response?.data?.error === 'invalid_grant') {
      console.log('\n⚠️⚠️⚠️ Authorization Code가 만료되었습니다! ⚠️⚠️⚠️');
      console.log('새 코드를 발급받으세요:');
      console.log('1. 브라우저에서 접속:');
      console.log(`   https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${CONFIG.clientId}&response_type=code&redirect_uri=${encodeURIComponent(CONFIG.redirectUri)}&scope=Files.ReadWrite.All`);
      console.log('2. 로그인 후 나오는 URL에서 code= 뒤의 값을 복사');
      console.log('3. proxy-server.js의 CONFIG.authCode 값을 새 코드로 교체');
      console.log('4. 서버 재시작\n');
    }
    
    return null;
  }
}

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

async function getValidAccessToken() {
  let tokens = loadTokens();

  // ✅ Render 환경변수에 REFRESH_TOKEN이 있으면 항상 우선 사용
  if (process.env.REFRESH_TOKEN) {
    if (!tokens || Date.now() >= (tokens.expires_at || 0) - 60000) {
      console.log('🔑 환경변수 Refresh Token으로 갱신 중...');
      const newTokens = await refreshAccessToken(process.env.REFRESH_TOKEN);
      if (newTokens) {
        return newTokens.access_token;
      }
    } else if (tokens) {
      return tokens.access_token;
    }
  }

  // 로컬 환경: 파일에서 토큰 로드
  if (!tokens) {
    console.log('⚠️ 저장된 Token이 없습니다. Device Flow를 시작합니다...');
    tokens = await getTokenViaDeviceFlow();
    if (!tokens) {
      throw new Error('Token 발급 실패. Render라면 REFRESH_TOKEN 환경변수를 설정하세요.');
    }
  }

  // Token 만료 시 갱신
  if (Date.now() >= tokens.expires_at - 60000) {
    tokens = await refreshAccessToken(tokens.refresh_token);
    if (!tokens) {
      throw new Error('Token 갱신 실패.');
    }
  }

  return tokens.access_token;
}
// ============================================================
// 재고 변경 이력 로그 관리
// ============================================================

function loadLogs() {
  try {
    if (fs.existsSync(LOG_FILE)) {
      const data = fs.readFileSync(LOG_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('로그 파일 읽기 실패:', error.message);
  }
  return [];
}

function saveLogs(logs) {
  try {
    fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
  } catch (error) {
    console.error('로그 파일 저장 실패:', error.message);
  }
}

function addLog(action, item, quantityChange, user = 'System') {
  const logs = loadLogs();
  const newLog = {
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    timestampKR: new Date().toLocaleString('ko-KR'),
    action,
    부품종류: item.부품종류,
    모델명: item.모델명,
    적용설비: item.적용설비,
    변경수량: quantityChange,
    변경전수량: item.현재수량 - quantityChange,
    변경후수량: item.현재수량,
    user
  };
  logs.unshift(newLog);
  
  if (logs.length > 1000) {
    logs.splice(1000);
  }
  
  saveLogs(logs);
  console.log(`📝 로그 추가: ${action} - ${item.모델명} (${quantityChange > 0 ? '+' : ''}${quantityChange}) by ${user}`);
}

// ============================================================
// OneDrive 엑셀 파일 읽기/쓰기
// ============================================================

let cachedData = null;
let lastFetchTime = null;
const CACHE_DURATION = 60 * 1000;

function invalidateCache() {
  cachedData = null;
  lastFetchTime = null;
  console.log('🗑️ 캐시 무효화 완료');
}

async function fetchExcelFromOneDrive() {
  const now = Date.now();
  if (cachedData && lastFetchTime && (now - lastFetchTime) < CACHE_DURATION) {
    console.log('📦 캐시된 데이터 사용');
    return cachedData;
  }

  try {
    const accessToken = await getValidAccessToken();

    console.log(`📥 OneDrive에서 "${CONFIG.excelFileName}" 다운로드 중...`);

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        responseType: 'arraybuffer'
      }
    );

    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    const worksheet = workbook.Sheets[CONFIG.sheetName];

    if (!worksheet) {
      console.error(`❌ 시트 "${CONFIG.sheetName}"를 찾을 수 없습니다.`);
      console.log(`사용 가능한 시트: ${workbook.SheetNames.join(', ')}`);
      return getDummyData();
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheet);

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

async function updateExcelOnOneDrive(data, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const accessToken = await getValidAccessToken();

      console.log(`📤 OneDrive에 "${CONFIG.excelFileName}" 업로드 중... (시도 ${attempt}/${retries})`);

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

      console.log('✅ OneDrive 엑셀 업데이트 완료!');
      invalidateCache();
      return true;

    } catch (error) {
      const errorCode = error.response?.data?.error?.code;
      const errorMessage = error.response?.data?.error?.message || error.message;
      
      console.error(`❌ OneDrive 엑셀 쓰기 실패 (시도 ${attempt}/${retries}):`, errorMessage);

      // resourceLocked 에러면 재시도
      if (errorCode === 'notAllowed' || errorCode === 'resourceLocked') {
        if (attempt < retries) {
          const waitTime = attempt * 2000; // 2초, 4초, 6초 대기
          console.log(`⏳ ${waitTime/1000}초 후 재시도합니다...`);
          await new Promise(resolve => setTimeout(resolve, waitTime));
          continue; // 다음 시도
        } else {
          console.error('❌ 파일이 잠겨있습니다. Excel 프로그램을 닫고 다시 시도하세요.');
        }
      }
      
      return false;
    }
  }
  
  return false;
}

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

app.get('/api/inventory/category/:categoryName', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const filtered = data.filter(item => item.부품종류 === req.params.categoryName);
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

app.post('/api/inventory/update', async (req, res) => {
  try {
    const { id, 현재수량 } = req.body;
    
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id === id);
    
    if (!item) {
      return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
    }

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = new Date().toLocaleString('ko-KR');

    const success = await updateExcelOnOneDrive(data);

    if (success) {
      const quantityChange = 현재수량 - oldQuantity;
      addLog('수정', item, quantityChange, 'API');
      
      res.json({ success: true, message: '업데이트 완료', data: item });
    } else {
      res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }

  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/manual-update', async (req, res) => {
  try {
    const { id, 현재수량, action } = req.body;
    
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id === id);
    
    if (!item) {
      return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
    }

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = new Date().toLocaleString('ko-KR');

    const success = await updateExcelOnOneDrive(data);

    if (success) {
      const quantityChange = 현재수량 - oldQuantity;
      addLog(action || '수정', item, quantityChange, 'Manual');
      
      res.json({ success: true, message: '업데이트 완료', data: item });
    } else {
      res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }

  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/logs', (req, res) => {
  try {
    const { limit = 50, filter } = req.query;
    let logs = loadLogs();
    
    if (filter) {
      logs = logs.filter(log => 
        log.모델명.includes(filter) || log.부품종류.includes(filter)
      );
    }
    
    logs = logs.slice(0, parseInt(limit));
    
    res.json({ success: true, data: logs, total: loadLogs().length });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/search', async (req, res) => {
  try {
    const { q } = req.query;
    
    if (!q || q.trim().length < 2) {
      return res.json({ success: true, data: [] });
    }
    
    const data = await fetchExcelFromOneDrive();
    const searchTerm = q.toLowerCase();
    
    const results = data.filter(item => 
      item.모델명.toLowerCase().includes(searchTerm) ||
      item.부품종류.toLowerCase().includes(searchTerm) ||
      item.적용설비.toLowerCase().includes(searchTerm)
    );
    
    res.json({ success: true, data: results });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/alerts', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const alerts = data
      .filter(item => item.현재수량 <= item.최소보유수량)
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

app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message, conversationHistory } = req.body;
    let inventoryData = await fetchExcelFromOneDrive();

    // ✅ 재고 데이터를 명확하게 포맷팅
    const inventoryTable = inventoryData.map(item => 
      `- ${item.부품종류} | ${item.모델명} | 적용설비: ${item.적용설비} | **현재수량: ${item.현재수량}개** | 최소수량: ${item.최소보유수량}개 | 상태: ${item.현재수량 <= item.최소보유수량 ? '⚠️부족' : '✅정상'}`
    ).join('\n');

    const systemPrompt = `
당신은 스페어파츠 재고 관리 AI 어시스턴트입니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 현재 재고 현황 (실시간 엑셀 데이터)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${inventoryTable}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔴 중요: 현재수량 참조 규칙
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. **재고 조회만 할 때**: 표의 현재수량을 그대로 말하세요.
   예: "SKF-6205 베어링은 현재 **12개** 있습니다"

2. **입출고 처리할 때**: 변경 후의 수량을 계산해서 말하세요.
   
   📥 입고 예시:
   - 현재수량: 4개, 입고: 3개
   - ✅ 올바른 답변: "SKF-6205 베어링 3개를 입고 처리했습니다. 현재 재고는 **7개**입니다."
   - ❌ 잘못된 답변: "현재 재고는 **4개**입니다" (변경 전 수량 X)

   📤 출고 예시:
   - 현재수량: 10개, 출고: 3개
   - ✅ 올바른 답변: "SKF-6205 베어링 3개를 출고 처리했습니다. 현재 재고는 **7개**입니다."
   - ❌ 잘못된 답변: "현재 재고는 **10개**입니다" (변경 전 수량 X)

3. **변경 후 수량 계산 방법**:
   - 입고: 현재수량 + 입고수량 = 변경 후 수량
   - 출고: 현재수량 - 출고수량 = 변경 후 수량 (최소 0)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🤖 당신의 역할
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. **재고 조회**: 위 표를 보고 정확한 현재수량을 알려줍니다.
2. **입출고 처리**: 사용자 요청에 따라 재고를 증감시키고, 변경 후 수량을 말합니다.
3. **재고 분석**: 부족/과다 재고를 파악하고 조언합니다.
4. **자연스러운 대화**: 재고 외 일상 대화도 가능하지만, 주 목적은 재고 관리입니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📝 입출고 명령 감지 규칙
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
**출고 키워드**: "사용했어", "썼어", "출고", "빼줘", "소모"
**입고 키워드**: "입고했어", "추가", "넣어줘", "들어왔어", "보충"

⚠️ 필수 정보: 모델명 + 수량
   예: "SKF-6205 3개 사용했어" ✅
   예: "베어링 좀 써" ❌ (수량 없음 → 확인 질문 필요)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ 입출고 명령 응답 형식
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
입출고가 필요한 경우, 응답 **맨 끝**에 이 형식을 추가하세요:

~~~INVENTORY_UPDATE
{
  "action": "출고" 또는 "입고",
  "items": [
    {"모델명": "SKF-6205", "수량": 3}
  ]
}
~~~

**예시 대화:**

사용자: "SKF-6205 베어링 3개 사용했어"
AI 응답: "알겠습니다. SKF-6205 베어링 3개를 출고 처리했습니다.
현재 재고는 12개입니다. (최소 보유수량: 5개)
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "SKF-6205", "수량": 3}]}
~~~"

사용자: "SKF-6205 지금 몇 개야?"
AI 응답: "SKF-6205 베어링은 현재 **12개** 있습니다. 최소 보유수량(5개)보다 많아서 정상입니다."

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ 재고 부족 경고
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
현재수량 ≤ 최소보유수량 → "⚠️ 재고 부족! 발주가 필요합니다" 경고

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
💬 응답 스타일
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- 한국어로 답변
- 친근하고 전문적인 톤
- 명확하지 않은 요청은 확인 질문
- 숫자는 **굵게** 강조
`;

    const contents = [];

    contents.push({
      role: 'user',
      parts: [{ text: systemPrompt }]
    });

    contents.push({
      role: 'model',
      parts: [{ text: '네, 스페어파츠 재고 관리를 도와드리겠습니다. 무엇을 도와드릴까요?' }]
    });

    if (conversationHistory && conversationHistory.length > 0) {
      conversationHistory.forEach(msg => {
        contents.push({
          role: msg.role === 'model' ? 'model' : 'user',
          parts: [{ text: msg.text }]
        });
      });
    }

    contents.push({
      role: 'user',
      parts: [{ text: message }]
    });

    const result = await model.generateContent({ contents });
    let responseText = result.response.text();

    let inventoryUpdated = false;
    let updateResult = null;

    if (responseText.includes('~~~INVENTORY_UPDATE')) {
      try {
        const jsonStart = responseText.indexOf('~~~INVENTORY_UPDATE') + '~~~INVENTORY_UPDATE'.length;
        const jsonEnd = responseText.indexOf('~~~', jsonStart);
        const jsonText = responseText.substring(jsonStart, jsonEnd).trim();
        const updateData = JSON.parse(jsonText);
        
        const { action, items } = updateData;

        for (const item of items) {
          const targetItem = inventoryData.find(d => d.모델명 === item.모델명);
          
          if (targetItem) {
            const oldQuantity = targetItem.현재수량;
            
            if (action === '출고') {
              targetItem.현재수량 = Math.max(0, targetItem.현재수량 - item.수량);
            } else if (action === '입고') {
              targetItem.현재수량 += item.수량;
            }
            targetItem.최종수정시각 = new Date().toLocaleString('ko-KR');
            
            const quantityChange = action === '입고' ? item.수량 : -item.수량;
            addLog(action, targetItem, quantityChange, 'AI');
          }
        }

        const success = await updateExcelOnOneDrive(inventoryData);
        
        if (success) {
          inventoryUpdated = true;
          updateResult = { success: true, action, items };
        }

        responseText = responseText.split('~~~INVENTORY_UPDATE')[0].trim();

      } catch (error) {
        console.error('재고 업데이트 처리 실패:', error);
      }
    }

    res.json({
      success: true,
      message: responseText,
      inventoryUpdated,
      updateResult,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('AI Chat Error:', error);
    res.status(500).json({ 
      success: false, 
      message: 'AI 응답 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.',
      error: error.message 
    });
  }
});

app.listen(PORT, () => {
  console.log(`\n🚀 백엔드 서버 실행 중: http://localhost:${PORT}`);
  console.log(`📋 API 엔드포인트:`);
  console.log(`   GET  /api/inventory                - 전체 재고`);
  console.log(`   GET  /api/inventory/categories     - 종류별 그룹`);
  console.log(`   GET  /api/inventory/category/:name - 특정 종류 상세`);
  console.log(`   GET  /api/inventory/summary        - 전체 요약`);
  console.log(`   POST /api/inventory/update         - 재고 수량 업데이트`);
  console.log(`   POST /api/inventory/manual-update  - 수동 재고 수정 ✨`);
  console.log(`   GET  /api/inventory/logs           - 재고 변경 이력 ✨`);
  console.log(`   GET  /api/inventory/search         - 검색 ✨`);
  console.log(`   GET  /api/inventory/alerts         - 재고 부족 알림 ✨`);
  console.log(`   POST /api/ai/chat                  - AI 채팅`);
  console.log(`\n📁 OneDrive 엑셀 파일: ${CONFIG.excelFileName}`);
  console.log(`📊 시트명: ${CONFIG.sheetName}\n`);
});

// ============================================================
// Device Code Flow로 Token 발급
// ============================================================
async function getTokenViaDeviceFlow() {
  try {
    console.log('\n📱 Device Code Flow 시작...\n');

    const deviceCodeResponse = await axios.post(
      'https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        scope: 'Files.ReadWrite Files.ReadWrite.All offline_access'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      }
    );

    const { user_code, device_code, verification_uri, expires_in, interval } = deviceCodeResponse.data;

    console.log('════════════════════════════════════════════════════');
    console.log('🔐 아래 단계를 따라하세요:');
    console.log('════════════════════════════════════════════════════');
    console.log(`\n1. 휴대폰이나 다른 기기에서 이 URL 접속:`);
    console.log(`   👉 ${verification_uri}`);
    console.log(`\n2. 화면에 이 코드를 입력하세요:`);
    console.log(`   👉 ${user_code}`);
    console.log(`\n3. 개인 Microsoft 계정으로 로그인하세요`);
    console.log(`   (${CONFIG.clientId.substring(0, 8)}...로 시작하는 앱)`);
    console.log(`\n⏰ ${expires_in}초 안에 완료해야 합니다.\n`);
    console.log('대기 중');

    const pollingInterval = (interval || 5) * 1000;
    const maxAttempts = Math.floor(expires_in / (interval || 5));

    for (let i = 0; i < maxAttempts; i++) {
      await new Promise(resolve => setTimeout(resolve, pollingInterval));

      try {
        const tokenResponse = await axios.post(
          'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
          new URLSearchParams({
            client_id: CONFIG.clientId,
            grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
            device_code: device_code
          }),
          {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
          }
        );

        const tokens = {
          access_token: tokenResponse.data.access_token,
          refresh_token: tokenResponse.data.refresh_token,
          expires_at: Date.now() + (tokenResponse.data.expires_in * 1000)
        };

        saveTokens(tokens);
        console.log('\n✅ 인증 성공! Token 발급 완료!\n');
        console.log('📄 Refresh Token이 onedrive_tokens.json에 저장되었습니다.\n');
        return tokens;

      } catch (error) {
        if (error.response?.data?.error === 'authorization_pending') {
          process.stdout.write('.');
        } else if (error.response?.data?.error === 'authorization_declined') {
          console.log('\n❌ 사용자가 권한을 거부했습니다.');
          return null;
        } else {
          throw error;
        }
      }
    }

    console.log('\n❌ 시간 초과: 인증을 완료하지 못했습니다.');
    return null;

  } catch (error) {
    console.error('\n❌ Device Code Flow 실패:', error.response?.data || error.message);
    return null;
  }
}