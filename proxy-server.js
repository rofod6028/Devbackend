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
// 환경 설정 (환경변수 우선, 없으면 직접 입력값 사용)
// ============================================================
const CONFIG = {
  clientId: process.env.CLIENT_ID || '5454a185-bc04-4e74-9597-e2305dd67d36',
  clientSecret: process.env.CLIENT_SECRET || 'Se98Q~SelMSaSB.Euko66Qqcny7wgcpuWy10ZbB0',
  redirectUri: process.env.REDIRECT_URI || 'http://localhost:5000/callback',
  excelFileName: process.env.EXCEL_FILE_NAME || '재고관리.xlsx',
  sheetName: process.env.SHEET_NAME || '재고관리'
};

const TOKEN_FILE = path.join(__dirname, 'onedrive_tokens.json');
// 현재 실행 중인 폴더의 'inventory_logs.json'을 확실히 지칭
const LOG_FILE = path.resolve(__dirname, 'inventory_logs.json');
let memoryLogs = []; // 서버 메모리에 최신 로그를 들고 있게 합니다.

// ============================================================
// Gemini AI 설정
// ============================================================
const genAI = new GoogleGenerativeAI(
  process.env.GEMINI_API_KEY || 'AIzaSyAkak8ZMrUHwGV01nPw69QCs1qnfwipZiA'
);
const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

// ============================================================
// Token 관리 - 메모리 캐시 추가
// ============================================================
let memoryTokens = null; // Render용 메모리 저장

function loadTokens() {
  // 메모리에 있으면 메모리 우선
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
  // 항상 메모리에 저장
  memoryTokens = tokens;
  
  // Render 환경이 아닐 때만 파일에도 저장
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
  // 1. 환경변수 REFRESH_TOKEN이 있다면 최우선으로 사용
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

  // 2. 저장된 토큰 파일 로드
  let tokens = loadTokens();

  // 3. 토큰이 아예 없으면 Device Flow 실행
  if (!tokens) {
    console.log('⚠️ 저장된 토큰이 없습니다. 기기 인증(Device Flow)을 시작합니다.');
    tokens = await getTokenViaDeviceFlow(); // 👈 여기서 사용자님이 만드신 함수를 호출합니다!
    if (!tokens) throw new Error('인증에 실패했습니다.');
    return tokens.access_token;
  }

  // 4. 토큰 만료 시 갱신 시도
  if (Date.now() >= tokens.expires_at - 60000) {
    console.log('🔄 토큰 만료됨. 갱신 중...');
    const refreshed = await refreshAccessToken(tokens.refresh_token);
    
    // 갱신 실패(Refresh Token 만료) 시 다시 Device Flow 실행
    if (!refreshed) {
      console.log('⚠️ 갱신 실패. 다시 기기 인증(Device Flow)을 진행합니다.');
      tokens = await getTokenViaDeviceFlow();
      if (!tokens) throw new Error('재인증 실패');
      return tokens.access_token;
    }
    return refreshed.access_token;
  }

  return tokens.access_token;
}

function saveLogs(logs) {
  memoryLogs = logs; // 메모리에 즉시 반영 (프론트엔드에서 바로 보이게 함)
  
  // Render 환경이 아닐 때만 파일로 저장 (Render는 재배포 시 파일이 날아가므로 메모리가 더 중요함)
  try {
    fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
  } catch (error) {
    console.error('❌ 로그 저장 실패:', error.message);
  }
}

function addLog(action, item, quantityChange, user = 'System') {
  const logs = loadLogs();
  const newLog = {
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    timestampKR: getKSTDate(), // 한국 시간 함수 사용
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
  if (logs.length > 1000) logs.splice(1000);
  saveLogs(logs);
  console.log(`📝 로그: ${action} - ${item.모델명} (${quantityChange > 0 ? '+' : ''}${quantityChange})`);
}
// [추가할 코드 1] 로그 파일을 읽어오는 함수
function loadLogs() {
  try {
    // 파일이 없으면 에러를 내지 말고 빈 배열([])을 반환하게 함
    if (fs.existsSync(LOG_FILE)) {
      const data = fs.readFileSync(LOG_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('❌ 로그 읽기 실패:', error.message);
  }
  return []; // 파일이 없거나 오류 시 빈 목록 반환
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
        headers: { 'Authorization': `Bearer ${accessToken}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    const worksheet = workbook.Sheets[CONFIG.sheetName];

    if (!worksheet) {
      console.error(`❌ 시트 "${CONFIG.sheetName}" 없음`);
      return getDummyData();
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheet);
        const mappedData = jsonData.map((row, index) => {
      // 엑셀 열 이름 중 '보관장소'를 찾습니다.
      const rowKeys = Object.keys(row);
      const foundKey = rowKeys.find(key => key.trim() === '보관장소');

      return {
        id: index + 1,
        대분류: row['대분류'] || '미분류',
        부품종류: row['부품종류'] || '',
        모델명: row['모델명'] || '',
        적용설비: row['적용설비'] || '',
        현재수량: Number(row['현재수량']) || 0,
        최소보유수량: Number(row['최소보유수량']) || 0,
        최종수정시각: row['최종수정시각'] || '',
        작업자: row['작업자'] || '',
        용도: row['용도'] || '',
        // ✨ storageKey 대신 찾은 키(foundKey)를 사용하여 안전하게 읽어옵니다.
        보관장소: foundKey ? row[foundKey] : '위치 미지정'
      };
    });

    cachedData = mappedData;
    lastFetchTime = now;
    console.log(`✅ OneDrive 데이터 로드 완료: ${mappedData.length}건`);
    return mappedData;

  } catch (error) {
    console.error('❌ OneDrive 읽기 실패:', error.response?.data || error.message);
    return getDummyData();
  }
}

async function updateExcelOnOneDrive(data, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const accessToken = await getValidAccessToken();

      const worksheet = XLSX.utils.json_to_sheet(data.map(item => ({
      '대분류': item.대분류 || '미분류',
      '부품종류': item.부품종류 || '',
      '모델명': item.모델명 || '',
      '적용설비': item.적용설비 || '',
      '현재수량': Number(item.현재수량) || 0,
      '최소보유수량': Number(item.최소보유수량) || 0,
      '최종수정시각': item.최종수정시각 || '',
      '작업자': item.작업자 || '',
      '용도': item.용도 || '',
      '보관장소': item.보관장소 || '위치 미지정' // ✨ 엑셀 헤더와 정확히 일치해야 함
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

      console.log('✅ OneDrive 업데이트 완료!');
      invalidateCache();
      return true;

    } catch (error) {
      const errorCode = error.response?.data?.error?.code;
      const errorDetail = error.response?.data?.error?.message;
      
      // 상세 로그 출력 (이게 핵심입니다!)
      console.error(`❌ OneDrive 쓰기 실패 (${attempt}/${retries})`);
      console.error(`❌ 에러 코드: ${errorCode}`);
      console.error(`❌ 상세 메시지: ${errorDetail || error.message}`);

      if ((errorCode === 'notAllowed' || errorCode === 'resourceLocked') && attempt < retries) {
        console.log("🔄 파일이 잠겨있거나 권한 문제로 재시도합니다...");
        await new Promise(resolve => setTimeout(resolve, attempt * 2000));
        continue;
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
    { id: 3, 부품종류: '오일필터', 모델명: 'MANN-W940', 적용설비: '컴프레서1', 현재수량: 20, 최소보유수량: 8, 최종수정시각: '2026-02-03 08:00' },
    { id: 4, 부품종류: '벨트', 모델명: 'Gates-A68', 적용설비: '모터A', 현재수량: 6, 최소보유수량: 3, 최종수정시각: '2026-02-01 09:30' },
    { id: 5, 부품종류: '패킹', 모델명: 'Teikoku-S1', 적용설비: '펌프A', 현재수량: 30, 최소보유수량: 10, 최종수정시각: '2026-02-03 07:45' },
  ];
}
const getKSTDate = () => {
  const curr = new Date();
  // 한국 시간(UTC+9) 계산
  const utc = curr.getTime() + (curr.getTimezoneOffset() * 60 * 1000);
  const KR_TIME_DIFF = 9 * 60 * 60 * 1000;
  const kstDate = new Date(utc + KR_TIME_DIFF);
  
  // '2024. 3. 4. 오후 12:30:45' 형식으로 반환
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
      const mainCat = item.대분류 || '미분류'; // ✨ 대분류를 기준으로 사용
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
    // ✨ 부품종류가 아니라 대분류(categoryName)로 필터링합니다.
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
  // ✨ 최소보유수량이 0보다 큰 항목들 중에서만 재고 부족을 찾습니다.
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
    const { id, 현재수량 } = req.body;
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id === id);
    if (!item) return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate(); // 한국 시간 함수 사용

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      // 로그를 남길 때 item 객체가 살아있는지 확인하며 안전하게 기록
      try {
        addLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual');
      } catch (logErr) {
        console.error('로그 기록 중 오류(무시됨):', logErr.message);
      }
      
      // 프론트엔드에 성공 응답을 명확히 보냄
      return res.status(200).json({ 
        success: true, 
        message: '업데이트 완료', 
        data: item 
      });
    } else {
      return res.status(500).json({ 
        success: false, 
        message: 'OneDrive 업데이트 실패' 
      });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/manual-update', async (req, res) => {
  try {
    const { id, 현재수량, action, user } = req.body;
const data = await fetchExcelFromOneDrive();

// 💡 숫자/문자열 상관없이 비교하도록 == 사용 및 로그 추가
const item = data.find(d => d.id == id); 

if (!item) {
  console.error(`❌ 항목 찾기 실패: 요청된 ID=${id}, 데이터 첫항목 ID=${data[0]?.id}`);
  return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
}

    if (!item) {
      return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
    }

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate(); // 한국 시간 함수 사용
    item.작업자 = user || 'Manual';

    const success = await updateExcelOnOneDrive(data);
    
    if (success) {
      // 로그 기록 중 에러가 나더라도 응답은 성공으로 보내도록 try-catch로 감싸기
      try {
        addLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual');
      } catch (logErr) {
        console.error('📝 로그 기록 오류(무시됨):', logErr.message);
      }

      return res.status(200).json({ 
        success: true, 
        message: '업데이트 완료', 
        data: item 
      });
    } else {
      return res.status(500).json({ 
        success: false, 
        message: 'OneDrive 업데이트 실패' 
      });
    }
  } catch (error) {
    console.error('❌ manual-update 서버 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
}); // 👈 여기서 함수의 중괄호가 닫혀야 합니다.

app.get('/api/inventory/alerts', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const alerts = data
      .filter(item => item.최소보유수량 > 0 && item.현재수량 <= item.최소보유수량) // ✨ 조건 추가
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
// [추가할 코드 2] 프론트엔드에 로그 데이터를 보내주는 통로
app.get('/api/inventory/logs', (req, res) => {
  try {
    const logs = loadLogs(); 
    const limit = parseInt(req.query.limit) || 100;
    res.json({ success: true, data: logs.slice(0, limit) });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});
// ✨ 검색 기능 API 추가 (404 에러 해결용)
app.get('/api/inventory/search', async (req, res) => {
  try {
    const query = req.query.q ? req.query.q.toLowerCase() : '';
    // 1. 최신 엑셀 데이터를 가져옵니다.
    const data = await fetchExcelFromOneDrive();
    
    // 2. 모델명, 부품종류, 적용설비 중 검색어가 포함된 항목만 필터링합니다.
    const filtered = data.filter(item => 
      (item.모델명 && item.모델명.toLowerCase().includes(query)) ||
      (item.부품종류 && item.부품종류.toLowerCase().includes(query)) ||
      (item.적용설비 && item.적용설비.toLowerCase().includes(query)) ||
      (item.대분류 && item.대분류.toLowerCase().includes(query))
    );

    console.log(`🔍 검색 수행: "${query}" -> ${filtered.length}건 발견`);
    res.json({ success: true, data: filtered });
  } catch (error) {
    console.error('❌ 검색 API 에러:', error.message);
    res.status(500).json({ success: false, message: '검색 중 오류가 발생했습니다.' });
  }
});
app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message, conversationHistory, user } = req.body;
    
    // 1. 최신 데이터 로드 (캐시 무시)
    invalidateCache();
    let inventoryData = await fetchExcelFromOneDrive();

    // 2. AI에게 전달할 재고 현황 테이블 생성 (보관장소 포함)
    const inventoryTable = inventoryData.map(item =>
      `- [${item.대분류}] ${item.모델명} | 현재: ${item.현재수량}개 | 위치: ${item.보관장소} | 용도: ${item.용도 || '정보 없음'}`
    ).join('\n');

    // 3. AI 지시사항 (프롬프트) 강화
    const systemPrompt = `당신은 스마트 재고 관리 전문가입니다. 반드시 아래 [최신 재고 현황]을 근거로 답변하세요.

[최신 재고 현황]
${inventoryTable}

[중요 지시]
1. 이전 대화 내용에 적힌 수량은 무시하십시오. 오직 아래 제공된 [최신 실시간 재고 현황]의 수량만 정답으로 간주합니다.
2. 사용자가 수동으로 수치를 변경했으므로, 당신의 이전 답변과 현재 수량이 달라도 [최신 현황]을 기준으로 답변하십시오.
3. 입출고 처리 후에는 반드시 원본 데이터의 수치를 당신이 직접 계산하지 말고, 수정 명령(INVENTORY_UPDATE)만 내리십시오.
4. 수정 시 답변 맨 마지막에 반드시 아래의 형식을 정확하게 포함하세요.
5. 마크다운 코드 블록(\`\`\`json)은 절대 사용하지 말고 반드시 ~~~ 기호만 사용하세요.

[응답 형식 예시]
친절한 설명 후 마지막에 아래 내용 추가:
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1}]}
~~~`;

    const contents = [
      { role: 'user', parts: [{ text: systemPrompt }] },
      { role: 'model', parts: [{ text: '네, 실시간 재고 현황을 바탕으로 입출고 관리를 도와드리겠습니다!' }] }
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

    // 4. AI 답변에서 명령어 추출 및 실행
if (responseText.includes('~~~INVENTORY_UPDATE')) {
  console.log("🤖 AI가 재고 수정 명령을 보냈습니다. 분석을 시작합니다...");
  try {
    const parts = responseText.split('~~~INVENTORY_UPDATE');
    if (!parts[1]) throw new Error("AI 응답 형식이 올바르지 않습니다. (구분자 부족)");

    let jsonPart = parts[1].split('~~~')[0].trim();
    jsonPart = jsonPart.replace(/```json|```/g, ''); // 마크다운 제거
    
    console.log("📝 파싱할 JSON 데이터:", jsonPart);
    const updateData = JSON.parse(jsonPart);
    const { action, items } = updateData;

    for (const item of items) {
  const targetItem = inventoryData.find(d => 
    String(d.모델명 || '').replace(/\s+/g, '').toLowerCase() === 
    String(item.모델명 || '').replace(/\s+/g, '').toLowerCase()
  );

  if (targetItem) {
    const changeQty = Number(item.수량) || 0;
    const oldQty = targetItem.현재수량; // 변경 전 수량 기억

    if (action === '출고') targetItem.현재수량 = Math.max(0, targetItem.현재수량 - changeQty);
    else if (action === '입고') targetItem.현재수량 += changeQty;
    
    targetItem.최종수정시각 = getKSTDate();
    targetItem.작업자 = user || 'AI 어시스턴트';

    // ✨ 핵심: 여기에 addLog를 추가해야 AI 변경 이력이 남습니다!
    // 출고는 마이너스(-), 입고는 플러스(+)로 수량 변화를 기록합니다.
    const finalChange = action === '출고' ? -changeQty : changeQty;
    
    try {
      addLog(action, targetItem, finalChange, user || 'AI 어시스턴트');
      console.log(`📝 AI 로그 기록 성공: ${targetItem.모델명}`);
    } catch (logErr) {
      console.error('❌ AI 로그 기록 실패:', logErr.message);
    }
  }
}

// 그 후 엑셀 업데이트 실행
const success = await updateExcelOnOneDrive(inventoryData);
    console.log("💾 OneDrive 저장 시도 결과:", success ? "성공" : "실패");
    
    if (success) {
      inventoryUpdated = true;
      updateResult = { success: true, action, items };
    }
  } catch (error) {
    console.error('❌ AI 반영 프로세스 중 오류 발생:', error.message);
    console.error('❌ 에러 상세 내용:', error.stack);
  }
} else {
  console.log("💡 AI 답변에 수정 명령이 포함되지 않았습니다.");
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