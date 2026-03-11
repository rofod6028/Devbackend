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
const LOG_FILE = path.join(__dirname, 'inventory_logs.json');

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
  memoryLogs = logs;
  if (!process.env.RENDER) {
    try {
      fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
    } catch (error) {
      console.error('로그 파일 저장 실패:', error.message);
    }
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
      console.error(`❌ OneDrive 쓰기 실패 (${attempt}/${retries}):`, error.message);

      if ((errorCode === 'notAllowed' || errorCode === 'resourceLocked') && attempt < retries) {
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
    const item = data.find(d => d.id === id);

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

[중요 규칙]
1. 사용자가 입고(추가, 넣기)나 출고(가져가기, 빼기)를 요청하면 반드시 재고를 수정해야 합니다.
2. 수정 시 답변 맨 마지막에 반드시 아래의 형식을 정확하게 포함하세요.
3. 마크다운 코드 블록(\`\`\`json)은 절대 사용하지 말고 반드시 ~~~ 기호만 사용하세요.

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
      try {
        const parts = responseText.split('~~~INVENTORY_UPDATE');
        let jsonText = parts[1].split('~~~')[0].trim();
        // 혹시라도 AI가 섞어 쓴 마크다운 제거
        jsonText = jsonText.replace(/```json|```/g, ''); 
        
        const updateData = JSON.parse(jsonText);
        const { action, items } = updateData;

        for (const item of items) {
          // ✨ 모델명 매칭 강화 (공백 제거 및 소문자 통일)
          const targetItem = inventoryData.find(d => 
            (d.모델명 || '').replace(/\s+/g, '').toLowerCase() === (item.모델명 || '').replace(/\s+/g, '').toLowerCase()
          );

          if (targetItem) {
            if (action === '출고') targetItem.현재수량 = Math.max(0, targetItem.현재수량 - item.수량);
            else if (action === '입고') targetItem.현재수량 += item.수량;
            
            targetItem.최종수정시각 = getKSTDate(); 
            const actualUser = user || 'AI 어시스턴트';
            targetItem.작업자 = actualUser; 

            addLog(action, targetItem, action === '입고' ? item.수량 : -item.수량, actualUser);
          } else {
            console.log(`⚠️ 모델명 매칭 실패: ${item.모델명}`);
          }
        }

        const success = await updateExcelOnOneDrive(inventoryData);
        if (success) {
          inventoryUpdated = true;
          updateResult = { success: true, action, items };
        }
        // 화면에는 기호 제외 텍스트만 노출
        responseText = parts[0].trim();
      } catch (error) {
        console.error('❌ AI 명령어 파싱 실패:', error.message);
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