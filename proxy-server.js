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
  inventorySheets: ['충전', '타정', '공통'],
  equipmentStandardSheet: '제조',       // 설비명 표준화 테이블 시트
  facilityLogSheetName: '설비이력',     // 설비별 이력 시트
  logSheetName: '사용내역종합',
  teamsWebhookUrl: process.env.TEAMS_WEBHOOK_URL  // Teams Incoming Webhook URL
};

// 환경변수 로딩 상태 로깅
console.log('📋 환경변수 설정 상태:');
console.log(`   Excel File: ${CONFIG.excelFileName ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Client ID: ${CONFIG.clientId ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Gemini Key: ${process.env.GEMINI_API_KEY ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Refresh Token: ${process.env.REFRESH_TOKEN ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Teams Webhook: ${CONFIG.teamsWebhookUrl ? '✅ 설정됨' : '❌ 미설정 (알림 비활성화)'}`);

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

async function refreshAccessToken(refreshToken, clientIdOverride) {
  const clientId = clientIdOverride || CONFIG.clientId;
  try {
    console.log('🔄 Access Token 갱신 중...');
    const params = {
      client_id: clientId,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    };
    // client_secret이 있으면 포함
    if (CONFIG.clientSecret) {
      params.client_secret = CONFIG.clientSecret;
    }
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams(params),
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
      const params = {
        client_id: CONFIG.clientId,
        refresh_token: process.env.REFRESH_TOKEN,
        grant_type: 'refresh_token'
      };
      if (CONFIG.clientSecret) {
        params.client_secret = CONFIG.clientSecret;
      }
      const response = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        new URLSearchParams(params),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );
      const newTokens = {
        access_token: response.data.access_token,
        refresh_token: response.data.refresh_token || process.env.REFRESH_TOKEN,
        expires_at: Date.now() + (response.data.expires_in * 1000)
      };
      saveTokens(newTokens);
      console.log('✅ 환경변수 REFRESH_TOKEN으로 갱신 성공!');
      return newTokens.access_token;
    } catch (err) {
      console.error('❌ 환경변수 토큰 갱신 실패:', err.response?.data || err.message);
      console.log('📁 로컬 저장 토큰으로 전환합니다...');
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
  if (memoryLogs && memoryLogs.length > 0) return memoryLogs;
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

// 설비명 표준화 테이블 캐시
let equipmentStandardMap = null;   // { '원본설비명': '표준설비명' }
let lastEquipmentFetchTime = null;
const EQUIPMENT_CACHE_DURATION = 5 * 60 * 1000; // 5분

// 설비이력 메모리 버퍼
let facilityLogs = [];

function invalidateCache() {
  cachedData = null;
  lastFetchTime = null;
}

// ============================================================
// 설비명 표준화 (제조 시트 기반)
// ============================================================
// 제조 시트 컬럼 구조:
//   원본설비명 | 표준설비명
//
// 공통부품 자동 판별 규칙:
//   설비명 베이스(#숫자 제거)가 같고, 동일 모델명이 2개 이상 설비에 존재하면
//   → 표준설비명이 "(공통)"으로 끝나도록 백엔드에서 override
//
async function loadEquipmentStandards(workbook) {
  const now = Date.now();
  if (equipmentStandardMap && lastEquipmentFetchTime && (now - lastEquipmentFetchTime) < EQUIPMENT_CACHE_DURATION) {
    return equipmentStandardMap;
  }

  const sheet = workbook.Sheets[CONFIG.equipmentStandardSheet];
  equipmentStandardMap = {};

  if (sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet);
    rows.forEach(row => {
      const original = String(row['원본설비명'] || '').trim();
      const standard = String(row['표준설비명'] || '').trim();
      if (original && standard) {
        equipmentStandardMap[original] = standard;
      }
    });
    console.log(`🏭 설비 표준화 테이블 로드: ${Object.keys(equipmentStandardMap).length}개 매핑`);
  } else {
    console.warn(`⚠️ "${CONFIG.equipmentStandardSheet}" 시트 없음 — 설비명 표준화 건너뜀`);
  }

  lastEquipmentFetchTime = now;
  return equipmentStandardMap;
}

// 원본설비명 → 표준설비명 변환
function normalizeEquipment(originalName) {
  // 엑셀 셀 내 줄바꿈(Alt+Enter) 제거 후 공백 정리
  originalName = String(originalName || '').replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!equipmentStandardMap) return originalName;
  return equipmentStandardMap[originalName] || originalName;
}

// 설비명 베이스 추출: "제트프레스 #1 (1공장)" → "제트프레스"
function getEquipmentBase(name) {
  return String(name || '')
    .replace(/#\s*\d+/g, '')     // #1, # 2 등 제거
    .replace(/\d+호기/g, '')     // 1호기, 2호기 제거
    .replace(/\s+/g, ' ')
    .trim();
}

// 전체 데이터에서 공통부품 자동 판별 후 표준설비명 override
function applyCommonEquipment(allData) {
  // 모델명별로 어떤 베이스 설비에 쓰이는지 집계
  const modelToBaseSets = {}; // { '모델명': Set<베이스설비명> }
  allData.forEach(item => {
    const stdName = item.표준설비명 || item.적용설비;
    const base = getEquipmentBase(stdName);
    const model = String(item.모델명 || '').trim();
    if (!model) return;
    if (!modelToBaseSets[model]) modelToBaseSets[model] = new Set();
    modelToBaseSets[model].add(base);
  });

  // 같은 베이스에서 2개 이상 설비에 걸친 모델 → (공통) override
  allData.forEach(item => {
    const stdName = item.표준설비명 || item.적용설비;
    const base = getEquipmentBase(stdName);
    const model = String(item.모델명 || '').trim();
    if (!model) return;

    // 같은 베이스 설비들 중 이 모델이 2개 이상의 서로 다른 표준설비명에 존재하는지 확인
    const sameBaseItems = allData.filter(d => {
      const dBase = getEquipmentBase(d.표준설비명 || d.적용설비);
      return dBase === base && String(d.모델명 || '').trim() === model;
    });
    const uniqueStdNames = new Set(sameBaseItems.map(d => d.표준설비명 || d.적용설비));

    if (uniqueStdNames.size >= 2) {
      // 공통 설비명: 베이스명 + "(공통)"
      item.표준설비명 = base + ' (공통)';
      item.isCommonPart = true;
    }
  });

  return allData;
}

// ============================================================
// 설비이력 관리
// ============================================================
function addFacilityLog(action, item, quantityChange, user) {
  const stdEquipment = item.표준설비명 || item.적용설비;
  const entry = {
    id: uuidv4(),
    timestampKR: getKSTDate(),
    action,
    원본시트: item.원본시트 || '',
    표준설비명: stdEquipment,
    원본설비명: item.적용설비,
    부품종류: item.부품종류,
    모델명: item.모델명,
    변경수량: quantityChange,
    변경전수량: item.현재수량 - quantityChange,
    변경후수량: item.현재수량,
    isCommonPart: item.isCommonPart || false,
    user
  };
  facilityLogs.unshift(entry);
  if (facilityLogs.length > 5000) facilityLogs.splice(5000);
  console.log(`🏭 설비이력: [${stdEquipment}] ${action} - ${item.모델명} (${quantityChange > 0 ? '+' : ''}${quantityChange})`);
}

async function saveFacilityLogsToOneDrive(workbook) {
  // 설비이력 시트에 현재까지의 facilityLogs를 저장 (updateExcelOnOneDrive 내부에서 호출)
  if (facilityLogs.length === 0) return workbook;
  const rows = [...facilityLogs].reverse(); // 오래된 순서로 저장
  const ws = XLSX.utils.json_to_sheet(rows);
  if (workbook.Sheets[CONFIG.facilityLogSheetName]) {
    workbook.Sheets[CONFIG.facilityLogSheetName] = ws;
  } else {
    XLSX.utils.book_append_sheet(workbook, ws, CONFIG.facilityLogSheetName);
  }
  return workbook;
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
      `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
      {
        headers: { 'Authorization': `Bearer ${accessToken}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    let allMappedData = [];

    // 설비명 표준화 테이블 먼저 로드 (제조 시트)
    await loadEquipmentStandards(workbook);

    // 재고 시트 순회 (충전, 타정, 공통)
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
          표준설비명: normalizeEquipment(String(row['적용설비'] || '').replace(/[\r\n]+/g, ' ').trim()),
          현재수량: Number(row['현재수량']) || 0,
          최소보유수량: Number(row['최소보유수량']) || 0,
          최종수정시각: row['최종수정시각'] || '',
          작업자: row['작업자'] || '',
          용도: row['용도'] || '',
          보관장소: foundKey ? row[foundKey] : '위치 미지정',
          isCommonPart: false
        };
      });
      allMappedData = [...allMappedData, ...mappedData];
    });

    // 로그 시트 로드 (사용내역종합) — memoryLogs가 이미 있으면 덮어쓰지 않음
    const logWorksheet = workbook.Sheets[CONFIG.logSheetName];
    if (logWorksheet && memoryLogs.length === 0) {
      const logJson = XLSX.utils.sheet_to_json(logWorksheet);
      // 오래된 순 저장 → 최신순으로 reverse, 상한 없이 전체 보관
      memoryLogs = logJson.reverse();
      console.log(`📜 로그 시트 로드 완료: ${memoryLogs.length}건`);
    }

    // 설비이력 시트 로드 — 엑셀 이력과 메모리 이력을 항상 병합 (휘발 방지)
    const facilityLogWs = workbook.Sheets[CONFIG.facilityLogSheetName];
    if (facilityLogWs) {
      const rows = XLSX.utils.sheet_to_json(facilityLogWs);
      const excelLogs = rows.reverse(); // 최신순
      const memoryIds = new Set(facilityLogs.map(l => l.id));
      const newFromExcel = excelLogs.filter(l => l.id && !memoryIds.has(l.id));
      if (newFromExcel.length > 0) {
        facilityLogs = [...facilityLogs, ...newFromExcel];
        facilityLogs.sort((a, b) => new Date(b.timestampKR || 0) - new Date(a.timestampKR || 0));
        facilityLogs = facilityLogs.slice(0, 5000);
        console.log(`🏭 설비이력 병합 완료: 총 ${facilityLogs.length}건 (엑셀에서 ${newFromExcel.length}건 추가)`);
      } else if (facilityLogs.length === 0) {
        facilityLogs = excelLogs.slice(0, 5000);
        console.log(`🏭 설비이력 초기 로드: ${facilityLogs.length}건`);
      }
    }

    // 공통부품 자동 판별 적용 (공통 탭 항목 제외)
    const nonCommonData = allMappedData.filter(d => d.원본시트 !== '공통');
    const commonTabData = allMappedData.filter(d => d.원본시트 === '공통');
    applyCommonEquipment(nonCommonData);
    allMappedData = [...nonCommonData, ...commonTabData];

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

      // 설비이력 시트 저장
      if (facilityLogs.length > 0) {
        const facilityRows = [...facilityLogs].reverse();
        const facilityWs = XLSX.utils.json_to_sheet(facilityRows);
        XLSX.utils.book_append_sheet(workbook, facilityWs, CONFIG.facilityLogSheetName);
      }

      // 설비명 표준화 시트(제조)는 읽기 전용 — 원본 그대로 복원
      try {
        const origAccessToken2 = await getValidAccessToken();
        const origResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
          { headers: { 'Authorization': `Bearer ${origAccessToken2}` }, responseType: 'arraybuffer' }
        );
        const origWb = XLSX.read(Buffer.from(origResponse.data), { type: 'buffer' });
        if (origWb.Sheets[CONFIG.equipmentStandardSheet]) {
          workbook.Sheets[CONFIG.equipmentStandardSheet] = origWb.Sheets[CONFIG.equipmentStandardSheet];
          if (!workbook.SheetNames.includes(CONFIG.equipmentStandardSheet)) {
            workbook.SheetNames.push(CONFIG.equipmentStandardSheet);
          }
        }
      } catch (e) {
        console.warn('⚠️ 설비 표준화 시트 복원 스킵:', e.message);
      }

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
// Teams 재고 부족 알림
// ============================================================

// 중복 알림 방지 — 동일 항목은 1시간에 1번만 알림
const alertCooldown = new Map();
const ALERT_COOLDOWN_MS = 60 * 60 * 1000;

async function sendTeamsAlert(lowStockItems) {
  if (!CONFIG.teamsWebhookUrl) {
    console.log('⚠️ TEAMS_WEBHOOK_URL 미설정 — 알림 스킵');
    return;
  }

  // 최소보유수량 > 0 이고 쿨다운 지난 항목만 필터
  const now = Date.now();
  const filtered = lowStockItems.filter(item => {
    if (item.최소보유수량 <= 0) return false;
    const lastAlerted = alertCooldown.get(item.id) || 0;
    return now - lastAlerted >= ALERT_COOLDOWN_MS;
  });

  if (filtered.length === 0) {
    console.log('ℹ️ Teams 알림 대상 없음 (쿨다운 또는 조건 미충족)');
    return;
  }

  filtered.forEach(item => alertCooldown.set(item.id, now));

  const critical = filtered.filter(i => i.현재수량 === 0);
  const warning  = filtered.filter(i => i.현재수량 > 0);

  const titleText = critical.length > 0 ? '🚨 긴급 재고 부족 알림' : '⚠️ 재고 부족 알림';
  const summaryParts = [];
  if (critical.length > 0) summaryParts.push(`🔴 재고 0: **${critical.length}건**`);
  if (warning.length  > 0) summaryParts.push(`🟡 부족 경고: **${warning.length}건**`);

  // 긴급(재고 0) 먼저, 그 다음 경고 순으로 정렬
  const sortedFiltered = [...filtered].sort((a, b) => a.현재수량 - b.현재수량);

  // 품목별 카드형 블록 생성 (설비명이 길어도 깔끔하게 표시)
  const itemBlocks = sortedFiltered.map(item => {
    const isCritical = item.현재수량 === 0;
    const badge = isCritical ? '🔴 재고 없음' : '🟡 부족 경고';
    const shortage = item.최소보유수량 - item.현재수량;
    const facilityName = String(item.표준설비명 || item.적용설비 || '-').replace(/[\r\n]+/g, ' ').trim();
    return {
      type: 'Container',
      style: isCritical ? 'attention' : 'warning',
      spacing: 'Small',
      items: [
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'stretch',
              items: [
                {
                  type: 'TextBlock',
                  text: `**${item.모델명 || '-'}**  ${badge}`,
                  wrap: true,
                  size: 'Default',
                  weight: 'Bolder',
                  color: isCritical ? 'Attention' : 'Warning'
                },
                {
                  type: 'TextBlock',
                  text: `📦 ${item.부품종류 || '-'}　|　🏭 ${facilityName}　|　📋 ${item.원본시트 || '-'}시트`,
                  wrap: true,
                  size: 'Small',
                  isSubtle: true,
                  spacing: 'None'
                }
              ]
            },
            {
              type: 'Column',
              width: 'auto',
              items: [
                {
                  type: 'TextBlock',
                  text: `현재 **${item.현재수량}**개`,
                  wrap: false,
                  size: 'Small',
                  color: isCritical ? 'Attention' : 'Warning',
                  weight: 'Bolder',
                  horizontalAlignment: 'Right'
                },
                {
                  type: 'TextBlock',
                  text: `최소 ${item.최소보유수량}개 / **${shortage}개 부족**`,
                  wrap: false,
                  size: 'Small',
                  isSubtle: true,
                  horizontalAlignment: 'Right',
                  spacing: 'None'
                }
              ]
            }
          ]
        }
      ]
    };
  });

  const card = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          {
            type: 'Container',
            style: 'emphasis',
            items: [
              { type: 'TextBlock', text: titleText, weight: 'Bolder', size: 'Large', color: 'Attention', wrap: true },
              { type: 'TextBlock', text: `🕐 ${getKSTDate()}　　${summaryParts.join('　|　')}`, size: 'Small', isSubtle: true, spacing: 'None', wrap: true }
            ]
          },
          { type: 'TextBlock', text: '─────────────────────', size: 'Small', isSubtle: true, spacing: 'Small' },
          ...itemBlocks,
          { type: 'TextBlock', text: '※ 최소보유수량 0 설정 항목은 알림 제외', size: 'Small', isSubtle: true, wrap: true, spacing: 'Medium' }
        ]
      }
    }]
  };

  try {
    await axios.post(CONFIG.teamsWebhookUrl, card);
    console.log(`✅ Teams 알림 전송 완료 — ${filtered.length}건 (긴급 ${critical.length}, 경고 ${warning.length})`);
  } catch (err) {
    console.error('❌ Teams 알림 전송 실패:', err.response?.data || err.message);
  }
}

// 재고 수정 후 저재고 체크 & 알림 트리거 (non-blocking)
function checkAndNotifyLowStock(data) {
  const lowStock = data.filter(d => d.최소보유수량 > 0 && d.현재수량 <= d.최소보유수량);
  if (lowStock.length > 0) {
    console.log(`📊 저재고 감지: ${lowStock.length}건 → Teams 알림 시도`);
    sendTeamsAlert(lowStock).catch(err => console.error('Teams 알림 오류:', err.message));
  }
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
        addFacilityLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual');
      } catch (logErr) {
        console.error('로그 기록 중 오류(무시됨):', logErr.message);
      }
      checkAndNotifyLowStock(data); // Teams 저재고 알림
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
        addFacilityLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual');
      } catch (logErr) {
        console.error('📝 로그 기록 오류(무시됨):', logErr.message);
      }
      checkAndNotifyLowStock(data); // Teams 저재고 알림
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
    let logs = loadLogs();
    const limit    = parseInt(req.query.limit)  || 100;
    const offset   = parseInt(req.query.offset) || 0;
    const facility = req.query.facility ? String(req.query.facility) : null;
    const partType = req.query.partType ? String(req.query.partType) : null;

    if (facility) {
      logs = logs.filter(l =>
        String(l.적용설비 || '').includes(facility) ||
        String(l.표준설비명 || '').includes(facility)
      );
    }
    if (partType) {
      logs = logs.filter(l => String(l.부품종류 || '') === partType);
    }

    const total = logs.length;
    res.json({ success: true, data: logs.slice(offset, offset + limit), total });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// 설비이력 전체 조회 (설비명 필터 가능)
app.get('/api/inventory/facility-logs', (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 200;
    const facility = req.query.facility ? String(req.query.facility) : null;
    const isCommon = req.query.isCommon === 'true'; // 공통탭 여부
    let logs = facilityLogs;
    if (facility) {
      logs = logs.filter(l => {
        // 일반 설비: 표준설비명 또는 원본설비명 매칭
        if (l.표준설비명 === facility || l.원본설비명 === facility) return true;
        // 공통부품 출고 이력: 원본시트가 '공통'이고 실제사용설비(표준설비명에 저장)가 매칭
        if (isCommon && l.isCommonPart && l.표준설비명 === facility) return true;
        // 어떤 설비든 공통부품 이력을 원본시트 기반으로 조회 (공통탭 대시보드용)
        if (isCommon && l.원본시트 === '공통') return true;
        return false;
      });
    }
    res.json({ success: true, data: logs.slice(0, limit), total: logs.length });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// 설비별 출고 요약 (대시보드용)
app.get('/api/inventory/facility-summary', async (req, res) => {
  try {
    // 설비별 총 출고 건수/수량 집계
    const summary = {};
    facilityLogs.forEach(log => {
      const facility = log.표준설비명 || log.원본설비명 || '미분류';
      if (!summary[facility]) {
        summary[facility] = { 표준설비명: facility, 출고건수: 0, 출고수량: 0, 입고건수: 0, 입고수량: 0, 최근이력: null };
      }
      const qty = Math.abs(Number(log.변경수량) || 0);
      if (log.action === '출고' || (log.변경수량 < 0)) {
        summary[facility].출고건수 += 1;
        summary[facility].출고수량 += qty;
      } else {
        summary[facility].입고건수 += 1;
        summary[facility].입고수량 += qty;
      }
      if (!summary[facility].최근이력) {
        summary[facility].최근이력 = log.timestampKR;
      }
    });
    res.json({ success: true, data: Object.values(summary) });
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

// ============================================================
// 방안C: 공통부품 출고 — 실제 사용 설비 포함 처리
// POST /api/inventory/common-update
// body: { id, 현재수량, action, user, 실제사용설비 }
// ============================================================
app.post('/api/inventory/common-update', async (req, res) => {
  try {
    const { id, 현재수량, action, user, 실제사용설비 } = req.body;

    if (!실제사용설비) {
      return res.status(400).json({ success: false, message: '공통부품 출고 시 실제사용설비는 필수입니다.' });
    }

    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id == id);
    if (!item) return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();
    item.작업자 = user || 'Manual';

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      // 로그에는 실제 설비명 기록
      const logItem = { ...item, 적용설비: 실제사용설비, 표준설비명: 실제사용설비, isCommonPart: true };
      addLog(action || '출고', logItem, 현재수량 - oldQuantity, user || 'Manual');
      addFacilityLog(action || '출고', logItem, 현재수량 - oldQuantity, user || 'Manual');
      checkAndNotifyLowStock(data);
      console.log(`🏭 공통부품 출고 기록 — ${item.모델명} → 실제설비: ${실제사용설비}`);
      return res.status(200).json({ success: true, message: '공통부품 출고 완료', data: item });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    console.error('❌ common-update 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message, conversationHistory, user } = req.body;

    invalidateCache();
    let inventoryData = await fetchExcelFromOneDrive();

    // ── 수정1·2: 공통 시트 여부를 원본시트 기준으로 정확히 판별 ──
    const commonItems = inventoryData.filter(item => item.원본시트 === '공통');
    const commonItemsList = commonItems.length > 0
      ? commonItems.map(item =>
          `  · ${item.모델명} (${item.부품종류} / 분류: ${item.적용설비} / 용도: ${String(item.용도 || '').slice(0, 40)})`
        ).join('\n')
      : '  (없음)';

    // 공통 탭 제외 실제 설비 목록 (공통부품 출고 시 선택지 제공용)
    const realFacilities = [...new Set(
      inventoryData
        .filter(d => d.원본시트 !== '공통')
        .map(d => String(d.표준설비명 || d.적용설비 || '').replace(/[\r\n]+/g, ' ').trim())
        .filter(Boolean)
    )].sort();

    // ── 수정4: 재고 현황에 핵심 정보 모두 포함 ──
    const inventoryTable = inventoryData.map(item => {
      const isCommon = item.원본시트 === '공통';
      const stockStatus = item.최소보유수량 > 0 && item.현재수량 <= item.최소보유수량 ? '⚠️부족' : '정상';
      const facilityLabel = isCommon
        ? `공통시트(분류:${item.적용설비})`
        : (item.표준설비명 || item.적용설비);
      return `원본시트:${item.원본시트} | 설비:${facilityLabel} | 모델명:${item.모델명} | 부품종류:${item.부품종류} | 현재수량:${item.현재수량} | 최소보유:${item.최소보유수량} | 재고:${stockStatus}`;
    }).join('\n');

    // ── 수정3: system_instruction 파라미터로 분리 ──
    const systemInstruction = `당신은 스마트 재고 관리 AI 어시스턴트입니다.
반드시 아래 [최신 재고 현황]만을 근거로 답변하고, 목록에 없는 부품은 없다고 명확히 말하세요.

[최신 재고 현황]
${inventoryTable}

[공통 시트 부품 목록 — 여러 설비 공용 부품]
${commonItemsList}

[등록된 실제 설비 목록 — 공통부품 출고 시 설비 선택 참고]
${realFacilities.join(', ')}

[절대 준수 규칙]
1. 입출고 처리 시 반드시 "원본시트(충전/타정/공통)" 값을 JSON에 포함할 것.
2. 마크다운 코드블록(\`\`\`json) 절대 금지. 반드시 ~~~ 기호만 사용.
3. 한 번에 50개 이상 변동 요청 시 두 번 재확인할 것.
4. 모델명 매칭 시 공백·대소문자 차이는 무시하고 찾을 것.
5. 재고 현황에 없는 부품은 "등록된 부품이 아닙니다"라고 명확히 답할 것.

[공통 시트 부품 출고 규칙 — 예외 없음]
6. 원본시트가 "공통"인 부품 출고 시, 사용자가 설비명을 명시하지 않았다면
   절대로 INVENTORY_UPDATE 명령 생성 금지. 반드시 먼저 질문:
   "어느 설비에 사용하실 예정인가요? (예: ${realFacilities.slice(0, 3).join(' / ')})"
7. 설비명 확인 후에만 실제사용설비 필드를 포함하여 명령 생성.
8. "그냥 출고", "확인 생략" 요청에도 공통 부품이면 설비 확인 절대 생략 금지.

[응답 형식 — 충전/타정 시트 일반 부품]
설명 후 마지막에:
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1, "원본시트": "충전"}]}
~~~

[응답 형식 — 공통 시트 부품 (설비 확인 완료 후)]
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1, "원본시트": "공통", "실제사용설비": "확인된설비명"}]}
~~~`;

    // 대화 이력만 contents에, systemInstruction은 별도 파라미터로
    const contents = [];
    if (conversationHistory?.length > 0) {
      conversationHistory.forEach(msg => {
        contents.push({ role: msg.role === 'model' ? 'model' : 'user', parts: [{ text: msg.text }] });
      });
    }
    contents.push({ role: 'user', parts: [{ text: message }] });

    const result = await model.generateContent({
      contents,
      systemInstruction: { parts: [{ text: systemInstruction }] }
    });
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

            // 공통부품인 경우 실제사용설비를 이력에 반영
            const logItem = { ...targetItem };
            if (item.실제사용설비) {
              logItem.적용설비 = item.실제사용설비;
              logItem.표준설비명 = item.실제사용설비;
              logItem.isCommonPart = true;
              console.log(`🏭 공통부품 실제 사용 설비: ${item.실제사용설비}`);
            }

            addLog(action, logItem, finalChange, user || 'AI 어시스턴트');
            addFacilityLog(action, logItem, finalChange, user || 'AI 어시스턴트');
          }
        }

        const success = await updateExcelOnOneDrive(inventoryData);
        if (success) {
          inventoryUpdated = true;
          updateResult = { success: true, action, items };
          checkAndNotifyLowStock(inventoryData); // Teams 저재고 알림
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

  if (CONFIG.teamsWebhookUrl) {
    console.log('✅ Teams Webhook 설정됨 - 재고 부족 알림 활성화');

    // 매일 오전 9시 (KST) 정기 재고 부족 알림
    // alertCooldown을 초기화하여 정기 체크는 항상 발송되도록 함
    setInterval(async () => {
      const kstHour = new Date(Date.now() + 9 * 60 * 60 * 1000).getUTCHours();
      const kstMin  = new Date(Date.now() + 9 * 60 * 60 * 1000).getUTCMinutes();
      if (kstHour === 9 && kstMin < 5) {
        console.log('⏰ 오전 9시 정기 재고 체크 실행');
        alertCooldown.clear(); // 정기 체크는 쿨다운 무시하고 전체 발송
        try {
          const data = await fetchExcelFromOneDrive();
          const lowStock = data.filter(d => d.최소보유수량 > 0 && d.현재수량 <= d.최소보유수량);
          if (lowStock.length > 0) {
            await sendTeamsAlert(lowStock);
          } else {
            console.log('✅ 정기 체크: 저재고 항목 없음');
          }
        } catch (err) {
          console.error('❌ 정기 재고 체크 오류:', err.message);
        }
      }
    }, 5 * 60 * 1000); // 5분마다 시각 확인

  } else {
    console.log('⚠️ TEAMS_WEBHOOK_URL 미설정 - Teams 알림 비활성화');
  }
});
