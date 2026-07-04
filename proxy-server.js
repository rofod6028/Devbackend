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
  inventorySheets: ['스페어파트'],
  facilityMasterSheet: '설비마스터',    // 설비 그룹→개별 호기 매핑 테이블
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

// 설비이력 메모리 버퍼
let facilityLogs = [];

function invalidateCache() {
  cachedData = null;
  lastFetchTime = null;
}

// ============================================================
// 설비명 정리 (공백/줄바꿈만 정리 — 별도 매핑 테이블 사용 안 함)
// ============================================================
// ⚠️ 과거에는 '제조' 시트("원본설비명→표준설비명" 매핑 테이블)를 참조했으나,
//    실제로는 관리되지 않는 사문화된 시트였고, 매핑 실수로 호기 표기(#3 등)가
//    지워지는 등 부작용만 있어 완전히 제거했다.
//    이제 적용설비 원본 문자열의 공백/줄바꿈만 정리해서 그대로 표준설비명으로 사용한다.
// ============================================================
// 전각(全角) 특수문자 → 반각 정규화
// ⚠️ 엑셀에서 한글 입력기를 쓰다 실수로 전각 샵(＃, U+FF03) 등이 섞여 들어가면
//    설비명 매칭(설비마스터 대조)이 실패할 수 있다. 모든 판별 로직에 들어가기 전에
//    이 정규화를 거쳐서, 어떤 문자가 섞여 들어와도 안전하게 반각으로 통일한다.
// ============================================================
function normalizeSpecialChars(str) {
  return String(str || '')
    .replace(/＃/g, '#')   // 전각 샵 → 반각 샵
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)) // 전각 숫자 → 반각 숫자
    .replace(/（/g, '(')
    .replace(/）/g, ')');
}

function normalizeEquipment(originalName) {
  return normalizeSpecialChars(String(originalName || '').replace(/[\r\n]+/g, ' ')).replace(/\s+/g, ' ').trim();
}

// ============================================================
// 설비마스터 기반 설비 판별
// ============================================================
// '설비마스터' 시트 구조: 원본시트 | 소속그룹(적용설비 매칭용) | 설비명(호기별)
//   - 여러 호기로 나뉘는 설비(예: 마블충전기#1~4)는 "소속그룹" 하나에 여러 "설비명" 행으로 등록
//   - 스페어파트 시트의 "적용설비" 값이 이 "소속그룹"과 일치하면, 해당 그룹에 속한
//     호기 목록 중 실사용자가 어느 것인지 선택/확인하도록 한다.
//   - "적용설비"가 이미 특정 호기 하나를 직접 가리키면(설비명과 일치) 확인 절차 없이 그대로 사용.
//   - 어느 쪽에도 매칭되지 않고 "~관련"/"기타" 같은 범용 태그이면, 전체 설비 목록 중에서 선택.
// ============================================================
let facilityMasterCache = null; // { groupToUnits, unitToSheet, allUnits, rawRows }

async function loadFacilityMaster(workbook) {
  const sheet = workbook.Sheets[CONFIG.facilityMasterSheet];
  const groupToUnits = {};   // { 소속그룹: [{설비명, 원본시트}] }
  const unitToSheet = {};    // { 설비명: 원본시트 }
  const allUnits = [];       // [{설비명, 원본시트, 소속그룹}]
  const rawRows = [];

  if (sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet);
    rows.forEach(row => {
      const 원본시트 = String(row['원본시트'] || '').trim();
      const 소속그룹 = normalizeEquipment(row['소속그룹(적용설비 매칭용)']);
      const 설비명 = normalizeEquipment(row['설비명(호기별)']);
      if (!설비명) return;
      rawRows.push(row);
      unitToSheet[설비명] = 원본시트 || '미분류';
      allUnits.push({ 설비명, 원본시트: 원본시트 || '미분류', 소속그룹 });
      if (소속그룹) {
        (groupToUnits[소속그룹] = groupToUnits[소속그룹] || []).push({ 설비명, 원본시트: 원본시트 || '미분류' });
      }
    });
    console.log(`🏭 설비마스터 로드: 그룹 ${Object.keys(groupToUnits).length}개, 개별 설비 ${allUnits.length}개`);
  } else {
    console.warn(`⚠️ "${CONFIG.facilityMasterSheet}" 시트를 찾을 수 없습니다 — 설비 확인 기능이 제한됩니다.`);
  }

  facilityMasterCache = { groupToUnits, unitToSheet, allUnits, rawRows };
  return facilityMasterCache;
}

// 스페어파트 한 행의 "적용설비" 원본 문자열을 보고
// 원본시트 / 표준설비명 / 설비확인필요여부 / 후보설비목록 을 판별한다.
function resolveFacility(rawEquipRaw, fm) {
  const cleaned = normalizeEquipment(rawEquipRaw);
  const blank = !cleaned || /^(none|nan|null|undefined)$/i.test(cleaned);

  if (blank) {
    return {
      원본시트: '공통',
      표준설비명: '(설비 미지정)',
      isCommonPart: true,
      후보설비목록: fm.allUnits.map(u => u.설비명).sort(),
      미등록: true
    };
  }

  // 1) 이미 특정 호기 하나를 직접 가리키는 경우 — 확인 불필요
  if (fm.unitToSheet[cleaned]) {
    return { 원본시트: fm.unitToSheet[cleaned], 표준설비명: cleaned, isCommonPart: false, 후보설비목록: null };
  }

  // 2) 여러 호기로 묶인 소속그룹에 매칭되는 경우 — 그 그룹 내에서만 확인
  if (fm.groupToUnits[cleaned]) {
    const units = fm.groupToUnits[cleaned];
    return {
      원본시트: units[0]?.원본시트 || '공통',
      표준설비명: cleaned,
      isCommonPart: true,
      후보설비목록: units.map(u => u.설비명).sort()
    };
  }

  // 3) "~관련" / "기타" 등 범용 태그 — 특정 설비군에 속하지 않음, 전체 목록에서 선택
  if (/관련$/.test(cleaned) || cleaned === '기타') {
    return {
      원본시트: '공통',
      표준설비명: cleaned,
      isCommonPart: true,
      후보설비목록: fm.allUnits.map(u => u.설비명).sort()
    };
  }

  // 4) 설비마스터에 등록되지 않은 개별 설비명 — 그대로 단일 설비로 취급하되 등록 필요 표시
  return { 원본시트: '미분류', 표준설비명: cleaned, isCommonPart: false, 후보설비목록: null, 미등록: true };
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

    // 설비마스터(그룹→호기 매핑) 먼저 로드
    const fm = await loadFacilityMaster(workbook);

    // 스페어파트 시트 하나만 순회
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
        const rawEquip = row['적용설비'] || '';
        const resolved = resolveFacility(rawEquip, fm);

        return {
          id: `${sheetName}_${index + 1}`,
          원본시트: resolved.원본시트,          // UI 상단 탭 분류(충전/타정/공통/미분류) — 화면 표시용, 실제 시트는 스페어파트 하나뿐
          대분류: row['대분류'] || '미분류',
          부품종류: row['부품종류'] || '',
          모델명: row['모델명'] || '',
          적용설비: row['적용설비'] || '',      // 엑셀 원본 그대로 (그룹명 또는 특정 설비명)
          표준설비명: resolved.표준설비명,
          isCommonPart: resolved.isCommonPart,  // true면 실사용 설비 확인 필요
          후보설비목록: resolved.후보설비목록,   // 확인 필요 시 선택 가능한 구체 설비 목록
          미등록설비: !!resolved.미등록,         // 설비마스터에 없는 값 — 진단용
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

    // 설비마스터에 없는(미등록) 적용설비 값을 진단 로그로 안내
    const unregistered = [...new Set(allMappedData.filter(d => d.미등록설비).map(d => d.적용설비))];
    if (unregistered.length > 0) {
      console.warn(`⚠️ 설비마스터에 없는 적용설비 값 ${unregistered.length}개 (임시로 "미분류"/개별설비 처리됨):`, unregistered.slice(0, 20));
    }

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

      // 스페어파트 시트 저장 — 원본시트(화면 표시용 분류)와 무관하게 전체를 하나의 시트에 저장
      const excelRows = data.map(item => ({
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
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(excelRows), '스페어파트');

      // 설비마스터 시트 보존 — 앱에서 직접 수정하지 않는 참조 테이블이므로
      // 로드 시점에 캐시해둔 원본 행을 그대로 다시 써서 유실 방지
      if (facilityMasterCache && facilityMasterCache.rawRows.length > 0) {
        XLSX.utils.book_append_sheet(
          workbook,
          XLSX.utils.json_to_sheet(facilityMasterCache.rawRows),
          CONFIG.facilityMasterSheet
        );
      }

      // 로그 시트 저장
      const logRows = [...memoryLogs].reverse();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(logRows), CONFIG.logSheetName);

      // 설비이력 시트 저장
      if (facilityLogs.length > 0) {
        const facilityRows = [...facilityLogs].reverse();
        const facilityWs = XLSX.utils.json_to_sheet(facilityRows);
        XLSX.utils.book_append_sheet(workbook, facilityWs, CONFIG.facilityLogSheetName);
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
    // 설비마스터(그룹→개별 호기 매핑)도 함께 내려줌 — 프론트엔드가 카드 목록/설비 선택지를
    // 정확한 "개별 호기" 단위로 구성할 수 있도록 함
    const facilityMaster = facilityMasterCache?.allUnits || [];
    res.json({ success: true, data, facilityMaster });
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

    // 설비 확인이 필요한 항목(isCommonPart)은 반드시 common-update로만 처리
    if (item.isCommonPart && Number(현재수량) < Number(item.현재수량)) {
      return res.status(400).json({ success: false, message: '이 부품은 실사용 설비 확인이 필요합니다. common-update를 사용해 주세요.' });
    }

    const oldQuantity = item.현재수량;
    const qtyDelta = Number(현재수량) - Number(oldQuantity);
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();
    item.작업자 = user || 'Manual';

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      try {
        addLog(action || '수정', item, qtyDelta, user || 'Manual');
        addFacilityLog(action || '수정', item, qtyDelta, user || 'Manual');
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

    // 후보설비목록이 정의된 항목이면, 그 목록 안의 설비인지 검증 (오타/임의 입력 방지)
    if (Array.isArray(item.후보설비목록) && item.후보설비목록.length > 0 && !item.후보설비목록.includes(실제사용설비)) {
      return res.status(400).json({ success: false, message: `"${실제사용설비}"는 이 부품의 사용 가능 설비 목록에 없습니다.` });
    }

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

    // 전체 특정 설비 목록(예시/일반 안내용)
    const realFacilities = (facilityMasterCache?.allUnits || []).map(u => u.설비명).sort();

    // ── 수정4: 재고 현황에 핵심 정보 모두 포함 — 설비확인필요 여부와 후보 목록을 그대로 노출 ──
    const inventoryTable = inventoryData.map(item => {
      const stockStatus = item.최소보유수량 > 0 && item.현재수량 <= item.최소보유수량 ? '⚠️부족' : '정상';
      const confirmNote = item.isCommonPart
        ? ` | [설비확인필요] 후보:${(item.후보설비목록 || []).join('/')}`
        : ` | [설비확정:${item.표준설비명}]`;
      return `모델명:${item.모델명} | 부품종류:${item.부품종류}${confirmNote} | 현재수량:${item.현재수량} | 최소보유:${item.최소보유수량} | 재고:${stockStatus}`;
    }).join('\n');

    // ── 수정3: system_instruction 파라미터로 분리 ──
    const systemInstruction = `당신은 스마트 재고 관리 AI 어시스턴트입니다.
반드시 아래 [최신 재고 현황]만을 근거로 답변하고, 목록에 없는 부품은 없다고 명확히 말하세요.

[최신 재고 현황]
각 줄은 하나의 부품입니다. "[설비확인필요]"가 붙은 부품은 여러 설비에서 공용으로 쓰는
부품이라 어느 설비에 사용/입고하는지 반드시 확인해야 하고, 그 뒤 "후보:" 목록이
선택 가능한 설비 목록입니다. "[설비확정:설비명]"이 붙은 부품은 이미 특정 설비 전용이라
추가로 물어볼 필요가 없습니다.
${inventoryTable}

[전체 설비 목록 — 참고용]
${realFacilities.join(', ')}

[절대 준수 규칙]
1. 마크다운 코드블록(\`\`\`json) 절대 금지. 반드시 ~~~ 기호만 사용.
2. 한 번에 50개 이상 변동 요청 시 두 번 재확인할 것.
3. 모델명 매칭 시 공백·대소문자 차이는 무시하고 찾을 것.
4. 재고 현황에 없는 부품은 "등록된 부품이 아닙니다"라고 명확히 답할 것.

[설비 확인 규칙 — 예외 없음]
5. 재고 현황에서 해당 부품이 "[설비확인필요]"로 표시되어 있으면, 사용자가 이미
   메시지에서 구체적 설비(호기)를 명시하지 않은 이상 절대로 INVENTORY_UPDATE 명령을
   생성하지 말고, 반드시 먼저 "어느 설비에 사용/입고하셨나요? (해당 부품의 "후보:" 목록 중에서
   실제 후보들을 예시로 제시)"라고 물어볼 것.
6. 사용자가 설비를 답하면(또는 이미 메시지에 명시했다면), 그 설비명이 해당 부품의
   "후보:" 목록에 있는 정확한 표기와 일치하도록 확인 후 실제사용설비 필드에 포함하여 명령 생성.
7. "[설비확정:설비명]" 부품은 이미 설비가 하나로 확정되어 있으므로 실제사용설비를
   묻지 않고 바로 명령을 생성해도 된다.
8. "그냥 출고", "확인 생략" 요청에도 "[설비확인필요]" 부품이면 설비 확인 절대 생략 금지.

[응답 형식 — 설비확정 부품(설비확인 불필요)]
설명 후 마지막에:
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1}]}
~~~

[응답 형식 — 설비확인필요 부품(설비 확인 완료 후)]
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1, "실제사용설비": "확인된정확한설비명"}]}
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

        // 모델명(공백/대소문자 무시) 기준으로 후보 항목 찾는 헬퍼
        const findCandidates = (모델명) => {
          const normModel = String(모델명 || '').replace(/\s+/g, '').toLowerCase();
          return inventoryData.filter(d => String(d.모델명 || '').replace(/\s+/g, '').toLowerCase() === normModel);
        };

        // ── 백엔드 안전망: 설비 확인이 필요한 부품(isCommonPart)인데 실제사용설비가 없으면 강제 차단 ──
        {
          const missingFacilityItems = items.filter(item => {
            if (item.실제사용설비 && item.실제사용설비.trim()) return false; // 이미 명시됨
            const candidates = findCandidates(item.모델명);
            return candidates.some(c => c.isCommonPart);
          });

          if (missingFacilityItems.length > 0) {
            const lines = missingFacilityItems.map(i => {
              const candidates = findCandidates(i.모델명);
              const match = candidates.find(c => c.isCommonPart) || candidates[0];
              const facs = match?.후보설비목록 || [];
              return `· ${i.모델명} → 후보 설비: ${facs.join(' / ')}`;
            });
            const clarifyMsg = `⚠️ 어느 설비에서 사용/입고하신 건지 확인이 필요합니다.\n${lines.join('\n')}\n\n정확한 설비명을 말씀해 주세요.`;
            console.log(`🚫 설비 미확인 차단: ${missingFacilityItems.map(i => i.모델명).join(', ')}`);
            return res.json({ success: true, message: clarifyMsg, inventoryUpdated: false, updateResult: null, timestamp: new Date().toISOString() });
          }
        }

        for (const item of items) {
          const candidates = findCandidates(item.모델명);

          let targetItem = null;
          if (candidates.length === 1) {
            targetItem = candidates[0];
          } else if (candidates.length > 1) {
            // 여러 후보 중 실제사용설비로 정확히 특정
            const facRaw = String(item.실제사용설비 || '').replace(/[\r\n]+/g, ' ').trim();
            if (facRaw) {
              targetItem = candidates.find(d => Array.isArray(d.후보설비목록) && d.후보설비목록.includes(facRaw))
                || candidates.find(d => d.표준설비명 === facRaw);
            }
            if (!targetItem) {
              // 설비를 특정하지 못하면 절대 임의의 행을 건드리지 않고 스킵 (몰빵 방지)
              console.warn(`⚠️ AI 인벤토리 업데이트: "${item.모델명}" 설비 특정 실패(후보 ${candidates.length}개, 실제사용설비="${facRaw}") → 스킵`);
              continue;
            }
          }

          if (targetItem) {
            const changeQty = Number(item.수량) || 0;
            const finalChange = action === '출고' ? -changeQty : changeQty;
            targetItem.현재수량 = action === '출고' ? Math.max(0, targetItem.현재수량 - changeQty) : targetItem.현재수량 + changeQty;
            targetItem.최종수정시각 = getKSTDate();
            targetItem.작업자 = user || 'AI 어시스턴트';

            // 설비확인이 필요했던 부품은 실제사용설비를 이력에 기록
            const logItem = { ...targetItem };
            if (targetItem.isCommonPart && item.실제사용설비) {
              logItem.적용설비 = item.실제사용설비;
              logItem.표준설비명 = item.실제사용설비;
              console.log(`🏭 실제 사용 설비: ${item.실제사용설비}`);
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
