/** @OnlyCurrentDoc */

/**
 * 정보컴퓨터학습튜터 - Google Sheets & Apps Script Web App
 * 중·고등학교 『정보』 (III. 알고리즘과 프로그래밍 - 문제 이해와 분석 · 추상화 · 알고리즘과 제어 구조)
 * 최소 권한 원칙: 현재 바인딩된 스프레드시트(@OnlyCurrentDoc)와 UrlFetchApp만 사용합니다.
 */

var TASK_SCHEMA_VERSION = 'v8_20261101';

function normalizeStudentCode_(rawCode) {
  var str = String(rawCode || '').trim().toUpperCase();
  var m = str.match(/^S?0*([1-9]|[1-9][0-9])$/);
  if (m) {
    var num = parseInt(m[1], 10);
    return 'S' + (num < 10 ? '0' + num : String(num));
  }
  return str;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('정보컴퓨터학습튜터')
    .addItem('초기 시트 및 20명 학급 코드 세팅/업데이트 (setup)', 'setup')
    .addToUi();
}

function doGet(e) {
  ensureSetup_();
  if (e && e.parameter && e.parameter.action) {
    return handleHttpApi_(e.parameter);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('정보컴퓨터학습튜터: 추상화와 알고리즘')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  ensureSetup_();
  var params = {};
  if (e && e.postData && e.postData.contents) {
    try {
      params = JSON.parse(e.postData.contents);
    } catch (err) {
      params = (e && e.parameter) || {};
    }
  } else if (e && e.parameter) {
    params = e.parameter;
  }
  return handleHttpApi_(params);
}

function handleHttpApi_(params) {
  var action = String(params.action || '');
  var result = { ok: false, message: '알 수 없는 요청입니다.' };
  if (action === 'enterStudent') {
    result = enterStudent(params.code);
  } else if (action === 'submitStudentAnswer') {
    result = submitStudentAnswer(params);
  } else if (action === 'askTutorQuestion') {
    result = askTutorQuestion(params);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 시트 4개(학생명단, 과제, 제출, 대화기록)가 없거나 과제 문항 버전이 업데이트된 경우 자동 반영합니다.
 * 다인원 동시 접속 시 속도 최적화를 위해 캐시 및 버전 플래그를 우선 확인합니다.
 */
function ensureSetup_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('SETUP_VER') === TASK_SCHEMA_VERSION) {
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var currentVer = props.getProperty('TASK_SCHEMA_VERSION');
  if (
    !ss.getSheetByName('학생명단') ||
    !ss.getSheetByName('과제') ||
    !ss.getSheetByName('제출') ||
    !ss.getSheetByName('대화기록') ||
    currentVer !== TASK_SCHEMA_VERSION
  ) {
    setup();
    props.setProperty('TASK_SCHEMA_VERSION', TASK_SCHEMA_VERSION);
  }
  cache.put('SETUP_VER', TASK_SCHEMA_VERSION, 21600);
}

function getCachedTasks_(ss) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('TASKS_' + TASK_SCHEMA_VERSION);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  var taskSheet = ss.getSheetByName('과제');
  var taskData = taskSheet.getDataRange().getValues();
  var tasks = [];
  for (var t = 1; t < taskData.length; t++) {
    if (!taskData[t][0]) continue;
    var dueVal = taskData[t][9];
    var dueStr = dueVal instanceof Date
      ? Utilities.formatDate(dueVal, 'Asia/Seoul', 'yyyy-MM-dd')
      : String(dueVal || '');
    tasks.push({
      id: String(taskData[t][0]).trim(),
      unit: String(taskData[t][1] || ''),
      title: String(taskData[t][2] || ''),
      description: String(taskData[t][3] || ''),
      step1Question: String(taskData[t][4] || ''),
      step2Question: String(taskData[t][5] || ''),
      step3Question: String(taskData[t][6] || ''),
      coachingGuide: String(taskData[t][7] || ''),
      practiceHint: String(taskData[t][8] || ''),
      dueDate: dueStr
    });
  }
  try {
    cache.put('TASKS_' + TASK_SCHEMA_VERSION, JSON.stringify(tasks), 3600);
  } catch (e) {}
  return tasks;
}

/**
 * 구글 시트에 학생명단, 과제, 제출, 대화기록 4개 탭을 만들고
 * 가상 학생(민트, 레몬, 자두)과 중학생 눈높이에 맞춘 『정보』 교과서 기반 과제 4개를 세팅합니다.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 학생명단(코드, 별명, 학습수준, 튜터 지시, 진행상태) — 20명 학급(S01~S20) 자동 발급 (괄호 번호 제거)
  var studentSheet = getOrCreateSheet_(ss, '학생명단', [
    '코드', '별명', '학습수준', '튜터 지시', '진행상태'
  ], '#1b5e20');
  if (studentSheet.getLastColumn() >= 6) {
    studentSheet.getRange(1, 6, Math.max(1, studentSheet.getLastRow()), studentSheet.getLastColumn() - 5).clear();
  }

  var nicknames = [
    '민트', '레몬', '자두', '체리', '피치',
    '멜론', '망고', '베리', '라임', '키위',
    '유자', '살구', '포도', '사과', '다래',
    '모과', '석류', '감귤', '자몽', '호두'
  ];
  var guideBasic = '어려운 정보 교과 용어는 일상적인 비유(마리오 게임 월드맵, 지하철 노선도, 자판기 등)로 쉽게 풀어서 설명하고, 한 번에 한 가지씩만 친절하게 질문하며 힌트를 주세요.';
  var guideStandard = '교과서의 핵심 개념(불필요한 요소 제거, 핵심 요소 추출, 모델링, 제어 구조)을 단계별로 짚어 주며, 학생이 작성한 답안에서 잘한 점과 보완할 점을 구체적으로 안내하세요.';
  var guideAdvanced = '정답을 바로 알려주지 말고, 학생이 스스로 빠뜨린 핵심 요소나 예외 상황(반례), 그래프 모델의 구조적 특징을 찾아내도록 소크라테스식 발문을 던지세요.';

  var existingMap = {};
  if (studentSheet.getLastRow() > 1) {
    var existRows = studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 5).getValues();
    for (var r = 0; r < existRows.length; r++) {
      var rowArr = existRows[r];
      var cKey = normalizeStudentCode_(rowArr[0]);
      // 별명에서 (1번) 등 괄호 번호 자동 제거
      rowArr[1] = String(rowArr[1] || '').replace(/\(\d+번\)/g, '').trim();
      existingMap[cKey] = rowArr.slice(0, 5);
    }
  }

  var students = [];
  for (var idx = 1; idx <= 20; idx++) {
    var sCode = 'S' + (idx < 10 ? '0' + idx : idx);
    if (existingMap[sCode]) {
      var existRow = existingMap[sCode];
      existRow[1] = nicknames[idx - 1];
      students.push(existRow);
    } else {
      var lvl = (idx % 3 === 1) ? '기초' : ((idx % 3 === 2) ? '심화' : '기본');
      var guide = (lvl === '기초') ? guideBasic : ((lvl === '심화') ? guideAdvanced : guideStandard);
      students.push([sCode, nicknames[idx - 1], lvl, guide, '미제출']);
    }
  }
  studentSheet.getRange(2, 1, students.length, 5).setValues(students);
  studentSheet.setColumnWidth(1, 80);
  studentSheet.setColumnWidth(2, 100);
  studentSheet.setColumnWidth(3, 90);
  studentSheet.setColumnWidth(4, 420);
  studentSheet.setColumnWidth(5, 120);

  // 2. 과제(번호, 단원, 제목, 설명, 1단계질문, 2단계질문, 3단계질문, 튜터코칭지침, 연습모드힌트, 마감)
  var taskSheet = getOrCreateSheet_(ss, '과제', [
    '번호', '단원', '제목', '설명', '1단계질문', '2단계질문', '3단계질문', '튜터코칭지침', '연습모드힌트', '마감'
  ], '#c62828');

  var tasks = [
    [
      'T03',
      'III-1-1. 문제 이해와 분석',
      '[문제 분석] 생활 속 문제의 초기 상태 · 현재 상태 · 목표 상태 분석하기',
      '교과서 p.99 & p.101~103(문제 이해와 분석): 아무 변화가 없는 처음 상황인 \'초기 상태\', 문제를 해결해 나가는 과정인 \'현재 상태\', 문제가 해결된 \'목표 상태\'를 구분하고 나만의 핵심 해결 방법을 찾아봅시다.',
      '[1단계: 초기 상태와 목표 상태 정의] 내가 해결하고 싶은 생활 속 문제(예: 아침 등교 지각, 과제 마감 지연 등)를 하나 정하고, 아무 변화가 없는 처음 상황인 \'초기 상태\'와 최종적으로 이루고 싶은 \'목표 상태\'를 적어 보세요.',
      '[2단계: 현재 상태 점검과 방해 원인 찾기] 문제를 해결해 나가는 과정인 \'현재 상태\'에서 목표 상태로 가는 것을 방해하는 가장 큰 원인은 무엇인가요? (힌트: 늦게 자는 습관, 알람 못 듣기, 스마트폰 사용 등 내가 직접 바꿀 수 있는 행동 중에서 1가지를 찾아보세요.)',
      '[3단계: 핵심 해결 실천 방법 정하기] 2단계에서 찾은 방해 원인을 해결하고 목표 상태에 도달하기 위해 내가 오늘부터 실천할 가장 중요한 해결 방법 1가지를 구체적으로 적어 보세요. (주의: 아래 입력창의 힌트 예시를 그대로 베끼지 말고 자신의 실제 상황에 맞게 작성하세요!)',
      '교과서 p.99~103 문제 이해와 분석 개념 적용. 아무 변화가 없는 처음 상태는 \'초기 상태\', 문제를 해결해 나가는 과정은 \'현재 상태\', 최종 개선된 상태는 \'목표 상태\'임을 구분해 코칭하세요. 학생이 힌트 예시를 그대로 베끼지 않고 자기 생활 속 실제 원인과 실천 방법을 구체적으로 적도록 격려하세요.',
      '\'초기 상태\'는 문제가 전혀 해결되지 않은 처음 상태이고, \'현재 상태\'는 문제를 해결하려고 노력해 나가는 과정이며, \'목표 상태\'는 문제가 해결된 모습이에요! 날씨나 버스 배차처럼 내가 바꿀 수 없는 것 말고, 취침 시간·알람 맞추기·준비물 챙기기처럼 내가 직접 실천할 수 있는 구체적인 방법을 적어 보세요.',
      '2026-11-15'
    ],
    [
      'T01',
      'III-1-2. 추상화 (핵심 주력)',
      '[추상화 1] 마리오 월드 지형도에서 길 찾기 노선도 추상화하기',
      '교과서 p.98(지하철 노선도) & p.106(시티 투어 노선도 제작): 복잡한 마리오 월드 지도(마리오의 집 → 쿠파 성)에서 길 찾기에 필요 없는 배경·장식을 지우고, 핵심 정보만 뽑아 가장 빠른 길을 찾아봅시다.',
      '[1단계: 불필요한 요소 제거] 위 마리오 월드 지도에서 \'가장 빠른 길 찾기\'와 관련이 없어서 지워도 되는 배경·장식 요소를 3가지 이상 찾아 적어 보세요.',
      '[2단계: 핵심 요소 추출] 배경 장식을 모두 지웠을 때, 목적지까지 가는 빠른 길을 찾기 위해 노선도에 꼭 남겨야 하는 핵심 정보는 무엇인가요?',
      '[3단계: 가장 빠른 길 찾기] 점과 선으로 단순하게 바꾼 노선도를 보고, [마리오의 집]에서 [쿠파 성]까지 가장 빠르게 갈 수 있는 경로는 어디를 거쳐 가는 길이며 총 몇 분이 걸리는지 적어 보세요.',
      '교과서 p.98~106 추상화 개념 적용. 한 번에 하나씩만 명료하게 질문하세요. 1단계에서는 구름, 버섯 나무, 뻐끔플라워, 코인, 구불구불한 길 모양 등 길 찾기와 무관한 장식을 찾게 하고, 2단계에서는 장소(거점), 연결된 길, 걸리는 시간(분)을 뽑게 하며, 3단계에서는 위쪽 길(2+4+3=9분)과 아래쪽 길(2+2+1+2=7분) 중 가장 빠른 길(7분)을 찾도록 안내하세요.',
      '지하철 노선도에는 실제 나무나 건물 모양이 그려져 있지 않아요! 1단계에서는 \'구름, 버섯 나무, 뻐끔플라워, 코인, 구불구불한 길 모양\'처럼 길 찾기와 상관없는 장식을 고르고, 2단계에서는 \'주요 장소(거점), 연결된 길, 이동 시간(분)\'을 적은 뒤, 3단계에서 위쪽 길(9분)과 아래쪽 길(7분) 중 더 빠른 길을 골라 보세요!',
      '2026-11-22'
    ],
    [
      'T02',
      'III-1-2. 추상화 (심화 모델링)',
      '[추상화 2] 쾨니히스베르크의 7개 다리와 섬 그래프 추상화하기',
      '교과서 p.105(쾨니히스베르크의 다리 문제): 강줄기와 4개의 육지·섬(A, B, C, D)을 잇는 7개의 다리를 한 번씩만 건널 수 있는지 알아보기 위해, 복잡한 마을 지도를 점과 선의 그래프로 단순하게 추상화해 봅시다.',
      '[1단계: 불필요한 정보 제거] 왼쪽 마을 지도에서 \'7개의 다리를 한 번씩만 모두 건널 수 있는가?\'라는 문제를 해결할 때, 문제 해결과 상관없어서 무시해도 되는 그림 속 요소를 3가지 이상 찾아 적어 보세요.',
      '[2단계: 핵심 요소 추출] 복잡한 마을 풍경을 모두 지운 뒤, 오른쪽 그림처럼 단순한 그래프로 표현하기 위해 반드시 남겨야 하는 핵심 정보는 무엇인가요?',
      '[3단계: 그래프에서 선의 개수 세기] 오른쪽 그래프처럼 육지(A, B, C, D)를 \'점\'으로, 다리를 \'선\'으로 바꾸었을 때, 각 점 A, B, C, D에 연결된 선(다리)은 각각 몇 개인지 적어 보세요.',
      '교과서 p.105 쾨니히스베르크 다리 추상화 개념 적용. 1단계에서는 그림에 표시된 다리의 길이(m), 강물의 폭(m), 유람선 배, 가로수 나무, 집, 풍차, 꽃밭 등 한붓그리기와 무관한 요소를 찾게 하세요. 2단계에서는 4개의 육지 구역(A, B, C, D)과 다리의 연결 관계를 찾게 하고, 3단계에서는 각 점에 연결된 선의 개수(A:3개, B:5개, C:3개, D:3개)만 정확히 세어 보도록 안내하세요.',
      '수학자 오일러는 복잡한 마을 지도를 그대로 보지 않고, 육지(A, B, C, D)는 \'점\'으로, 다리는 \'선\'으로 단순화(추상화)했어요! 그림 속 집, 나무, 배, 다리 길이(m), 강폭(m)은 다리를 건너는 순서와 상관이 없답니다. 오른쪽 그래프에서 점 A, B, C, D에 연결된 선이 각각 몇 개인지 세어 보세요!',
      '2026-11-29'
    ],
    [
      'T04',
      'III-1-3. 알고리즘과 제어 구조',
      '[알고리즘 설계] 스마트 음료 자판기 블록 조립으로 제어 구조 완성하기',
      '교과서 p.109~113 & p.157(알고리즘의 세 가지 제어 구조): 자판기의 음료 버튼을 직접 눌러 음료가 나오는 과정을 체험해 보고, 1단계 빈칸(2·4단계)을 채운 뒤 스크래치 블록을 마우스로 드래그하여 선택 구조(2단계)와 반복 구조(3단계)를 완성해 봅시다.',
      '[1단계: 순차 구조 빈칸 채우기] 자판기에서 음료를 뽑는 기본 순서입니다. 빈칸 ( ㉠ ), ( ㉡ )에 들어갈 알맞은 말을 적어 보세요.\n• 1단계: 동전이나 지폐를 투입한다 → 2단계: ( ㉠ ) → 3단계: 배출구로 나온 음료를 꺼낸다 → 4단계: ( ㉡ )',
      '[2단계: 선택 구조 스크래치 블록 드래그 조립] 위 스크래치 조립판으로 블록을 직접 마우스로 끌어다 놓아(Drag & Drop) 선택 구조 블록을 연결해 보세요! ("만약 ~라면" 먼저 놓기 / "아니면(부족하다면)" 먼저 놓기 모두 의미가 통하면 복수 정답으로 인정됩니다.)',
      '[3단계: 반복 구조 스크래치 블록 드래그 조립] 위 스크래치 조립판으로 블록을 직접 마우스로 끌어다 놓아(Drag & Drop) 금액이 부족할 때 목표 금액이 될 때까지 반복하는 블록을 연결해 보세요!',
      '교과서 p.113 순차(Sequential), 선택(Selection), 반복(Iteration) 3대 제어 구조와 p.157 음료 자판기 예제를 쉽게 안내하세요. 1단계는 2단계(음료 버튼 누르기)와 4단계(거스름돈 반환받기) 빈칸 채우기이며, 2·3단계는 스크래치 블록을 마우스로 직접 드래그해 연결하도록 도와주세요.',
      '자판기는 ①차례대로 진행되는 [순차 구조(1단계: 동전 투입 → 2단계(㉠): 음료 버튼 누르기 → 3단계: 음료 꺼내기 → 4단계(㉡): 거스름돈 받기)], ②돈이 충분한지 판단하는 [선택 구조], ③목표 금액이 될 때까지 되풀이하는 [반복 구조]로 작동해요! 위쪽의 스크래치 블록을 마우스로 직접 끌어다 조립판에 연결해 보세요!',
      '2026-12-06'
    ]
  ];

  // 과제 시트 최신 문항으로 업데이트
  if (taskSheet.getLastRow() > 1) {
    taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 10).clearContent();
  }
  taskSheet.getRange(2, 1, tasks.length, tasks[0].length).setValues(tasks);
  taskSheet.setColumnWidth(1, 70);
  taskSheet.setColumnWidth(2, 160);
  taskSheet.setColumnWidth(3, 280);
  taskSheet.setColumnWidth(4, 320);
  taskSheet.setColumnWidth(5, 280);
  taskSheet.setColumnWidth(6, 280);
  taskSheet.setColumnWidth(7, 280);
  taskSheet.setColumnWidth(8, 300);
  taskSheet.setColumnWidth(9, 300);
  taskSheet.setColumnWidth(10, 100);

  // 3. 제출(시각, 코드, 과제번호, 1단계답안, 2단계답안, 3단계답안, 종합피드백) — 학생별·과제별 최종 답안 1건만 유지(중복 행 자동 정리)
  var submitSheet = getOrCreateSheet_(ss, '제출', [
    '시각', '코드', '과제번호', '1단계답안', '2단계답안', '3단계답안', '종합피드백'
  ], '#1565c0');
  submitSheet.setColumnWidth(1, 150);
  submitSheet.setColumnWidth(2, 80);
  submitSheet.setColumnWidth(3, 90);
  submitSheet.setColumnWidth(4, 260);
  submitSheet.setColumnWidth(5, 260);
  submitSheet.setColumnWidth(6, 260);
  submitSheet.setColumnWidth(7, 300);

  if (submitSheet.getLastRow() > 2) {
    var allSub = submitSheet.getRange(2, 1, submitSheet.getLastRow() - 1, 7).getValues();
    var latestByKey = {};
    var orderKeys = [];
    for (var sIdx = 0; sIdx < allSub.length; sIdx++) {
      var rCode = normalizeStudentCode_(allSub[sIdx][1]);
      var rTask = String(allSub[sIdx][2] || '').trim();
      if (!rCode || !rTask) continue;
      var k = rCode + '__' + rTask;
      allSub[sIdx][1] = rCode;
      if (latestByKey[k] && latestByKey[k][6] && !allSub[sIdx][6]) {
        allSub[sIdx][6] = latestByKey[k][6];
      }
      if (!latestByKey[k]) orderKeys.push(k);
      latestByKey[k] = allSub[sIdx];
    }
    if (orderKeys.length < allSub.length) {
      var deduped = [];
      for (var okIdx = 0; okIdx < orderKeys.length; okIdx++) {
        deduped.push(latestByKey[orderKeys[okIdx]]);
      }
      submitSheet.getRange(2, 1, submitSheet.getLastRow() - 1, 7).clearContent();
      if (deduped.length > 0) {
        submitSheet.getRange(2, 1, deduped.length, 7).setValues(deduped);
      }
    }
  }

  // 4. 대화기록(시각, 코드, 과제번호, 질문, 답)
  var chatSheet = getOrCreateSheet_(ss, '대화기록', [
    '시각', '코드', '과제번호', '질문', '답'
  ], '#6a1b9a');
  chatSheet.setColumnWidth(1, 150);
  chatSheet.setColumnWidth(2, 80);
  chatSheet.setColumnWidth(3, 90);
  chatSheet.setColumnWidth(4, 300);
  chatSheet.setColumnWidth(5, 420);

  // 기본 시트(Sheet1/시트1)가 비어 있으면 정리
  var defaultSheet = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }

  PropertiesService.getScriptProperties().setProperty('TASK_SCHEMA_VERSION', TASK_SCHEMA_VERSION);
  return '초기 시트 세팅 및 중학생 눈높이 과제 업데이트(학생명단, 과제, 제출, 대화기록)가 완료되었습니다.';
}

function getOrCreateSheet_(ss, name, headers, headerColor) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground(headerColor || '#37474f')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  return sheet;
}

function formatNow_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 학생 코드로 입장하여 학생 정보, 과제 목록, 본인 전용 제출 답안(및 선생님의 종합피드백), 본인 전용 대화 기록만 불러옵니다.
 * (다른 학생의 제출 기록이나 대화 기록은 절대 반환하지 않아 개인별 시트 보안을 철저히 유지합니다.)
 */
function enterStudent(rawCode) {
  ensureSetup_();
  var code = normalizeStudentCode_(rawCode);
  if (!code) {
    return { ok: false, message: '학생 코드를 입력해 주세요. (예: S1 ~ S20 또는 S01 ~ S20)' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생명단');
  var studentData = studentSheet.getDataRange().getValues();
  var student = null;

  for (var i = 1; i < studentData.length; i++) {
    if (normalizeStudentCode_(studentData[i][0]) === code) {
      student = {
        code: normalizeStudentCode_(studentData[i][0]),
        nickname: String(studentData[i][1] || '').replace(/\(\d+번\)/g, '').trim(),
        level: String(studentData[i][2] || '기본'),
        instruction: String(studentData[i][3] || ''),
        status: String(studentData[i][4] || '미제출')
      };
      break;
    }
  }

  if (!student) {
    return {
      ok: false,
      message: '등록되지 않은 학생 코드입니다. S1 ~ S20 (또는 S01 ~ S20) 범위의 번호를 입력해 주세요.'
    };
  }

  // 과제 목록 읽기 (CacheService 고속 캐싱 적용 — 20명 동시 접속 시 지연 방지)
  var tasks = getCachedTasks_(ss);

  // 해당 학생(code) 본인의 과제별 최종 제출 내역만 필터링하여 읽기 (타 학생 데이터 원천 차단)
  var submitSheet = ss.getSheetByName('제출');
  var submitData = submitSheet.getDataRange().getValues();
  var submissionsByTask = {};
  for (var s = 1; s < submitData.length; s++) {
    var sCode = normalizeStudentCode_(submitData[s][1]);
    if (sCode === code) {
      var tId = String(submitData[s][2] || '').trim();
      var timeVal = submitData[s][0];
      var timeStr = timeVal instanceof Date
        ? Utilities.formatDate(timeVal, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
        : String(timeVal || '');
      submissionsByTask[tId] = {
        timestamp: timeStr,
        taskId: tId,
        step1: String(submitData[s][3] || ''),
        step2: String(submitData[s][4] || ''),
        step3: String(submitData[s][5] || ''),
        teacherFeedback: String(submitData[s][6] || '')
      };
    }
  }

  // 해당 학생(code) 본인의 대화기록만 필터링하여 읽기 (타 학생 데이터 원천 차단)
  var chatSheet = ss.getSheetByName('대화기록');
  var chatData = chatSheet.getDataRange().getValues();
  var chats = [];
  for (var c = 1; c < chatData.length; c++) {
    var cCode = normalizeStudentCode_(chatData[c][1]);
    if (cCode === code) {
      var cTime = chatData[c][0];
      var cTimeStr = cTime instanceof Date
        ? Utilities.formatDate(cTime, 'Asia/Seoul', 'MM-dd HH:mm')
        : String(cTime || '');
      chats.push({
        timestamp: cTimeStr,
        taskId: String(chatData[c][2] || ''),
        question: String(chatData[c][3] || ''),
        answer: String(chatData[c][4] || '')
      });
    }
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var sheetUrl = ss.getUrl();
  var submitSheetId = submitSheet.getSheetId();

  return {
    ok: true,
    student: student,
    tasks: tasks,
    submissions: submissionsByTask,
    chats: chats.slice(-20),
    sheetUrl: sheetUrl,
    submitSheetUrl: sheetUrl + '#gid=' + submitSheetId,
    aiMode: apiKey ? 'GEMINI_LIVE' : 'PRACTICE_MODE'
  };
}

/**
 * 학생이 1단계~3단계 답안을 제출하면 '제출' 탭에서 해당 (학생코드, 과제번호) 행을 찾아
 * 매번 줄을 늘리지 않고 **최종 답안 딱 1번(한 행)만 덮어쓰기(Upsert) 저장**합니다.
 * 기존 선생님의 종합피드백이 있으면 보존하며, '학생명단'의 진행상태를 '제출완료'로 변경합니다.
 */
function submitStudentAnswer(payload) {
  ensureSetup_();
  var code = normalizeStudentCode_((payload && payload.code) || '');
  var taskId = String((payload && payload.taskId) || 'T01').trim();
  var step1 = String((payload && payload.step1) || '').trim();
  var step2 = String((payload && payload.step2) || '').trim();
  var step3 = String((payload && payload.step3) || '').trim();
  var combinedAnswer = String((payload && payload.answer) || '').trim();

  if (!code) {
    return { ok: false, message: '먼저 학생 코드로 입장해 주세요.' };
  }

  if (!step1 && !step2 && !step3 && combinedAnswer) {
    step1 = combinedAnswer;
  }

  if (!step1 && !step2 && !step3) {
    return { ok: false, message: '1단계~3단계 답안을 작성한 뒤 제출해 주세요.' };
  }

  var lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var submitSheet = ss.getSheetByName('고성탈출_형성평가');
    if (!submitSheet) {
      submitSheet = ss.insertSheet('고성탈출_형성평가');
      submitSheet.getRange(1, 1, 1, 11).setValues([[
        '제출일시', '학생코드', '학생별명', '과제ID', '맵단계', '1단계: 문제 정의 및 목표 분석', 
        '2단계: 핵심 요소 추출 및 경로 배제 분석', '3단계: 최적 탈출 경로 및 총 소요 시간 도출', 
        '클리어시간(분)', '교사평가(루브릭 상/중/하)', '교사용 피드백 환류'
      ]]);
      submitSheet.getRange(1, 1, 1, 11)
        .setBackground('#7f0000')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      submitSheet.setFrozenRows(1);
    }
    var submitData = submitSheet.getDataRange().getValues();
    var nowStr = formatNow_();
    var nickname = String((payload && payload.nickname) || '').trim();
    var stageNum = String((payload && payload.stageNum) || '').trim();
    var clearMinutes = String((payload && payload.clearMinutes) || '').trim();

    var existingTeacherFeedback = '';
    var matchedRowIndices = [];
    for (var i = 1; i < submitData.length; i++) {
      if (
        normalizeStudentCode_(submitData[i][1]) === code &&
        String(submitData[i][2] || '').trim() === taskId
      ) {
        matchedRowIndices.push(i + 1); // 1-based sheet row index
        if (String(submitData[i][6] || '').trim() !== '') {
          existingTeacherFeedback = String(submitData[i][6] || '').trim();
        }
      }
    }

    var rowValues = [
      nowStr,
      code,
      nickname,
      taskId,
      stageNum,
      step1,
      step2,
      step3,
      clearMinutes,
      '', // 교사평가(루브릭)
      existingTeacherFeedback
    ];

    if (matchedRowIndices.length > 0) {
      // 첫 번째 매칭 행에 최종 답안 덮어쓰기 (최종답안 1건만 유지)
      submitSheet.getRange(matchedRowIndices[0], 1, 1, 7).setValues([rowValues]);
      // 만약 이전에 중복 저장된 행이 더 있었다면 아래쪽부터 삭제
      for (var d = matchedRowIndices.length - 1; d >= 1; d--) {
        submitSheet.deleteRow(matchedRowIndices[d]);
      }
    } else {
      submitSheet.appendRow(rowValues);
    }

    updateStudentStatus_(ss, code, '제출완료 (' + taskId + ')');
    SpreadsheetApp.flush();

    var submitSheetUrl = ss.getUrl() + '#gid=' + submitSheet.getSheetId();

    return {
      ok: true,
      timestamp: nowStr,
      taskId: taskId,
      status: '제출완료',
      teacherFeedback: existingTeacherFeedback,
      submitSheetUrl: submitSheetUrl,
      message: '최종 답안 1건이 구글 시트 [제출] 탭에 깔끔하게 업데이트(저장)되었습니다!'
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * 학생이 질문 버튼을 눌렀을 때 실행됩니다.
 */
function askTutorQuestion(payload) {
  ensureSetup_();
  var code = normalizeStudentCode_((payload && payload.code) || '');
  var taskId = String((payload && payload.taskId) || 'T01').trim();
  var question = String((payload && payload.question) || '').trim();
  var step1 = String((payload && payload.step1) || '').trim();
  var step2 = String((payload && payload.step2) || '').trim();
  var step3 = String((payload && payload.step3) || '').trim();

  if (!code) {
    return { ok: false, message: '먼저 학생 코드로 입장해 주세요.' };
  }
  if (!question) {
    return { ok: false, message: '튜터에게 물어볼 질문을 입력해 주세요.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var student = findStudent_(ss, code);
  var task = findTask_(ss, taskId);

  if (!student) {
    return { ok: false, message: '학생 정보를 찾을 수 없습니다.' };
  }

  if (student.status === '미제출') {
    updateStudentStatus_(ss, code, '작성중');
    student.status = '작성중';
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var replyText = '';
  var modeUsed = 'PRACTICE_MODE';

  if (apiKey && String(apiKey).trim() !== '') {
    try {
      replyText = callGeminiTutor_(String(apiKey).trim(), student, task, step1, step2, step3, question);
      modeUsed = 'GEMINI_LIVE';
    } catch (err) {
      replyText = buildSmartPracticeReply_(student, task, step1, step2, step3, question) +
        '\n\n(※ 알림: Gemini API 호출 중 일시적 오류가 발생하여 스마트 연습 모드 코칭으로 전환되었습니다.)';
    }
  } else {
    replyText = buildSmartPracticeReply_(student, task, step1, step2, step3, question);
  }

  var nowStr = formatNow_();
  var chatSheet = ss.getSheetByName('대화기록');
  chatSheet.appendRow([nowStr, code, taskId, question, replyText]);
  SpreadsheetApp.flush();

  return {
    ok: true,
    timestamp: nowStr.substring(5, 16),
    taskId: taskId,
    question: question,
    answer: replyText,
    status: student.status,
    modeUsed: modeUsed
  };
}

function findStudent_(ss, code) {
  var norm = normalizeStudentCode_(code);
  var sheet = ss.getSheetByName('학생명단');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeStudentCode_(data[i][0]) === norm) {
      return {
        code: norm,
        nickname: String(data[i][1] || '').replace(/\(\d+번\)/g, '').trim(),
        level: String(data[i][2] || '기본'),
        instruction: String(data[i][3] || ''),
        status: String(data[i][4] || '미제출')
      };
    }
  }
  return null;
}

function findTask_(ss, taskId) {
  var sheet = ss.getSheetByName('과제');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === taskId) {
      return {
        id: String(data[i][0] || '').trim(),
        unit: String(data[i][1] || ''),
        title: String(data[i][2] || ''),
        description: String(data[i][3] || ''),
        step1Question: String(data[i][4] || ''),
        step2Question: String(data[i][5] || ''),
        step3Question: String(data[i][6] || ''),
        coachingGuide: String(data[i][7] || ''),
        practiceHint: String(data[i][8] || '')
      };
    }
  }
  return {
    id: taskId,
    unit: 'III-1-2. 추상화',
    title: '정보 추상화 과제',
    description: '',
    step1Question: '',
    step2Question: '',
    step3Question: '',
    coachingGuide: '',
    practiceHint: '불필요한 요소를 제거하고 핵심 요소를 추출하여 단순하게 표현해 보세요.'
  };
}

function updateStudentStatus_(ss, code, newStatus) {
  var norm = normalizeStudentCode_(code);
  var sheet = ss.getSheetByName('학생명단');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeStudentCode_(data[i][0]) === norm) {
      sheet.getRange(i + 1, 5).setValue(newStatus);
      break;
    }
  }
}

function callGeminiTutor_(apiKey, student, task, step1, step2, step3, question) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(apiKey);

  var systemInstruction = [
    '당신은 중학교 『정보』 교과의 친절하고 지혜로운 AI 코칭 튜터입니다.',
    '절대로 최종 정답이나 완성된 답안 문장을 그대로 대신 써 주지 마세요.',
    '중학생 눈높이에 맞게 쉬운 단어를 사용하고, 한 번에 딱 한 가지 질문만 던져 학생이 스스로 생각하게 도와주세요.',
    '---',
    '[학생 정보]',
    '- 별명: ' + student.nickname + ' (' + student.code + ')',
    '- 학습수준: ' + student.level,
    '- 학생 맞춤 튜터 지시: ' + student.instruction,
    '---',
    '[현재 과제 정보]',
    '- 단원: ' + task.unit,
    '- 과제 제목: ' + task.title,
    '- 과제 설명: ' + task.description,
    '- 1단계 질문: ' + task.step1Question,
    '- 2단계 질문: ' + task.step2Question,
    '- 3단계 질문: ' + task.step3Question,
    '- 교사 코칭 루브릭: ' + task.coachingGuide
  ].join('\n');

  var userPrompt = [
    '[학생이 현재까지 작성한 단계별 답안 초안]',
    '- 1단계 초안: ' + (step1 || '(아직 작성 전)'),
    '- 2단계 초안: ' + (step2 || '(아직 작성 전)'),
    '- 3단계 초안: ' + (step3 || '(아직 작성 전)'),
    '',
    '[학생의 질문]',
    question,
    '',
    '위 초안과 질문을 바탕으로 중학생이 이해하기 쉽게 3~4문장 이내로 코칭해 주세요.'
  ].join('\n');

  var payload = {
    system_instruction: {
      parts: [{ text: systemInstruction }]
    },
    contents: [
      {
        role: 'user',
        parts: [{ text: userPrompt }]
      }
    ],
    generationConfig: {
      temperature: 0.6,
      maxOutputTokens: 500
    }
  };

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var json = JSON.parse(res.getContentText());
  if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts) {
    return json.candidates[0].content.parts[0].text;
  }
  throw new Error('Gemini 응답 파싱 실패');
}

function buildSmartPracticeReply_(student, task, step1, step2, step3, question) {
  var q = question.toLowerCase();
  var opener = '';
  if (student.level === '기초') {
    opener = '[' + student.nickname + ' · 기초 맞춤 코칭] 어렵지 않아요! 쉬운 비유로 하나씩 같이 생각해 볼까요?\n';
  } else if (student.level === '심화') {
    opener = '[' + student.nickname + ' · 심화 탐구 코칭] 좋은 질문이에요! 핵심 원리를 한 걸음 더 생각해 볼까요?\n';
  } else {
    opener = '[' + student.nickname + ' · 단계별 개념 코칭] 교과서 핵심 개념으로 힌트를 드릴게요!\n';
  }

  var analysisLines = [];

  if (task.id === 'T01') {
    if (q.indexOf('1단계') !== -1 || q.indexOf('불필요') !== -1 || q.indexOf('지워') !== -1) {
      analysisLines.push('• [1단계 힌트] 마리오가 쿠파 성까지 가장 빠른 길을 찾을 때, 하늘의 눈 달린 구름이나 버섯 나무, 토관 속 뻐끔플라워, 구불구불한 길 모양이 시간을 계산하는 데 필요할까요? 길 찾기와 상관없는 장식 3가지를 골라 보세요!');
    } else if (q.indexOf('2단계') !== -1 || q.indexOf('핵심') !== -1 || q.indexOf('남겨') !== -1) {
      analysisLines.push('• [2단계 힌트] 장식을 모두 지웠다면, 빠른 길을 찾기 위해 꼭 알아야 하는 것은 "어느 장소를 거쳐 가는지(거점)", "장소끼리 길이 연결되어 있는지", 그리고 "몇 분이 걸리는지(이동 시간)"예요!');
    } else if (q.indexOf('3단계') !== -1 || q.indexOf('빠른') !== -1 || q.indexOf('경로') !== -1 || q.indexOf('몇 분') !== -1) {
      analysisLines.push('• [3단계 힌트] [1-1 버섯 언덕] 갈림길에서 위쪽 [1-2 해변 다리]로 가는 길(2분+4분+3분)과 아래쪽 [지하 도관 → 요시 섬]으로 가는 길(2분+2분+1분+2분)의 시간을 각각 더해 보세요! 어느 쪽이 더 빠른가요?');
    } else {
      analysisLines.push('• [노선도 추상화 힌트] ' + task.practiceHint);
    }
  } else if (task.id === 'T02') {
    if (q.indexOf('1단계') !== -1 || q.indexOf('불필요') !== -1 || q.indexOf('무시') !== -1) {
      analysisLines.push('• [1단계 힌트] 왼쪽 마을 그림을 자세히 보세요! 강물 위의 유람선 배, 물고기, 강폭(90m), 다리 길이(120m 등), 빨간 지붕 집, 풍차, 가로수 나무들이 "다리를 한 번씩만 건널 수 있는가?"를 따질 때 필요할까요?');
    } else if (q.indexOf('2단계') !== -1 || q.indexOf('핵심') !== -1) {
      analysisLines.push('• [2단계 힌트] 오른쪽 그래프에 무엇만 남아 있는지 살펴보세요! 4개의 육지 구역(점 A, B, C, D)과 각 육지를 연결하는 7개의 다리(선)만 남아 있답니다.');
    } else {
      analysisLines.push('• [3단계 힌트] 오른쪽 그래프에서 점 A, B, C, D에서 뻗어 나오는 노란색 선(다리)이 각각 몇 줄인지 손가락으로 하나씩 세어 적어 보세요!');
    }
  } else if (task.id === 'T03') {
    analysisLines.push('• [문제 이해와 분석 힌트] 아무 변화가 없는 처음 상황은 "초기 상태", 지금 문제를 해결하려고 노력해 나가는 과정은 "현재 상태", 문제가 완전히 해결된 모습은 "목표 상태"예요! 3단계에서는 예시를 그대로 베끼지 말고 내가 오늘 바로 실천할 수 있는 나만의 행동을 1가지 적어 보세요.');
  } else if (task.id === 'T04') {
    analysisLines.push('• [자판기 블록 코칭] 화면에 있는 주황색·보라색 스크래치 블록을 클릭하거나 마우스로 끌어다 빈칸에 맞춰 보세요! 2단계는 "만약 투입 금액 >= 음료 가격 이라면 ~ 아니면 ~" 순서로, 3단계는 "투입 금액 >= 음료 가격이 될 때까지 동전 추가 투입 반복" 순서로 조립하면 됩니다.');
  }

  return opener + analysisLines.join('\n') + '\n\n(안내: 현재 연습 모드로 작동 중이며, 교사가 GEMINI_API_KEY를 설정하면 실시간 생성 코칭으로 전환됩니다.)';
}
