/** @OnlyCurrentDoc */

/**
 * 정보컴퓨터학습튜터 - Google Sheets & Apps Script Web App
 * 중·고등학교 『정보』 (III. 알고리즘과 프로그래밍 - 문제 이해와 분석 · 추상화 · 알고리즘과 제어 구조)
 * 최소 권한 원칙: 현재 바인딩된 스프레드시트(@OnlyCurrentDoc)와 UrlFetchApp만 사용합니다.
 */

var TASK_SCHEMA_VERSION = 'v10_castle_escape_sequential';

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
    .createMenu('🏰 고성 미궁 추상화 경연대회')
    .addItem('▶ 2026 경연대회 전용 시트 및 과제 4종 자동 세팅 (setup)', 'setup')
    .addItem('▶ 학생 제출 데이터 시트 서식 새로고침', 'setupCastleEscapeSheet_')
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
  var tasks = [
    [
      'T01',
      'III. 문제 해결과 프로그래밍 — 1. 추상화 [9정03-01]',
      '[MAP 1: 용암 지하성당] 불필요한 우회로 가지치기 및 핵심 경로 추출',
      '고대 성채 지하의 용암 카타콤에서 막다른 방인 [C.상층보물실]과 용암 위험 지연로인 [E.화산암우회로]를 배제하고, [D.지하성소교각]과 [F.흑요석룬대교]를 거쳐 6분 최적 탈출 경로를 완성하는 추상화 탐구 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '2022 개정 교육과정 [9정03-01] 추상화 개념 적용. 불필요한 장식(용암 마그마 폭포, 해골 촛대)과 위험 지연로(E: 7분 소요), 막다른 방(C)을 소거하고, [A->B->D->F->G]의 6분 최단 경로를 가중치 방향 그래프로 도출하도록 안내하세요.',
      '막다른 방인 [C.상층보물실(막다른길)]과 위험 지연로인 [E.화산암우회로(느림)]를 배제하고, [D.지하성소교각]과 [F.흑요석룬대교]를 거치면 6분 최적 탈출 경로가 완성됩니다.',
      '2026-11-27'
    ],
    [
      'T02',
      'III. 문제 해결과 프로그래밍 — 1. 추상화 [9정03-01]',
      '[MAP 2: 빙하 수로 미궁] 수저 동굴 우회로 배제 및 V자 수로·승강기 최적화',
      '혹한의 청빙 지하 미궁에서 막다른 [F.얼음수정동굴]과 시간 지연 함정인 [D.빙하우회절벽]을 배제하고, V자 지하 수로와 [C.크리스탈승강기]를 활용하여 7분 최단 경로를 모델링하는 추상화 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '빙하 수로의 고저 차이를 극복하기 위해 승강기(Lift)를 핵심 거점으로 선정하고, 빙하 절벽 우회로(D)와 수정 동굴(F)을 상태 공간에서 배제하는 사고력을 평가합니다.',
      '비효율적인 [D.빙하우회절벽(지연)]과 막다른 길인 [F.얼음수정동굴(막다른길)]을 배제하고, [C.크리스탈승강기]를 타면 7분 최단 경로로 탈출할 수 있습니다.',
      '2026-11-27'
    ],
    [
      'T03',
      'III. 문제 해결과 프로그래밍 — 1. 추상화 [9정03-01]',
      '[MAP 3: 기계 시계탑] 중앙 파괴교각 우회 및 다층 기어 가중치 경로 최적화',
      '거대 톱니바퀴가 회전하는 시계탑에서 통행이 불가한 [C.중앙붕괴교각]을 피하고, [D.서쪽승강기계실]과 [F.상층외벽비탈]을 거쳐 8분 만에 시계탑 최상층 탈출구로 향하는 가중치 그래프 최적화 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '파괴된 교각(C: 통행 불가)을 탐색 대상에서 즉시 소거(Pruning)하고, 회전 기어 승강로를 경유하는 대안 경로의 가중치를 계산하는 문제 해결력을 코칭합니다.',
      '통행이 불가능한 [C.중앙붕괴교각(통행불가)]과 막다른 길 [E.폐기어보관소]를 배제하고, [D.서쪽승강기계실]을 경유하면 8분 만에 탈출할 수 있습니다.',
      '2026-11-27'
    ],
    [
      'T04',
      'III. 문제 해결과 프로그래밍 — 1. 추상화 [9정03-01]',
      '[MAP 4: 심연 최종 탈출] 2400x960 심연 상태 공간 축약 및 불사신 사신 데스 회피',
      '고성 최심부 2배 깊이의 심연 미궁에서 [X1.고문실(막다른길)], [X2.폐쇄철문(통행불가)], [W.미혹마법진(역송함정)]을 가지치기하고, [C.수직엘리베이터]와 [G.심연대교]를 거쳐 8분 만에 대성문으로 탈출하는 최고난도 알고리즘 추상화 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '2400x960의 거대한 상태 공간(State Space)을 노드 14개, 간선 13개의 단순 방향 그래프로 축약하고, 시공 왜곡 함정(W: +5분 지연)을 논리적으로 배제하는 최고난도 추상화 능력을 평가합니다.',
      '막다른 [X1.고문실(막다른길)], [X2.폐쇄철문(통행불가)], [W.미혹마법진(함정)]을 배제하고, [C.승강도크] 엘리베이터로 하강 후 [G.심연대교]를 거쳐 8분 만에 탈출하세요.',
      '2026-11-27'
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
    // 3. 고성탈출_형성평가 전용 시트 생성 및 고딕 와인색 서식 지정
  setupCastleEscapeSheet_(ss);

  // 4. 루브릭_평가기준표 시트 생성
  setupRubricSheet_(ss);

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

function setupCastleEscapeSheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, '고성탈출_형성평가', [
    '제출일시', '학생코드', '학생별명', '과제ID', '맵단계', '1단계: 문제 정의 및 목표 분석', 
    '2단계: 핵심 요소 추출 및 경로 배제 분석', '3단계: 최적 탈출 경로 및 총 소요 시간 도출', 
    '클리어시간(분)', '2022 성취기준 루브릭 평가', '교사 피드백 환류'
  ], '#7f0000');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 260);
  sheet.setColumnWidth(7, 340);
  sheet.setColumnWidth(8, 320);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 160);
  sheet.setColumnWidth(11, 240);
  return sheet;
}

function setupRubricSheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var rSheet = getOrCreateSheet_(ss, '루브릭_평가기준표', [
    '등급', '성취 수준 기준 (2022 개정 [9정03-01])', '세부 채점 지침 (핵심 경로 및 가지치기 분석)'
  ], '#283593');
  var rubricData = [
    ['상 (탁월)', '문제 해결에 필요한 핵심 요소를 완벽히 추출하고, 불필요한 우회로와 함정의 배제 이유를 타당하게 제시하며 최적 경로를 정확히 계산함.', '• 1단계(출발/도착/목표시간) 완벽\n• 2단계: 통행불가/막다른길/우회로 배제 사유 명확\n• 3단계: 최단 시간 합산식 완벽 도출'],
    ['중 (보통)', '핵심 경로는 도출하였으나, 우회로 배제 사유의 구체성이 다소 부족하거나 시간 계산에 사소한 오차가 있음.', '• 핵심 경로는 찾았으나 배제 이유가 단순 기술에 그침\n• 부분적 가중치 계산 누락'],
    ['하 (노력)', '단순 직관이나 추측에 의존하여 경로를 선택하였으며, 핵심 요소 추출과 배제 기준의 연계성이 미흡함.', '• 막다른 길이나 지연 요소를 인지하지 못함\n• 3단계 추상화 논리 서술 부족']
  ];
  if (rSheet.getLastRow() <= 1) {
    rSheet.getRange(2, 1, rubricData.length, 3).setValues(rubricData);
    rSheet.setColumnWidth(1, 90);
    rSheet.setColumnWidth(2, 420);
    rSheet.setColumnWidth(3, 460);
  }
  return rSheet;
}
