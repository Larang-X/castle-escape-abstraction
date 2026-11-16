/** @OnlyCurrentDoc */

/**
 * ============================================================================
 * [2026 교육용 웹콘텐츠 개발 경연대회 출품작]
 * 고성 탈출 게임으로 추상화 배워보기 — Google Apps Script 백엔드 API
 * 2022 개정 교육과정 중학교 『정보』 [9정03-01]·[9정03-02] 문제 상태 구조화 & 추상화 알고리즘 형성평가 시스템
 * ============================================================================
 */

var TASK_SCHEMA_VERSION = 'v12_plain_names';

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
    .addItem('▶ 전체 시트 및 과제 4종 자동 세팅 (setup)', 'setup')
    .addItem('▶ 고성탈출_형성평가 서식 새로고침', 'setupCastleEscapeSheet_')
    .addToUi();
}

function doGet(e) {
  ensureSetup_();
  if (e && e.parameter && e.parameter.action) {
    return handleHttpApi_(e.parameter);
  }
  return HtmlService.createHtmlOutput('<h3>🏰 2026 고성 미궁 추상화 백엔드 API 가동 중</h3><p>정상 연결되었습니다.</p>')
    .setTitle('고성 탈출 추상화 튜터 API')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    !ss.getSheetByName('고성탈출_형성평가') ||
    !ss.getSheetByName('루브릭_평가기준표') ||
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

function getOrCreateSheet_(ss, name, headers, headerBg) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (headers && headers.length > 0) {
    var curHeaders = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, headers.length).getValues()[0] : [];
    var needsHeader = false;
    for (var h = 0; h < headers.length; h++) {
      if (curHeaders[h] !== headers[h]) {
        needsHeader = true;
        break;
      }
    }
    if (needsHeader || sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground(headerBg || '#1a237e')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 학생명단 (S01 ~ S20)
  var studentSheet = getOrCreateSheet_(ss, '학생명단', ['학생코드', '별명', '수준', '맞춤지도지침', '진행상태'], '#1b5e20');
  var nicknames = [
    '민트', '호두', '보라', '바다', '단풍', '노을', '별빛', '초록', '달빛', '하늘',
    '햇살', '은하', '송이', '도토리', '구름', '라임', '체리', '망고', '올리브', '새싹'
  ];
  var guideBasic = '어려운 정보 교과 용어는 일상적인 비유(고성 미궁 맵, 노선도 등)로 쉽게 풀어서 설명하고, 한 번에 한 가지씩만 친절하게 질문하세요.';
  var guideStandard = '1단계(문제 정의), 2단계(핵심 요소 추출 및 경로 배제 분석), 3단계(최적 탈출 경로 및 총 소요 시간 도출)의 연결 관계를 스스로 정리하도록 유도하세요.';
  var guideAdvanced = '단순히 길을 찾는 것을 넘어, 복잡한 맵을 장소와 길만 남긴 노선도로 단순화하고 불필요한 우회로를 사전에 가지치기하는 추상화의 계산 효율성을 깊이 있게 탐구하도록 심화 질문을 던지세요.';

  var students = [];
  for (var idx = 1; idx <= 20; idx++) {
    var sCode = 'S' + (idx < 10 ? '0' + idx : String(idx));
    var lvl = (idx % 3 === 1) ? '기초' : ((idx % 3 === 2) ? '심화' : '기본');
    var guide = (lvl === '기초') ? guideBasic : ((lvl === '심화') ? guideAdvanced : guideStandard);
    students.push([sCode, nicknames[idx - 1], lvl, guide, '미제출']);
  }
  studentSheet.getRange(2, 1, students.length, 5).setValues(students);
  studentSheet.setColumnWidth(1, 80);
  studentSheet.setColumnWidth(2, 100);
  studentSheet.setColumnWidth(3, 90);
  studentSheet.setColumnWidth(4, 420);
  studentSheet.setColumnWidth(5, 120);

  // 2. 과제 (T01 ~ T04: MAP 1 ~ MAP 4 순차)
  var taskSheet = getOrCreateSheet_(ss, '과제', [
    '번호', '단원', '제목', '설명', '1단계질문', '2단계질문', '3단계질문', '튜터코칭지침', '연습모드힌트', '마감'
  ], '#c62828');

  var tasks = [
    [
      'T01',
      'III. 문제 해결과 프로그래밍 — [9정03-01]·[9정03-02]',
      '[MAP 1: 용암 지하성당] 불필요한 우회로 가지치기 및 핵심 경로 추출',
      '고대 성채 지하의 용암 카타콤에서 막다른 방인 [C.보물방]과 용암 위험 지연로인 [E.용암우회로]를 배제하고, [D.지하성소다리]와 [F.흑요석룬대교]를 거쳐 6분 최적 탈출 경로를 완성하는 추상화 탐구 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '2022 개정 교육과정 [9정03-01] 추상화 개념 적용. 불필요한 장식과 위험 지연로(E: 7분 소요), 막다른 방(C)을 소거하고, [A->B->D->F->G]의 6분 최단 경로를 노선도처럼 단순하게 정리하도록 안내하세요.',
      '막다른 방인 [C.보물방(막다른길)]과 위험 지연로인 [E.용암우회로(느림)]를 배제하고, [D.지하성소다리]와 [F.흑요석룬대교]를 거치면 6분 최적 탈출 경로가 완성됩니다.',
      '2026-11-27'
    ],
    [
      'T02',
      'III. 문제 해결과 프로그래밍 — [9정03-01]·[9정03-02]',
      '[MAP 2: 빙하 수로 미궁] 예티동굴·얼음감옥 가지치기 및 크리스탈 승강대 최적 경로',
      '얼어붙은 지하 수로에서 시간이 오래 걸리는 [C.예티동굴(우회로)]와 막다른 [X.얼음감옥(막다른길)]을 배제하고, [D.크리스탈수로]에서 [E.크리스탈승강대]를 타고 7분 만에 탈출하는 추상화 탐구 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '예티가 지키는 우회로(5분)와 막다른 얼음감옥을 왜 빼야 하는지, 크리스탈 승강대를 타면 왜 더 빠른지 시간을 더해서 스스로 설명하도록 코칭합니다.',
      '시간이 오래 걸리는 [C.예티동굴(우회로)]와 막다른 길인 [X.얼음감옥(막다른길)]을 배제하고, [D.크리스탈수로]에서 [E.크리스탈승강대]를 타면 7분 최적 경로로 탈출할 수 있습니다.',
      '2026-11-27'
    ],
    [
      'T03',
      'III. 문제 해결과 프로그래밍 — [9정03-01]·[9정03-02]',
      '[MAP 3: 기계 시계탑] 서쪽 종탑 우회로 가지치기 및 중앙 태엽 승강기 최적 경로',
      '거대 톱니바퀴가 도는 시계탑에서 느린 [C.서쪽종탑(우회로)]과 막다른 [X.태엽창고(막다른길)]를 배제하고, [D.아래층진자회랑]에서 [F.중앙태엽승강기]를 타고 8분 만에 시계탑 꼭대기 종각으로 가는 추상화 탐구 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '4분이 더 걸리는 서쪽 종탑 우회로와 막다른 태엽창고를 빼고, 중앙 태엽 승강기를 타는 길이 왜 가장 빠른지 시간을 더해서 설명하도록 코칭합니다.',
      '느린 [C.서쪽종탑(우회로)]과 막다른 [X.태엽창고(막다른길)]를 배제하고, [D.아래층진자회랑]에서 [F.중앙태엽승강기]를 타면 8분 만에 꼭대기 종각에 도달할 수 있습니다.',
      '2026-11-27'
    ],
    [
      'T04',
      'III. 문제 해결과 프로그래밍 — [9정03-01]·[9정03-02]',
      '[MAP 4: 심연 최종 탈출] 아주 큰 미궁 단순화 및 불사신 사신 피하기',
      '고성 가장 깊은 곳의 아주 큰 미궁에서 [X1.고문실(막다른길)], [X2.폐쇄철문(통행불가)], [W.미혹마법진(함정)]을 가지치기하고, [C.승강장] 엘리베이터와 [G.심연대교]를 거쳐 8분 만에 대성문으로 탈출하는 가장 어려운 추상화 과제입니다.',
      '[1단계: 문제 정의 및 목표 분석] 탈출 출발 장소, 최종 탈출구, 그리고 최적 목표 시간(분)을 명시해 보세요.',
      '[2단계: 핵심 요소 추출 및 경로 배제 분석] 탈출에 꼭 필요한 핵심 장소 간 \'구간별 소요 시간\'과, 가지치기하여 탈출로에서 \'배제(소거)한 우회로 및 위험 요소의 구체적 배제 이유\'를 서술해 보세요.',
      '[3단계: 최적 탈출 경로 및 총 소요 시간 도출] 2단계의 핵심 구간을 종합하여 출발지부터 탈출구까지의 \'전체 최적 탈출 경로 순서\'와 \'총 소요 시간(합산 식)\'을 명확히 계산하여 적어 보세요.',
      '아주 큰 맵에서 필요 없는 길을 지우면 노선도처럼 단순해진다는 점을 느끼게 하고, 미혹마법진(W)에 걸리면 시간이 5분 늘어나는 이유를 설명하도록 코칭합니다.',
      '막다른 [X1.고문실(막다른길)], [X2.폐쇄철문(통행불가)], [W.미혹마법진(함정)]을 배제하고, [C.승강장] 엘리베이터로 하강 후 [G.심연대교]를 거쳐 8분 만에 탈출하세요.',
      '2026-11-27'
    ]
  ];

  if (taskSheet.getLastRow() > 1) {
    taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 10).clearContent();
  }
  taskSheet.getRange(2, 1, tasks.length, tasks[0].length).setValues(tasks);
  taskSheet.setColumnWidth(1, 70);
  taskSheet.setColumnWidth(2, 160);
  taskSheet.setColumnWidth(3, 240);
  taskSheet.setColumnWidth(4, 320);
  taskSheet.setColumnWidth(5, 220);
  taskSheet.setColumnWidth(6, 260);
  taskSheet.setColumnWidth(7, 240);
  taskSheet.setColumnWidth(8, 280);
  taskSheet.setColumnWidth(9, 260);
  taskSheet.setColumnWidth(10, 100);

  // 3. 고성탈출_형성평가 전용 제출 시트
  setupCastleEscapeSheet_(ss);

  // 4. 루브릭_평가기준표 시트
  setupRubricSheet_(ss);

  // 5. 대화기록 시트
  getOrCreateSheet_(ss, '대화기록', ['일시', '학생코드', '학생별명', '과제번호', '질문내용', 'AI답변', '모델'], '#4a148c');

  SpreadsheetApp.flush();
}

function setupCastleEscapeSheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, '고성탈출_형성평가', [
    '제출일시', '학생코드', '학생별명', '과제ID', '맵단계', '1단계: 문제 정의 및 목표 분석', 
    '2단계: 핵심 요소 추출 및 경로 배제 분석', '3단계: 최적 탈출 경로 및 총 소요 시간 도출', 
    '클리어시간(분)', '교사평가(루브릭 상/중/하)', '교사용 피드백 환류'
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
  if (sheet.getLastRow() === 1) {
    var samples = [
      [
        '2026-11-15 19:24:08', 'S01', '민트', 'T01', 'MAP 1: 용암 지하성당',
        '• 출발지: A.입구제단\n• 목표지: G.지하봉인문\n• 최적 목표 시간: 6분',
        '• 주요 핵심 구간별 소요 시간:\n  - [A.입구제단 ➔ B.청동계단]: 1분\n  - [B.청동계단 ➔ D.지하성소다리]: 2분\n  - [D.지하성소다리 ➔ F.흑요석룬대교]: 2분\n  - [F.흑요석룬대교 ➔ G.지하봉인문]: 1분\n• 우회로 및 위험 요소 배제 이유:\n  - [C.보물방(막다른길)]: 탈출로가 없는 막다른 방이므로 배제함\n  - [E.용암우회로(느림)]: 용암 위험과 7분이 걸리는 비효율적 우회로이므로 배제함',
        '• 전체 최적 탈출 경로:\nA.입구제단 ➔ B.청동계단 ➔ D.지하성소다리 ➔ F.흑요석룬대교 ➔ G.지하봉인문\n• 총 최단 탈출 시간 계산: 1분 + 2분 + 2분 + 1분 = 총 6분\n• 최종 탈출 전략: 막다른 보물방(C)과 용암 우회로(E)를 가지치기하여 도출한 유일한 6분 최단 경로임',
        '6', '상 (탁월)', '불필요한 우회로(E)와 막다른 길(C)을 완벽히 분별하여 6분 노선도처럼 단순하게 추상화함. 탁월함!'
      ],
      [
        '2026-11-15 19:35:12', 'S02', '레몬', 'T02', 'MAP 2: 빙하 수로 미궁',
        '• 출발지: A.빙하절벽입구\n• 목표지: G.얼음신전성문\n• 최적 목표 시간: 7분',
        '• 주요 핵심 구간별 소요 시간:\n  - [A.빙하절벽입구 ➔ B.수로갈림길]: 1분\n  - [B.수로갈림길 ➔ D.크리스탈수로]: 2분\n  - [D.크리스탈수로 ➔ E.크리스탈승강대]: 1분\n  - [E.크리스탈승강대 ➔ F.수정회랑]: 1분\n  - [F.수정회랑 ➔ G.얼음신전성문]: 2분\n• 우회로 및 위험 요소 배제 이유:\n  - [C.예티동굴(우회로)]: 설인이 지키고 5분이 지연되는 비효율적 우회로이므로 배제함\n  - [X.얼음감옥(막다른길)]: 낭떠러지로 연결되는 막다른 방이므로 배제함',
        '• 전체 최적 탈출 경로:\nA.빙하절벽입구 ➔ B.수로갈림길 ➔ D.크리스탈수로 ➔ E.크리스탈승강대 ➔ F.수정회랑 ➔ G.얼음신전성문\n• 총 최단 탈출 시간 계산: 1분 + 2분 + 1분 + 1분 + 2분 = 총 7분\n• 최종 탈출 전략: 예티동굴(C)와 얼음감옥(X)을 배제하고 수직 승강기(E)를 활용한 7분 최단 경로임',
        '7', '상 (탁월)', '빙하 절벽(C)과 동굴(X)을 배제하고 크리스탈 승강기를 핵심 거점으로 정확히 추출함. 우수함!'
      ]
    ];
    sheet.getRange(2, 1, samples.length, 11).setValues(samples);
  }
  return sheet;
}

function setupRubricSheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var rSheet = getOrCreateSheet_(ss, '루브릭_평가기준표', [
    '등급', '성취 수준 기준 (2022 개정 [9정03-01]·[9정03-02])', '세부 채점 지침 (핵심 경로 및 가지치기 분석)'
  ], '#283593');

  var rubrics = [
    [
      '상 (탁월)',
      '[9정03-01] 문제의 초기 상태(출발지)와 목표 상태(탈출구), 제한 조건을 완벽히 구조화하고, [9정03-02] 불필요한 우회로와 함정을 논리적으로 완전하게 제거(가지치기)하여 핵심 장소와 이동 시간을 바탕으로 가장 빠른 탈출 방법(알고리즘)을 명확히 만듦.',
      '1~3단계 전체 완성도 우수. [1단계: 상태 구조화] 완벽 명시. [2단계: 우회로 및 위험 요소 배제 사유] 논리적 서술. [3단계: 최적 경로 합산 식] 완벽 도출.'
    ],
    [
      '중 (보통)',
      '[9정03-01] 문제의 초기/목표 상태를 파악하였으나, [9정03-02] 일부 우회로의 구체적 배제 근거 서술이 다소 추상적이거나 최적 알고리즘의 총 소요 시간 계산 과정에 경미한 오차가 있음.',
      '최적 경로는 도출하였으나 우회로 배제 사유의 구체성이 다소 부족하거나 시간 합산에 사소한 오차 발생.'
    ],
    [
      '하 (노력요함)',
      '[9정03-01] 문제의 상태 구조화가 미흡하여 막다른 길이나 지연 함정을 핵심 경로에 포함하거나, [9정03-02] 직관에 의존하여 최적 탈출 알고리즘 도출을 완수하지 못함.',
      '출발지와 목표지의 단순 연결에 그치거나 시간 비교 없이 비효율적인 우회 경로를 탈출로로 오인함. 추가 튜터 코칭 필요.'
    ]
  ];

  if (rSheet.getLastRow() > 1) {
    rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 3).clearContent();
  }
  rSheet.getRange(2, 1, rubrics.length, rubrics[0].length).setValues(rubrics);
  rSheet.setColumnWidth(1, 90);
  rSheet.setColumnWidth(2, 360);
  rSheet.setColumnWidth(3, 480);
}

function enterStudent(rawCode) {
  ensureSetup_();
  var code = normalizeStudentCode_(rawCode);
  if (!code) {
    return { ok: false, message: '학생 코드를 입력해 주세요. (예: S01 ~ S20)' };
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
      message: '등록되지 않은 학생 코드입니다. S01 ~ S20 범위의 번호를 입력해 주세요.'
    };
  }

  var tasks = getCachedTasks_(ss);

  // 고성탈출_형성평가 시트에서 해당 학생 답안 읽기
  var submitSheet = ss.getSheetByName('고성탈출_형성평가');
  var submissionsByTask = {};
  if (submitSheet && submitSheet.getLastRow() > 1) {
    var submitData = submitSheet.getDataRange().getValues();
    for (var s = 1; s < submitData.length; s++) {
      var sCode = normalizeStudentCode_(submitData[s][1]);
      if (sCode === code) {
        var tId = String(submitData[s][3] || '').trim();
        var timeVal = submitData[s][0];
        var timeStr = timeVal instanceof Date
          ? Utilities.formatDate(timeVal, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
          : String(timeVal || '');
        submissionsByTask[tId] = {
          timestamp: timeStr,
          taskId: tId,
          stageNum: String(submitData[s][4] || ''),
          step1: String(submitData[s][5] || ''),
          step2: String(submitData[s][6] || ''),
          step3: String(submitData[s][7] || ''),
          clearMinutes: String(submitData[s][8] || ''),
          rubricScore: String(submitData[s][9] || ''),
          teacherFeedback: String(submitData[s][10] || '')
        };
      }
    }
  }

  // 대화기록 시트에서 해당 학생 기록 읽기
  var chatSheet = ss.getSheetByName('대화기록');
  var chats = [];
  if (chatSheet && chatSheet.getLastRow() > 1) {
    var chatData = chatSheet.getDataRange().getValues();
    for (var c = 1; c < chatData.length; c++) {
      var cCode = normalizeStudentCode_(chatData[c][1]);
      if (cCode === code) {
        var cTime = chatData[c][0];
        var cTimeStr = cTime instanceof Date
          ? Utilities.formatDate(cTime, 'Asia/Seoul', 'MM-dd HH:mm')
          : String(cTime || '');
        chats.push({
          timestamp: cTimeStr,
          taskId: String(chatData[c][3] || ''),
          question: String(chatData[c][4] || ''),
          answer: String(chatData[c][5] || '')
        });
      }
    }
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var sheetUrl = ss.getUrl();
  var submitSheetId = submitSheet ? submitSheet.getSheetId() : 0;

  return {
    ok: true,
    student: student,
    tasks: tasks,
    submissions: submissionsByTask,
    chats: chats,
    hasApiKey: !!apiKey,
    sheetUrl: sheetUrl,
    submitSheetId: submitSheetId
  };
}

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
    var submitSheet = setupCastleEscapeSheet_(ss);
    var submitData = submitSheet.getDataRange().getValues();
    var nowStr = formatNow_();
    var nickname = String((payload && payload.nickname) || '').trim();
    var stageNum = String((payload && payload.stageNum) || '').trim();
    var clearMinutes = String((payload && payload.clearMinutes) || '').trim();

    var clearMinutes = String((payload && payload.clearMinutes) || '').trim();
    if (!clearMinutes || clearMinutes === '0') {
      var defaultTargetMins = { 'T01': '6', 'T02': '7', 'T03': '8', 'T04': '8' };
      clearMinutes = defaultTargetMins[taskId] || '6';
    }

    var defaultFeedbacks = {
      'T01': '불필요한 우회로(E)와 막다른 길(C)을 완벽히 분별하여 노선도처럼 단순하게 추상화함. 탁월함!',
      'T02': '빙하 절벽(C)과 동굴(X)을 배제하고 크리스탈 승강기를 핵심 거점으로 정확히 추출함. 우수함!',
      'T03': '느린 서쪽 종탑 우회로를 빼고 중앙 태엽 승강기 경유 8분 최적 경로를 도출함. 우수함!',
      'T04': '아주 큰 심연 맵에서 미혹 함정(W)을 소거하고 대성문 8분 탈출 경로를 완벽히 도출함. 탁월함!'
    };

    var incomingRubric = String((payload && payload.rubric) || '').trim();
    var incomingFeedback = String((payload && payload.teacherFeedback) || '').trim();

    var existingRubric = '';
    var existingTeacherFeedback = '';
    var matchedRowIndices = [];
    for (var i = 1; i < submitData.length; i++) {
      if (
        normalizeStudentCode_(submitData[i][1]) === code &&
        String(submitData[i][3] || '').trim() === taskId
      ) {
        matchedRowIndices.push(i + 1); // 1-based sheet row index
        if (String(submitData[i][9] || '').trim() !== '') {
          existingRubric = String(submitData[i][9] || '').trim();
        }
        if (String(submitData[i][10] || '').trim() !== '') {
          existingTeacherFeedback = String(submitData[i][10] || '').trim();
        }
      }
    }

    var finalRubric = existingRubric || incomingRubric || '상 (탁월)';
    var finalFeedback = existingTeacherFeedback || incomingFeedback || defaultFeedbacks[taskId] || '핵심 요소를 완벽히 추출하고 최적 경로를 도출함.';

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
      finalRubric,
      finalFeedback
    ];

    if (matchedRowIndices.length > 0) {
      // 해당 학생과 과제의 기존 행에 최종 답안 덮어쓰기 (11개 열 전체)
      submitSheet.getRange(matchedRowIndices[0], 1, 1, 11).setValues([rowValues]);
      // 만약 과거 중복 행이 있었다면 아래쪽부터 삭제
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
      rubricScore: finalRubric,
      teacherFeedback: finalFeedback,
      submitSheetUrl: submitSheetUrl,
      message: '최종 답안이 구글 시트 [고성탈출_형성평가] 탭에 성공적으로 기록되었습니다!'
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

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
  if (!task) {
    return { ok: false, message: '과제 정보를 찾을 수 없습니다.' };
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var replyText = '';
  var modelName = 'Rule-based Practice Tutor';

  if (apiKey) {
    try {
      replyText = callGeminiApi_(apiKey, student, task, step1, step2, step3, question);
      modelName = 'gemini-1.5-flash';
    } catch (err) {
      replyText = buildSmartPracticeReply_(student, task, step1, step2, step3, question);
    }
  } else {
    replyText = buildSmartPracticeReply_(student, task, step1, step2, step3, question);
  }

  // 대화기록 시트에 저장
  var chatSheet = ss.getSheetByName('대화기록');
  if (chatSheet) {
    chatSheet.appendRow([
      formatNow_(),
      student.code,
      student.nickname,
      task.id,
      question,
      replyText,
      modelName
    ]);
  }

  return {
    ok: true,
    reply: replyText,
    model: modelName
  };
}

function findStudent_(ss, code) {
  var sheet = ss.getSheetByName('학생명단');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeStudentCode_(data[i][0]) === code) {
      return {
        code: normalizeStudentCode_(data[i][0]),
        nickname: String(data[i][1] || '').replace(/\(\d+번\)/g, '').trim(),
        level: String(data[i][2] || '기본'),
        instruction: String(data[i][3] || '')
      };
    }
  }
  return null;
}

function findTask_(ss, taskId) {
  var sheet = ss.getSheetByName('과제');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === taskId) {
      return {
        id: String(data[i][0]).trim(),
        unit: String(data[i][1] || ''),
        title: String(data[i][2] || ''),
        description: String(data[i][3] || ''),
        coachingGuide: String(data[i][7] || ''),
        practiceHint: String(data[i][8] || '')
      };
    }
  }
  return null;
}

function updateStudentStatus_(ss, code, newStatus) {
  var sheet = ss.getSheetByName('학생명단');
  var data = sheet.getDataRange().getValues();
  var norm = normalizeStudentCode_(code);
  for (var i = 1; i < data.length; i++) {
    if (normalizeStudentCode_(data[i][0]) === norm) {
      sheet.getRange(i + 1, 5).setValue(newStatus);
      break;
    }
  }
}

function buildSmartPracticeReply_(student, task, step1, step2, step3, question) {
  var q = question.toLowerCase();
  var opener = '[' + student.nickname + ' 학생을 위한 ' + student.level + ' 맞춤형 힌트]\n';
  var analysisLines = [];

  if (task.id === 'T01') {
    // MAP 1: 용암 지하성당
    if (q.indexOf('1단계') !== -1 || q.indexOf('목표') !== -1 || q.indexOf('출발') !== -1) {
      analysisLines.push('• [MAP 1 - 1단계 힌트] 출발 장소는 [A.입구제단], 목표 장소는 [G.지하봉인문]이며, 최적 목표 시간은 [6분]입니다.');
    } else if (q.indexOf('2단계') !== -1 || q.indexOf('배제') !== -1 || q.indexOf('우회') !== -1 || q.indexOf('가지치기') !== -1) {
      analysisLines.push('• [MAP 1 - 2단계 힌트] 막다른 방인 [C.보물방(막다른길)]과 용암 위험이 도사리고 7분이나 걸리는 [E.용암우회로(느림)]는 왜 탈출로에서 배제(가지치기)해야 할까요? 꼭 지나야 할 [B.청동계단(1분)], [D.지하성소다리(2분)], [F.흑요석룬대교(2분)]의 구간별 시간을 계산해 보세요.');
    } else if (q.indexOf('3단계') !== -1 || q.indexOf('최적') !== -1 || q.indexOf('계산') !== -1 || q.indexOf('경로') !== -1) {
      analysisLines.push('• [MAP 1 - 3단계 힌트] 전체 최적 경로 순서: A.입구제단 ➔ B.청동계단 ➔ D.지하성소다리 ➔ F.흑요석룬대교 ➔ G.지하봉인문 입니다. 총 최단 탈출 시간은 1분 + 2분 + 2분 + 1분 = 총 6분입니다.');
    } else {
      analysisLines.push('• [MAP 1 종합 코칭] ' + task.practiceHint);
    }
  } else if (task.id === 'T02') {
    // MAP 2: 빙하 수로 미궁
    if (q.indexOf('1단계') !== -1 || q.indexOf('목표') !== -1 || q.indexOf('출발') !== -1) {
      analysisLines.push('• [MAP 2 - 1단계 힌트] 출발 장소는 [A.빙하절벽입구], 목표 장소는 [G.얼음신전성문]이며, 최적 목표 시간은 [7분]입니다.');
    } else if (q.indexOf('2단계') !== -1 || q.indexOf('배제') !== -1 || q.indexOf('우회') !== -1 || q.indexOf('가지치기') !== -1) {
      analysisLines.push('• [MAP 2 - 2단계 힌트] 예티가 지키며 5분이나 걸리는 [C.예티동굴(우회로)]와 막다른 길인 [X.얼음감옥(막다른길)]은 왜 빼야 할까요? [D.크리스탈수로]에서 [E.크리스탈승강대]를 타면 어떤 점이 좋은지 구간별 시간을 계산해 보세요.');
    } else if (q.indexOf('3단계') !== -1 || q.indexOf('최적') !== -1 || q.indexOf('계산') !== -1 || q.indexOf('경로') !== -1) {
      analysisLines.push('• [MAP 2 - 3단계 힌트] 전체 최적 경로 순서: A.빙하절벽입구 ➔ B.수로갈림길 ➔ D.크리스탈수로 ➔ E.크리스탈승강대 ➔ F.수정회랑 ➔ G.얼음신전성문 입니다. 총 시간은 1분 + 2분 + 1분 + 1분 + 2분 = 총 7분입니다.');
    } else {
      analysisLines.push('• [MAP 2 종합 코칭] ' + task.practiceHint);
    }
  } else if (task.id === 'T03') {
    // MAP 3: 기계 시계탑
    if (q.indexOf('1단계') !== -1 || q.indexOf('목표') !== -1 || q.indexOf('출발') !== -1) {
      analysisLines.push('• [MAP 3 - 1단계 힌트] 출발 장소는 [A.지하태엽실], 목표 장소는 [H.시계탑정상종각]이며, 최적 목표 시간은 [8분]입니다.');
    } else if (q.indexOf('2단계') !== -1 || q.indexOf('배제') !== -1 || q.indexOf('우회') !== -1 || q.indexOf('가지치기') !== -1) {
      analysisLines.push('• [MAP 3 - 2단계 힌트] 4분이 더 걸리는 [C.서쪽종탑(우회로)]과 막다른 [X.태엽창고(막다른길)]는 왜 빼야 할까요? [D.아래층진자회랑]에서 [F.중앙태엽승강기]를 타고 올라가는 길의 시간을 계산해 보세요.');
    } else if (q.indexOf('3단계') !== -1 || q.indexOf('최적') !== -1 || q.indexOf('계산') !== -1 || q.indexOf('경로') !== -1) {
      analysisLines.push('• [MAP 3 - 3단계 힌트] 전체 최적 경로 순서: A.지하태엽실 ➔ B.제1기어교차로 ➔ D.아래층진자회랑 ➔ F.중앙태엽승강기 ➔ G.상층시계문자판 ➔ H.시계탑정상종각 입니다. 총 시간은 1분 + 2분 + 2분 + 2분 + 1분 = 총 8분입니다.');
    } else {
      analysisLines.push('• [MAP 3 종합 코칭] ' + task.practiceHint);
    }
  } else if (task.id === 'T04') {
    // MAP 4: 심연 최종 탈출
    if (q.indexOf('1단계') !== -1 || q.indexOf('목표') !== -1 || q.indexOf('출발') !== -1) {
      analysisLines.push('• [MAP 4 - 1단계 힌트] 출발 장소는 [A.첨탑옥좌], 최종 목표 장소는 [Z.심연대성문]이며, 최적 목표 시간은 [8분]입니다.');
    } else if (q.indexOf('2단계') !== -1 || q.indexOf('배제') !== -1 || q.indexOf('우회') !== -1 || q.indexOf('가지치기') !== -1) {
      analysisLines.push('• [MAP 4 - 2단계 힌트] 아주 큰 맵에서 막다른 [X1.고문실], [X2.폐쇄철문], 그리고 5분이 더 걸리는 함정 [W.미혹마법진]을 먼저 지우고, [C.승강장] 엘리베이터와 [G.심연대교]가 있는 길만 남겨 보세요.');
    } else if (q.indexOf('3단계') !== -1 || q.indexOf('최적') !== -1 || q.indexOf('계산') !== -1 || q.indexOf('경로') !== -1) {
      analysisLines.push('• [MAP 4 - 3단계 힌트] 전체 최적 경로 순서: A.첨탑옥좌 ➔ B.서쪽회랑 ➔ C.승강장 ➔ D.고문실갈림길 ➔ E.심연중앙수로 ➔ F.대교진입로 ➔ G.심연대교 ➔ Z.심연대성문 입니다. 총 소요 시간은 8분입니다.');
    } else {
      analysisLines.push('• [MAP 4 종합 코칭] ' + task.practiceHint);
    }
  }

  return opener + analysisLines.join('\n') + '\n\n(💡 팁: 화면 우측 3단계 형성평가 워크시트에 위 핵심 분석 내용을 정리하여 [구글 시트에 최종 답안 제출] 버튼을 누르세요.)';
}

function callGeminiApi_(apiKey, student, task, step1, step2, step3, question) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
  var prompt = [
    '너는 2022 개정 교육과정 중학교 정보 교과 [9정03-01] 문제 상태 구조화와 [9정03-02] 추상화·알고리즘 설계 단원의 친절한 AI 튜터다.',
    '학생 정보: ' + student.nickname + ' (' + student.level + ' 수준)',
    '맞춤 지도 지침: ' + student.instruction,
    '현재 과제: ' + task.title,
    '과제 설명: ' + task.description,
    '코칭 지침: ' + task.coachingGuide,
    '학생이 작성 중인 답안:',
    '- 1단계: ' + (step1 || '(미작성)'),
    '- 2단계: ' + (step2 || '(미작성)'),
    '- 3단계: ' + (step3 || '(미작성)'),
    '학생의 질문: ' + question,
    '지침: 정답을 곧바로 알려주지 말고, 불필요한 우회로나 막다른 길을 배제하고 핵심 경로를 추출하는 추상화 사고력을 기를 수 있도록 단계별 소크라테스식 힌트를 친절하게 제시하라.'
  ].join('\n');

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.7, maxOutputTokens: 600 }
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var json = JSON.parse(response.getContentText());
  if (json.candidates && json.candidates[0] && json.candidates[0].content) {
    return json.candidates[0].content.parts[0].text;
  }
  throw new Error('Gemini 응답 파싱 실패');
}

function formatNow_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}
