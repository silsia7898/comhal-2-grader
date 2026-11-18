// Exam rounds, initial statistics, rankings, and lecture catalog
window.APP_DATA = {
  EXAM_ROUNDS: [
    {
      id: 'guide-sample',
      name: '대표기출따라하기',
      fileName: '대표기출따라하기_문제.xlsx',
      description: '기본작업, 계산작업(함수), 분석작업(부분합/피벗), 기타작업(차트/매크로) 대표 유형'
    },
    {
      id: 'mock-01',
      name: '제1회 실전 모의고사',
      fileName: '제1회_실전모의고사_문제.xlsx',
      description: '최신 출제 경향 반영 모의고사 1회차'
    },
    {
      id: 'mock-02',
      name: '제2회 실전 모의고사',
      fileName: '제2회_실전모의고사_문제.xlsx',
      description: '최신 출제 경향 반영 모의고사 2회차'
    },
    {
      id: 'mock-03',
      name: '제3회 실전 모의고사',
      fileName: '제3회_실전모의고사_문제.xlsx',
      description: '최신 출제 경향 반영 모의고사 3회차'
    },
    {
      id: 'past-2025-1',
      name: '2025년 기출복원 1회',
      fileName: '2025_기출복원_1회_문제.xlsx',
      description: '상설시험 기출복원 1회차'
    },
    {
      id: 'past-2025-2',
      name: '2025년 기출복원 2회',
      fileName: '2025_기출복원_2회_문제.xlsx',
      description: '상설시험 기출복원 2회차'
    }
  ],

  INITIAL_STATS: {
    participantCount: 2103,
    averageScore: 39.01,
    distribution: [
      { range: '0~19점', count: 412 },
      { range: '20~39점', count: 685 },
      { range: '40~59점', count: 498 },
      { range: '60~69점', count: 214 },
      { range: '70~89점', count: 221 },
      { range: '90~100점', count: 73 }
    ]
  },

  INITIAL_STUDENT_RANKINGS: [
    { name: '오******', score: 2035 },
    { name: '황******', score: 1895 },
    { name: '소******', score: 1880 },
    { name: '이******', score: 1785 },
    { name: '김******', score: 1740 },
    { name: '박******', score: 1520 },
    { name: '강******', score: 1515 },
    { name: '신******', score: 1485 },
    { name: '홍******', score: 1455 },
    { name: '정******', score: 1390 }
  ],

  INITIAL_TEACHER_RANKINGS: [
    { name: 'T******', score: 765 },
    { name: '강******', score: 640 },
    { name: '박******', score: 345 },
    { name: '김******', score: 280 },
    { name: '이******', score: 245 }
  ],

  VIDEO_LECTURES: [
    {
      part: '문제1 · 기본작업 (20점)',
      title: '고급 필터 및 조건부 서식 핵심 정리',
      duration: '18분',
      summary: 'AND/OR 조건식 작성법, 수식을 이용한 조건부 서식 행 전체 적용 방법'
    },
    {
      part: '문제2 · 계산작업 (40점)',
      title: '필수 함수 15선 (IF, VLOOKUP, INDEX/MATCH, SUMIFS)',
      duration: '34분',
      summary: '배점 비중이 가장 높은 계산작업에서 자주 틀리는 절대참조($) 및 중첩 함수 풀이'
    },
    {
      part: '문제3 · 분석작업 (20점)',
      title: '부분합, 피벗 테이블, 데이터 통합 & 목표값 찾기',
      duration: '22분',
      summary: '정렬 선행 조건, 피벗 테이블 레이아웃 및 값 필드 설정 감점 방지 팁'
    },
    {
      part: '문제4 · 기타작업 (20점)',
      title: '차트 편집 및 매크로 기록/도형 연결 완벽 대비',
      duration: '19분',
      summary: '차트 제목/데이터 레이블/보조축 설정과 매크로 단추 지정 시 주의사항'
    }
  ],

  MOCK_QUESTIONS: [
    {
      id: 'q1',
      partKey: 'task1',
      partLabel: '문제1 · 기본작업',
      points: 20,
      type: 'choice',
      question: '고급 필터에서 "부서가 영업부이거나(OR) 기본급이 3,500,000 이상인" 데이터를 추출하려고 합니다. 조건 범위 작성 방법으로 올바른 것은?',
      options: [
        '부서(영업부)와 기본급(>=3500000)을 같은 행에 나란히 입력한다.',
        '부서(영업부)와 기본급(>=3500000)을 서로 다른 행에 입력한다.',
        '조건 필드 이름 아래에 =OR(영업부, >=3500000)을 입력한다.',
        '조건 범위를 비워두고 자동 필터를 적용한다.'
      ],
      answerIndex: 1,
      explanation: '고급 필터에서 AND 조건은 같은 행에, OR 조건(이거나)은 서로 다른 행에 입력해야 합니다.'
    },
    {
      id: 'q2',
      partKey: 'task2',
      partLabel: '문제2 · 계산작업 (1)',
      points: 20,
      type: 'formula',
      question: '아래 표에서 D2 셀에 필기(B2)와 실기(C2)의 평균이 70 이상이면 "합격", 아니면 "불합격"을 구하는 수식을 작성하세요.',
      tableHeaders: ['A (성명)', 'B (필기)', 'C (실기)', 'D (합격여부)'],
      tableRows: [['김민수', '80', '70', '?']],
      placeholder: '예: =IF(AVERAGE(B2:C2)>=70,"합격","불합격")',
      acceptedPatterns: [
        /^\s*=\s*IF\s*\(\s*AVERAGE\s*\(\s*B2\s*:\s*C2\s*\)\s*>=\s*70\s*,\s*["']합격["']\s*,\s*["']불합격["']\s*\)\s*$/i,
        /^\s*=\s*IF\s*\(\s*\(\s*B2\s*\+\s*C2\s*\)\s*\/\s*2\s*>=\s*70\s*,\s*["']합격["']\s*,\s*["']불합격["']\s*\)\s*$/i
      ],
      sampleAnswer: '=IF(AVERAGE(B2:C2)>=70,"합격","불합격")',
      explanation: 'AVERAGE(B2:C2)로 평균을 구한 뒤 IF 함수로 70 이상 여부를 판정합니다.'
    },
    {
      id: 'q3',
      partKey: 'task2',
      partLabel: '문제2 · 계산작업 (2)',
      points: 20,
      type: 'choice',
      question: '지정된 범위($B$2:$B$10)에서 "영업부"의 급여($C$2:$C$10) 합계를 구할 때 가장 알맞은 수식은?',
      options: [
        '=SUM($C$2:$C$10)',
        '=SUMIF($B$2:$B$10, "영업부", $C$2:$C$10)',
        '=COUNTIF($B$2:$B$10, "영업부")',
        '=VLOOKUP("영업부", $B$2:$C$10, 2, FALSE)'
      ],
      answerIndex: 1,
      explanation: '조건을 만족하는 셀의 합계를 구할 때는 SUMIF(조건범위, 조건, 합계범위) 함수를 사용합니다.'
    },
    {
      id: 'q4',
      partKey: 'task3',
      partLabel: '문제3 · 분석작업',
      points: 20,
      type: 'choice',
      question: '[부분합] 기능을 실행하기 전에 반드시 먼저 수행해야 하는 작업으로 옳은 것은?',
      options: [
        '그룹화할 기준 필드를 기준으로 정렬(오름차순 또는 내림차순)한다.',
        '모든 셀의 병합을 해제하고 테두리를 지운다.',
        '피벗 테이블을 먼저 생성한 뒤 부분합을 실행한다.',
        '데이터 유효성 검사를 설정한다.'
      ],
      answerIndex: 0,
      explanation: '부분합을 실행하기 전에는 반드시 그룹화할 항목(기준 필드)을 기준으로 정렬을 먼저 해야 올바르게 묶입니다.'
    },
    {
      id: 'q5',
      partKey: 'task4',
      partLabel: '문제4 · 기타작업',
      points: 20,
      type: 'choice',
      question: '매크로가 포함된 통합 문서를 저장할 때, 매크로가 정상적으로 유지되는 파일 형식(확장자)은?',
      options: [
        '.xlsx',
        '.csv',
        '.xlsm',
        '.txt'
      ],
      answerIndex: 2,
      explanation: '매크로가 포함된 엑셀 파일은 반드시 Excel 매크로 사용 통합 문서(.xlsm) 형식으로 저장해야 매크로가 사라지지 않습니다.'
    }
  ]
};
