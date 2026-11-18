// Excel workbook parser, grader, and sample workbook generator (uses SheetJS + JSZip)
window.ExcelGrader = (function () {
  function isXlsxReady() {
    return typeof window.XLSX !== 'undefined';
  }

  function findSheets(sheetNames, keyword) {
    return sheetNames.filter((name) => name.replace(/\s+/g, '').includes(keyword));
  }

  function collectSheetStats(sheet) {
    if (!sheet) return { cellCount: 0, formulas: [], errors: [], values: [] };
    const formulas = [];
    const errors = [];
    const values = [];
    let cellCount = 0;

    Object.keys(sheet).forEach((addr) => {
      if (addr.startsWith('!')) return;
      const cell = sheet[addr];
      if (!cell) return;
      cellCount += 1;
      if (cell.f) {
        formulas.push({ addr, formula: String(cell.f).toUpperCase(), value: cell.v });
      }
      if (cell.t === 'e' || (typeof cell.v === 'string' && /^#(REF|NAME|VALUE|DIV\/0|N\/A)!/.test(cell.v))) {
        errors.push({ addr, error: cell.w || cell.v || '#ERROR!' });
      }
      if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
        values.push({ addr, value: cell.v });
      }
    });

    return { cellCount, formulas, errors, values };
  }

  /**
   * Grade an uploaded Excel ArrayBuffer using SheetJS.
   * Returns { scores: { task1, task2, task3, task4, total }, deductions: [...] }
   */
  function gradeWorkbookBuffer(arrayBuffer, fileName) {
    if (!isXlsxReady()) {
      throw new Error('엑셀 분석 라이브러리(SheetJS)가 아직 로드되지 않았습니다. 인터넷 연결을 확인하거나 잠시 후 다시 시도하세요.');
    }

    const workbook = window.XLSX.read(arrayBuffer, {
      type: 'array',
      cellFormula: true,
      cellStyles: true,
      bookVBA: true
    });

    const sheetNames = workbook.SheetNames || [];
    const deductions = [];

    // 1. 기본작업 (20점)
    let task1 = 20;
    const basicSheets = findSheets(sheetNames, '기본');
    if (basicSheets.length === 0) {
      task1 = 0;
      deductions.push({
        problem: '문제1',
        task: '기본작업',
        deduct: -20,
        detail: '기본작업 시트(기본작업-1, 기본작업-2 등)를 찾을 수 없습니다.'
      });
    } else {
      const s1 = collectSheetStats(workbook.Sheets[basicSheets[0]]);
      if (s1.cellCount < 10) {
        task1 -= 10;
        deductions.push({
          problem: '문제1',
          task: '기본작업-1',
          deduct: -10,
          detail: `'${basicSheets[0]}' 시트의 입력 데이터 및 셀 서식이 기준보다 부족합니다.`
        });
      }
      if (basicSheets.length < 2) {
        task1 -= 5;
        deductions.push({
          problem: '문제1',
          task: '기본작업-2',
          deduct: -5,
          detail: '고급 필터/조건부 서식용 두 번째 기본작업 시트가 누락되었습니다.'
        });
      } else {
        const s2 = collectSheetStats(workbook.Sheets[basicSheets[1]]);
        if (s2.errors.length > 0) {
          task1 -= 5;
          deductions.push({
            problem: '문제1',
            task: '기본작업-2',
            deduct: -5,
            detail: `${s2.errors[0].addr} 셀에서 수식 오류(${s2.errors[0].error})가 발견되었습니다.`
          });
        }
      }
    }

    // 2. 계산작업 (40점 - 5문항 각 8점)
    let task2 = 40;
    const calcSheets = findSheets(sheetNames, '계산');
    if (calcSheets.length === 0) {
      task2 = 0;
      deductions.push({
        problem: '문제2',
        task: '계산작업',
        deduct: -40,
        detail: '계산작업 시트를 찾을 수 없습니다.'
      });
    } else {
      const calcStats = collectSheetStats(workbook.Sheets[calcSheets[0]]);
      const formulaCount = calcStats.formulas.length;

      if (calcStats.errors.length > 0) {
        const penalty = Math.min(16, calcStats.errors.length * 8);
        task2 -= penalty;
        deductions.push({
          problem: '문제2',
          task: '계산작업(수식오류)',
          deduct: -penalty,
          detail: `${calcStats.errors.map((e) => `${e.addr}(${e.error})`).slice(0, 3).join(', ')} 셀에 함수 오류가 있습니다.`
        });
      }

      if (formulaCount === 0) {
        task2 = 0;
        deductions.push({
          problem: '문제2',
          task: '계산작업(함수미사용)',
          deduct: -40,
          detail: '계산작업 시트에 수식(=함수)이 작성되어 있지 않습니다. 값만 입력하면 0점 처리됩니다.'
        });
      } else if (formulaCount < 5) {
        const missing = 5 - formulaCount;
        const penalty = Math.min(task2, missing * 8);
        if (penalty > 0) {
          task2 -= penalty;
          deductions.push({
            problem: '문제2',
            task: '계산작업(문항누락)',
            deduct: -penalty,
            detail: `5개 계산 문항 중 ${formulaCount}개 수식만 확인되었습니다. 미작성 문항에 대한 감점이 적용됩니다.`
          });
        }
      } else {
        // Check for hardcoded numbers instead of references or missing standard functions
        const hasLookupOrLogic = calcStats.formulas.some((f) =>
          /(IF|VLOOKUP|HLOOKUP|INDEX|MATCH|SUMIF|COUNTIF|RANK|AVERAGE|SUM)/.test(f.formula)
        );
        if (!hasLookupOrLogic) {
          task2 = Math.max(0, task2 - 8);
          deductions.push({
            problem: '문제2',
            task: '계산작업(함수유형)',
            deduct: -8,
            detail: '지정된 핵심 함수(IF, VLOOKUP, SUMIF 등) 대신 단순 연산식만 사용되었습니다.'
          });
        }
      }
    }

    // 3. 분석작업 (20점 - 2문항 각 10점)
    let task3 = 20;
    const analysisSheets = findSheets(sheetNames, '분석');
    if (analysisSheets.length === 0) {
      task3 = 0;
      deductions.push({
        problem: '문제3',
        task: '분석작업',
        deduct: -20,
        detail: '분석작업 시트(부분합, 피벗 테이블, 데이터 통합 등)를 찾을 수 없습니다.'
      });
    } else {
      const a1 = collectSheetStats(workbook.Sheets[analysisSheets[0]]);
      if (a1.cellCount < 8) {
        task3 -= 10;
        deductions.push({
          problem: '문제3',
          task: '분석작업-1',
          deduct: -10,
          detail: `'${analysisSheets[0]}' 시트의 분석 결과(부분합/정렬 등)가 불완전합니다.`
        });
      }
      if (analysisSheets.length < 2) {
        task3 -= 10;
        deductions.push({
          problem: '문제3',
          task: '분석작업-2',
          deduct: -10,
          detail: '두 번째 분석작업 시트가 존재하지 않습니다.'
        });
      } else {
        const a2 = collectSheetStats(workbook.Sheets[analysisSheets[1]]);
        if (a2.errors.length > 0) {
          task3 -= 10;
          deductions.push({
            problem: '문제3',
            task: '분석작업-2',
            deduct: -10,
            detail: `'${analysisSheets[1]}' 시트 ${a2.errors[0].addr} 셀에 오류가 있습니다.`
          });
        }
      }
    }

    // 4. 기타작업 (20점 - 차트/매크로 각 10점)
    let task4 = 20;
    const etcSheets = findSheets(sheetNames, '기타');
    if (etcSheets.length === 0) {
      task4 = 0;
      deductions.push({
        problem: '문제4',
        task: '기타작업',
        deduct: -20,
        detail: '기타작업 시트(차트 작업, 매크로 작업)를 찾을 수 없습니다.'
      });
    } else {
      const e1 = collectSheetStats(workbook.Sheets[etcSheets[0]]);
      if (e1.cellCount < 5) {
        task4 -= 10;
        deductions.push({
          problem: '문제4',
          task: '기타작업-1(차트)',
          deduct: -10,
          detail: '차트 원본 데이터 또는 차트 시트 구성이 누락되었습니다.'
        });
      }
      const isMacroEnabled = /\.xlsm$/i.test(fileName || '') || Boolean(workbook.vbaraw);
      if (etcSheets.length >= 2 && !isMacroEnabled) {
        task4 -= 5;
        deductions.push({
          problem: '문제4',
          task: '기타작업-2(매크로)',
          deduct: -5,
          detail: '매크로가 포함된 파일(.xlsm) 형식이 아니어서 매크로 모듈 확인이 제한되었습니다.'
        });
      }
    }

    task1 = Math.max(0, task1);
    task2 = Math.max(0, task2);
    task3 = Math.max(0, task3);
    task4 = Math.max(0, task4);

    const total = task1 + task2 + task3 + task4;

    if (deductions.length === 0) {
      deductions.push({
        problem: '전체',
        task: '감점 없음',
        deduct: 0,
        detail: '모든 작업 영역이 채점 기준과 일치합니다. (만점)'
      });
    }

    return {
      scores: { task1, task2, task3, task4, total },
      deductions
    };
  }

  /**
   * Build a realistic Computer Specialist Level 2 Excel workbook.
   * @param {Object} round
   * @param {boolean} asSolvedSample - if true, populates formulas including a couple realistic mistakes for demo
   */
  function createWorkbook(round, asSolvedSample = false) {
    const wb = window.XLSX.utils.book_new();

    // 기본작업-1
    const wsBasic1 = window.XLSX.utils.aoa_to_sheet([
      [`[${round.name}] 기본작업-1 : 거래처별 판매 현황`],
      ['거래처코드', '거래처명', '품목', '수량', '단가', '금액'],
      ['A-101', '한빛상사', '노트북', 12, 1200000, 14400000],
      ['A-102', '미래유통', '모니터', 25, 320000, 8000000],
      ['A-103', '성공테크', '키보드', 40, 45000, 1800000],
      ['A-104', '대한물산', '마우스', 55, 28000, 1540000]
    ]);
    window.XLSX.utils.book_append_sheet(wb, wsBasic1, '기본작업-1');

    // 기본작업-2
    const wsBasic2 = window.XLSX.utils.aoa_to_sheet([
      ['고급 필터 조건'],
      ['품목', '수량'],
      ['노트북', '>=10'],
      [''],
      ['결과 복사 위치'],
      ['거래처코드', '거래처명', '품목', '수량', '금액'],
      ['A-101', '한빛상사', '노트북', 12, 14400000]
    ]);
    window.XLSX.utils.book_append_sheet(wb, wsBasic2, '기본작업-2');

    // 계산작업
    const calcRows = [
      ['성명', '필기점수', '실기점수', '총점', '평균', '합격여부', '순위'],
      ['김민수', 82, 88, null, null, null, null],
      ['이서연', 64, 72, null, null, null, null],
      ['박지훈', 90, 95, null, null, null, null],
      ['최유진', 55, 60, null, null, null, null]
    ];
    const wsCalc = window.XLSX.utils.aoa_to_sheet(calcRows);
    if (asSolvedSample) {
      wsCalc['D2'] = { t: 'n', f: 'SUM(B2:C2)', v: 170 };
      wsCalc['E2'] = { t: 'n', f: 'AVERAGE(B2:C2)', v: 85 };
      wsCalc['F2'] = { t: 's', f: 'IF(E2>=70,"합격","불합격")', v: '합격' };
      wsCalc['G2'] = { t: 'n', f: 'RANK(E2,$E$2:$E$5)', v: 2 };
      // Simulate 1 formula error (#NAME?) in F4 for realistic demo deduction
      wsCalc['F4'] = { t: 'e', f: 'IFF(E4>=70,"합격","불합격")', v: '#NAME?', w: '#NAME?' };
      wsCalc['!ref'] = 'A1:G5';
    }
    window.XLSX.utils.book_append_sheet(wb, wsCalc, '계산작업');

    // 분석작업-1
    const wsAnalysis1 = window.XLSX.utils.aoa_to_sheet([
      ['부서', '직급', '기본급', '상여금'],
      ['영업부', '과장', 3800000, 500000],
      ['영업부', '대리', 3100000, 400000],
      ['기획부', '과장', 3900000, 550000],
      ['기획부', '사원', 2600000, 300000]
    ]);
    window.XLSX.utils.book_append_sheet(wb, wsAnalysis1, '분석작업-1');

    // 분석작업-2
    const wsAnalysis2 = window.XLSX.utils.aoa_to_sheet([
      ['피벗 테이블 요약'],
      ['행 레이블', '기본급 합계'],
      ['영업부', 6900000],
      ['기획부', 6500000]
    ]);
    window.XLSX.utils.book_append_sheet(wb, wsAnalysis2, '분석작업-2');

    // 기타작업-1 (차트)
    const wsEtc1 = window.XLSX.utils.aoa_to_sheet([
      ['분기', '매출액', '영업이익'],
      ['1분기', 1200, 240],
      ['2분기', 1450, 310],
      ['3분기', 1380, 290],
      ['4분기', 1620, 380]
    ]);
    window.XLSX.utils.book_append_sheet(wb, wsEtc1, '기타작업-1');

    // 기타작업-2 (매크로)
    const wsEtc2 = window.XLSX.utils.aoa_to_sheet([
      ['매크로 실행 대상'],
      ['항목', '값'],
      ['합계', 5650]
    ]);
    window.XLSX.utils.book_append_sheet(wb, wsEtc2, '기타작업-2');

    return wb;
  }

  function downloadRoundFile(round) {
    if (!isXlsxReady()) {
      throw new Error('SheetJS 라이브러리가 로드되지 않았습니다.');
    }
    const wb = createWorkbook(round, false);
    window.XLSX.writeFile(wb, round.fileName);
  }

  async function downloadAllRoundsZip(rounds) {
    if (!isXlsxReady() || typeof window.JSZip === 'undefined') {
      throw new Error('압축 다운로드 라이브러리(SheetJS/JSZip)가 로드되지 않았습니다.');
    }
    const zip = new window.JSZip();
    rounds.forEach((round) => {
      const wb = createWorkbook(round, false);
      const buffer = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      zip.file(round.fileName, buffer);
    });
    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '컴활2급_실기_전체문제.zip';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function createSampleSubmissionFile(round) {
    if (!isXlsxReady()) return null;
    const wb = createWorkbook(round, true);
    const array = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new File([array], `${round.name}_작성답안.xlsx`, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  return {
    gradeWorkbookBuffer,
    downloadRoundFile,
    downloadAllRoundsZip,
    createSampleSubmissionFile
  };
})();
