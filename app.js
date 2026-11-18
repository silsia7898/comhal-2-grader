(function () {
  const STORAGE_KEY = 'comsilgi_grader_state_v1';
  const DATA = window.APP_DATA;

  // DOM Elements
  const noticeBanner = document.getElementById('noticeBanner');
  const closeNoticeBtn = document.getElementById('closeNoticeBtn');
  const roundSelect = document.getElementById('roundSelect');
  const downloadRoundBtn = document.getElementById('downloadRoundBtn');
  const downloadAllBtn = document.getElementById('downloadAllBtn');
  const teacherIdInput = document.getElementById('teacherIdInput');
  const studentNameInput = document.getElementById('studentNameInput');
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const dropzoneText = document.getElementById('dropzoneText');
  const loadSampleBtn = document.getElementById('loadSampleBtn');
  const clearFileBtn = document.getElementById('clearFileBtn');
  const queueCountEl = document.getElementById('queueCount');
  const startGradeBtn = document.getElementById('startGradeBtn');
  const printBtn = document.getElementById('printBtn');

  const scoreTask1El = document.getElementById('scoreTask1');
  const scoreTask2El = document.getElementById('scoreTask2');
  const scoreTask3El = document.getElementById('scoreTask3');
  const scoreTask4El = document.getElementById('scoreTask4');
  const totalScoreValueEl = document.getElementById('totalScoreValue');
  const passStatusBadge = document.getElementById('passStatusBadge');

  const participantCountEl = document.getElementById('participantCount');
  const averageScoreEl = document.getElementById('averageScore');
  const deductionTableBody = document.getElementById('deductionTableBody');

  const studentRankingChart = document.getElementById('studentRankingChart');
  const teacherRankingChart = document.getElementById('teacherRankingChart');

  const navHelp = document.getElementById('navHelp');
  const navMockExam = document.getElementById('navMockExam');
  const openMockExamBtn = document.getElementById('openMockExamBtn');
  const openStatsBtn = document.getElementById('openStatsBtn');
  const openVideoBtn = document.getElementById('openVideoBtn');

  const openLoginBtn = document.getElementById('openLoginBtn');
  const navUserWrap = document.getElementById('navUserWrap');
  const myHistoryBtn = document.getElementById('myHistoryBtn');
  const navUserLabel = document.getElementById('navUserLabel');
  const logoutBtn = document.getElementById('logoutBtn');

  const modalOverlay = document.getElementById('modalOverlay');
  const modalTitle = document.getElementById('modalTitle');
  const modalBody = document.getElementById('modalBody');
  const modalCloseBtn = document.getElementById('modalCloseBtn');

  const toggleTimerBtn = document.getElementById('toggleTimerBtn');
  const navbarTimerBadge = document.getElementById('navbarTimerBadge');
  const navbarTimerText = document.getElementById('navbarTimerText');

  const toastContainer = document.getElementById('toastContainer');

  // State
  let selectedFile = null;
  let timerInterval = null;
  let timerRemainingSec = 40 * 60;
  let timerRunning = false;
  let appState = loadPersistedState();

  function loadPersistedState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        return JSON.parse(raw);
      }
    } catch (_) {
      // ignore storage errors
    }
    return {
      stats: JSON.parse(JSON.stringify(DATA.INITIAL_STATS)),
      studentRankings: JSON.parse(JSON.stringify(DATA.INITIAL_STUDENT_RANKINGS)),
      teacherRankings: JSON.parse(JSON.stringify(DATA.INITIAL_TEACHER_RANKINGS))
    };
  }

  function savePersistedState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
    } catch (_) {
      // ignore storage errors
    }
  }

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.remove();
    }, 3200);
  }

  function initRounds() {
    roundSelect.innerHTML = '';
    DATA.EXAM_ROUNDS.forEach((round) => {
      const option = document.createElement('option');
      option.value = round.id;
      option.textContent = round.name;
      roundSelect.appendChild(option);
    });
  }

  function getSelectedRound() {
    const id = roundSelect.value;
    return DATA.EXAM_ROUNDS.find((r) => r.id === id) || DATA.EXAM_ROUNDS[0];
  }

  function renderStatsHeader() {
    participantCountEl.textContent = appState.stats.participantCount.toLocaleString();
    averageScoreEl.textContent = appState.stats.averageScore.toFixed(2);
  }

  function renderEmptyDeductionRows(minRows = 12) {
    deductionTableBody.innerHTML = '';
    for (let i = 0; i < minRows; i++) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td></td><td></td><td></td><td></td>';
      deductionTableBody.appendChild(tr);
    }
  }

  function renderDeductions(deductions) {
    deductionTableBody.innerHTML = '';
    deductions.forEach((item) => {
      const tr = document.createElement('tr');
      tr.className = 'has-data';
      const deductBadge =
        item.deduct < 0
          ? `<span class="deduct-pill">${item.deduct}</span>`
          : `<span class="ok-pill">0</span>`;

      tr.innerHTML = `
        <td class="cell-center">${escapeHtml(item.problem)}</td>
        <td>${escapeHtml(item.task)}</td>
        <td class="cell-center">${deductBadge}</td>
        <td>${escapeHtml(item.detail)}</td>
      `;
      deductionTableBody.appendChild(tr);
    });

    // Pad with empty zebra rows to keep the clean table appearance
    const remaining = Math.max(0, 11 - deductions.length);
    for (let i = 0; i < remaining; i++) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td></td><td></td><td></td><td></td>';
      deductionTableBody.appendChild(tr);
    }
  }

  function setTaskScoreBox(el, score, maxScore) {
    el.textContent = `${score} / ${maxScore}`;
    el.classList.remove('perfect', 'deducted');
    if (score === maxScore) {
      el.classList.add('perfect');
    } else {
      el.classList.add('deducted');
    }
  }

  function resetScoresUI() {
    [scoreTask1El, scoreTask2El, scoreTask3El, scoreTask4El].forEach((el) => {
      el.textContent = '-';
      el.classList.remove('perfect', 'deducted');
    });
    totalScoreValueEl.textContent = '0';
    passStatusBadge.classList.add('hidden');
    passStatusBadge.classList.remove('pass', 'fail');
    renderEmptyDeductionRows();
  }

  function updateSelectedFile(file) {
    if (!file) {
      selectedFile = null;
      dropzone.classList.remove('has-file');
      dropzoneText.textContent = '채점할 파일 선택';
      clearFileBtn.hidden = true;
      fileInput.value = '';
      return;
    }

    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      showToast('엑셀 파일(.xlsx, .xlsm, .xls)만 업로드할 수 있습니다.', 'error');
      return;
    }

    selectedFile = file;
    dropzone.classList.add('has-file');
    dropzoneText.textContent = file.name;
    clearFileBtn.hidden = false;
  }

  function maskName(raw) {
    const trimmed = (raw || '').trim();
    if (!trimmed) return '';
    if (trimmed.includes('*')) return trimmed;
    const first = trimmed.charAt(0);
    return `${first}******`;
  }

  function updateRankingsAndStats(totalScore) {
    // Update participant count & average
    const prevCount = appState.stats.participantCount;
    const prevTotalSum = prevCount * appState.stats.averageScore;
    const nextCount = prevCount + 1;
    const nextAvg = (prevTotalSum + totalScore) / nextCount;
    appState.stats.participantCount = nextCount;
    appState.stats.averageScore = Number(nextAvg.toFixed(2));

    // Update distribution bucket
    const dist = appState.stats.distribution;
    if (totalScore < 20) dist[0].count += 1;
    else if (totalScore < 40) dist[1].count += 1;
    else if (totalScore < 60) dist[2].count += 1;
    else if (totalScore < 70) dist[3].count += 1;
    else if (totalScore < 90) dist[4].count += 1;
    else dist[5].count += 1;

    // Update student ranking if nickname provided
    const studentName = maskName(studentNameInput.value);
    if (studentName) {
      const existing = appState.studentRankings.find((r) => r.name === studentName);
      if (existing) {
        existing.score += totalScore;
      } else {
        appState.studentRankings.push({ name: studentName, score: 1350 + totalScore });
      }
      appState.studentRankings.sort((a, b) => b.score - a.score);
      appState.studentRankings = appState.studentRankings.slice(0, 10);
    }

    // Update teacher ranking if teacher ID provided
    const teacherName = maskName(teacherIdInput.value);
    if (teacherName) {
      const existing = appState.teacherRankings.find((r) => r.name === teacherName);
      if (existing) {
        existing.score += Math.max(10, Math.round(totalScore / 2));
      } else {
        appState.teacherRankings.push({ name: teacherName, score: Math.max(50, totalScore) });
      }
      appState.teacherRankings.sort((a, b) => b.score - a.score);
      appState.teacherRankings = appState.teacherRankings.slice(0, 5);
    }

    savePersistedState();
    renderStatsHeader();
    renderRankings();
  }

  /**
   * Render clean horizontal bar chart in SVG matching reference image style
   */
  function renderHorizontalBarChart(container, items, minDomain, maxDomain, ticks) {
    const width = 520;
    const rowHeight = 25;
    const topPad = 14;
    const bottomPad = 30;
    const leftPad = 74;
    const rightPad = 24;
    const chartWidth = width - leftPad - rightPad;
    const height = topPad + items.length * rowHeight + bottomPad;

    const topColors = ['#b46e32', '#facc15', '#fbcfe8'];

    let svg = `<svg class="bar-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="누적점수 순위 차트">`;

    // Vertical gridlines + x-axis ticks
    ticks.forEach((tick) => {
      const ratio = (tick - minDomain) / (maxDomain - minDomain);
      const x = leftPad + ratio * chartWidth;
      svg += `<line x1="${x}" y1="${topPad - 4}" x2="${x}" y2="${height - bottomPad}" stroke="#e2e8f0" stroke-width="1" />`;
      svg += `<text x="${x}" y="${height - 10}" text-anchor="middle" font-size="10.5" fill="#64748b">${tick.toLocaleString()}</text>`;
    });

    // Left axis line
    svg += `<line x1="${leftPad}" y1="${topPad - 4}" x2="${leftPad}" y2="${height - bottomPad}" stroke="#94a3b8" stroke-width="1" />`;

    items.forEach((item, idx) => {
      const y = topPad + idx * rowHeight;
      const barH = Math.min(16, rowHeight - 8);
      const barY = y + (rowHeight - barH) / 2;
      const clamped = Math.max(minDomain, Math.min(maxDomain, item.score));
      const barW = Math.max(6, ((clamped - minDomain) / (maxDomain - minDomain)) * chartWidth);
      const fill = idx < 3 ? topColors[idx] : '#cbd5e1';

      svg += `<text x="${leftPad - 8}" y="${barY + barH - 3}" text-anchor="end" font-size="11" fill="#334155">${escapeHtml(item.name)}</text>`;
      svg += `<rect x="${leftPad}" y="${barY}" width="${barW}" height="${barH}" rx="2" fill="${fill}">
                <title>${escapeHtml(item.name)}: ${item.score.toLocaleString()}점</title>
              </rect>`;
    });

    svg += '</svg>';
    container.innerHTML = svg;
  }

  function renderRankings() {
    const sMax = Math.max(2200, ...appState.studentRankings.map((i) => i.score + 100));
    renderHorizontalBarChart(
      studentRankingChart,
      appState.studentRankings,
      1200,
      sMax,
      [1200, 1400, 1600, 1800, 2000, 2200]
    );

    const tMax = Math.max(800, ...appState.teacherRankings.map((i) => i.score + 50));
    renderHorizontalBarChart(
      teacherRankingChart,
      appState.teacherRankings,
      0,
      tMax,
      [0, 200, 400, 600, 800]
    );
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Modal Helpers
  function openModal(title, html) {
    modalTitle.textContent = title;
    modalBody.innerHTML = html;
    modalOverlay.classList.remove('hidden');
  }

  function closeModal() {
    modalOverlay.classList.add('hidden');
  }

  // Timer Logic (40분 실전 시험)
  function formatTimer(sec) {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  }

  function toggleExamTimer() {
    navbarTimerBadge.classList.remove('hidden');
    if (timerRunning) {
      clearInterval(timerInterval);
      timerRunning = false;
      toggleTimerBtn.classList.remove('running');
      showToast('실전 타이머가 일시정지되었습니다.');
      return;
    }

    if (timerRemainingSec <= 0) {
      timerRemainingSec = 40 * 60;
    }

    timerRunning = true;
    toggleTimerBtn.classList.add('running');
    navbarTimerText.textContent = formatTimer(timerRemainingSec);
    showToast('40분 실전 시험 타이머가 시작되었습니다.', 'success');

    timerInterval = setInterval(() => {
      timerRemainingSec -= 1;
      navbarTimerText.textContent = formatTimer(timerRemainingSec);
      if (timerRemainingSec <= 300) {
        navbarTimerBadge.classList.add('warning');
      }
      if (timerRemainingSec <= 0) {
        clearInterval(timerInterval);
        timerRunning = false;
        toggleTimerBtn.classList.remove('running');
        showToast('시험 시간 40분이 종료되었습니다!', 'error');
      }
    }, 1000);
  }

  // Event Listeners
  closeNoticeBtn.addEventListener('click', () => {
    noticeBanner.classList.add('hidden');
  });

  roundSelect.addEventListener('change', () => {
    resetScoresUI();
  });

  downloadRoundBtn.addEventListener('click', () => {
    const round = getSelectedRound();
    try {
      window.ExcelGrader.downloadRoundFile(round);
      showToast(`'${round.fileName}' 다운로드를 시작합니다.`, 'success');
    } catch (err) {
      showToast(err.message, 'error');
    }
  });

  downloadAllBtn.addEventListener('click', async () => {
    downloadAllBtn.disabled = true;
    try {
      await window.ExcelGrader.downloadAllRoundsZip(DATA.EXAM_ROUNDS);
      showToast('전체 모의고사 ZIP 파일 다운로드를 시작합니다.', 'success');
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      downloadAllBtn.disabled = false;
    }
  });

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      updateSelectedFile(e.target.files[0]);
    }
  });

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      updateSelectedFile(e.dataTransfer.files[0]);
    }
  });

  clearFileBtn.addEventListener('click', () => {
    updateSelectedFile(null);
    resetScoresUI();
  });

  loadSampleBtn.addEventListener('click', () => {
    const round = getSelectedRound();
    const sampleFile = window.ExcelGrader.createSampleSubmissionFile(round);
    if (!sampleFile) {
      showToast('엑셀 라이브러리가 준비되지 않았습니다. 잠시 후 다시 시도하세요.', 'error');
      return;
    }
    updateSelectedFile(sampleFile);
    showToast('샘플 답안 파일이 불러와졌습니다. [채점시작]을 눌러보세요.', 'success');
  });

  startGradeBtn.addEventListener('click', async () => {
    if (!selectedFile) {
      showToast('먼저 채점할 엑셀 파일(.xlsx/.xlsm)을 선택하거나 샘플 파일을 불러오세요.', 'error');
      return;
    }

    startGradeBtn.disabled = true;
    queueCountEl.textContent = '1';

    try {
      const buffer = await selectedFile.arrayBuffer();
      // Brief realistic processing delay for queue feedback
      await new Promise((resolve) => setTimeout(resolve, 450));

      const result = window.ExcelGrader.gradeWorkbookBuffer(buffer, selectedFile.name);
      applyGradeResultToDashboard(result.scores, result.deductions, `${getSelectedRound().name} (엑셀 채점)`);

      showToast(`채점 완료! 총점 ${result.scores.total}점입니다.`, 'success');
    } catch (err) {
      showToast(err.message || '파일을 분석하는 중 오류가 발생했습니다.', 'error');
    } finally {
      queueCountEl.textContent = '0';
      startGradeBtn.disabled = false;
    }
  });

  printBtn.addEventListener('click', () => {
    window.print();
  });

  toggleTimerBtn.addEventListener('click', toggleExamTimer);

  // Modals
  navHelp.addEventListener('click', () => {
    openModal(
      '채점프로그램 사용 도움말',
      `
        <ol class="help-list">
          <li><strong>회차 선택 및 문제 다운로드:</strong> 상단에서 풀고자 하는 모의고사 회차를 선택한 뒤 <code>해당회차 다운</code> 또는 <code>전체 다운 (ZIP)</code>으로 실습 엑셀 파일을 받으세요.</li>
          <li><strong>정보 입력 (선택):</strong> 선생님 ID와 학생 닉네임을 입력하면 채점 결과가 하단 누적점수 순위에 반영됩니다.</li>
          <li><strong>작성 파일 업로드:</strong> 작성이 완료된 <code>.xlsx</code> 또는 <code>.xlsm</code> 파일을 점선 영역에 끌어다 놓거나 클릭해 선택하세요.</li>
          <li><strong>채점 시작:</strong> <code>채점시작</code> 버튼을 누르면 기본작업(20점), 계산작업(40점), 분석작업(20점), 기타작업(20점) 점수와 감점 상세 내역이 표시됩니다.</li>
          <li><strong>실전 타이머:</strong> <code>문제 유형 선택</code> 우측 시계 아이콘을 누르면 40분 실전 시험 타이머가 작동합니다.</li>
        </ol>
      `
    );
  });

  openStatsBtn.addEventListener('click', () => {
    const dist = appState.stats.distribution;
    const total = dist.reduce((acc, d) => acc + d.count, 0) || 1;
    const passCount = dist[4].count + dist[5].count;
    const passRate = ((passCount / total) * 100).toFixed(1);

    const barsHtml = dist
      .map((item) => {
        const pct = Math.round((item.count / total) * 100);
        return `
          <div style="margin-bottom:10px;">
            <div style="display:flex; justify-content:space-between; font-size:12.5px; margin-bottom:3px;">
              <strong>${item.range}</strong>
              <span>${item.count.toLocaleString()}명 (${pct}%)</span>
            </div>
            <div style="background:#e2e8f0; border-radius:999px; height:10px; overflow:hidden;">
              <div style="width:${Math.max(3, pct)}%; height:100%; background:${
                item.range.startsWith('70') || item.range.startsWith('90') ? '#22c55e' : '#3b82f6'
              };"></div>
            </div>
          </div>
        `;
      })
      .join('');

    openModal(
      '점수별 통계보기',
      `
        <div class="stats-summary-grid">
          <div class="stats-kpi">
            <div class="stats-kpi-label">누적 응시 인원</div>
            <div class="stats-kpi-value">${appState.stats.participantCount.toLocaleString()}명</div>
          </div>
          <div class="stats-kpi">
            <div class="stats-kpi-label">전체 평균 점수</div>
            <div class="stats-kpi-value">${appState.stats.averageScore.toFixed(2)}점</div>
          </div>
          <div class="stats-kpi">
            <div class="stats-kpi-label">합격률 (70점↑)</div>
            <div class="stats-kpi-value" style="color:#16a34a;">${passRate}%</div>
          </div>
        </div>
        <div>${barsHtml}</div>
      `
    );
  });

  openVideoBtn.addEventListener('click', () => {
    const listHtml = DATA.VIDEO_LECTURES.map(
      (lec) => `
        <div class="lecture-item">
          <div>
            <div style="font-size:11.5px; font-weight:700; color:#dc2626; margin-bottom:2px;">${escapeHtml(lec.part)} · ${escapeHtml(lec.duration)}</div>
            <h4>${escapeHtml(lec.title)}</h4>
            <p>${escapeHtml(lec.summary)}</p>
          </div>
          <button type="button" class="btn btn-outline" onclick="alert('${escapeHtml(lec.title)}\\n\\n핵심 포인트: ${escapeHtml(lec.summary)}')">학습하기</button>
        </div>
      `
    ).join('');

    openModal('동영상 강좌 및 영역별 핵심 가이드', `<div class="lecture-list">${listHtml}</div>`);
  });

  modalCloseBtn.addEventListener('click', closeModal);
  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) closeModal();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !modalOverlay.classList.contains('hidden')) {
      closeModal();
    }
  });

  // =========================================================
  // 대시보드 점수 반영 + 로그인 + 모의문제 풀기 기능
  // =========================================================
  const USERS_KEY = 'comsilgi_users_v1';
  const SESSION_KEY = 'comsilgi_session_v1';
  let currentUser = loadSession();

  function loadUsers() {
    try {
      return JSON.parse(localStorage.getItem(USERS_KEY) || '{}');
    } catch (_) {
      return {};
    }
  }

  function saveUsers(users) {
    try {
      localStorage.setItem(USERS_KEY, JSON.stringify(users));
    } catch (_) {}
  }

  function loadSession() {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (_) {
      return null;
    }
  }

  function saveSession(user) {
    currentUser = user;
    try {
      if (user) localStorage.setItem(SESSION_KEY, JSON.stringify(user));
      else localStorage.removeItem(SESSION_KEY);
    } catch (_) {}
    renderAuthUI();
  }

  function renderAuthUI() {
    if (currentUser) {
      openLoginBtn.classList.add('hidden');
      navUserWrap.classList.remove('hidden');
      navUserLabel.textContent = `👤 ${currentUser.username} (${currentUser.role})`;
      if (!studentNameInput.value) studentNameInput.value = currentUser.username;
      if (currentUser.teacherId && !teacherIdInput.value) teacherIdInput.value = currentUser.teacherId;
    } else {
      openLoginBtn.classList.remove('hidden');
      navUserWrap.classList.add('hidden');
    }
  }

  function recordUserHistory(sourceLabel, totalScore) {
    if (!currentUser) return;
    const users = loadUsers();
    const record = {
      date: new Date().toLocaleString('ko-KR'),
      source: sourceLabel,
      score: totalScore
    };
    if (!users[currentUser.username]) {
      users[currentUser.username] = { ...currentUser, history: [] };
    }
    users[currentUser.username].history = [record, ...(users[currentUser.username].history || [])].slice(0, 20);
    saveUsers(users);
  }

  function applyGradeResultToDashboard(scores, deductions, sourceLabel) {
    setTaskScoreBox(scoreTask1El, scores.task1, 20);
    setTaskScoreBox(scoreTask2El, scores.task2, 40);
    setTaskScoreBox(scoreTask3El, scores.task3, 20);
    setTaskScoreBox(scoreTask4El, scores.task4, 20);

    totalScoreValueEl.textContent = scores.total;
    passStatusBadge.classList.remove('hidden', 'pass', 'fail');
    if (scores.total >= 70) {
      passStatusBadge.textContent = '합격 기준 충족 (70점 이상)';
      passStatusBadge.classList.add('pass');
    } else {
      passStatusBadge.textContent = '불합격 (70점 미만 · 보완 필요)';
      passStatusBadge.classList.add('fail');
    }

    renderDeductions(deductions);
    updateRankingsAndStats(scores.total);
    recordUserHistory(sourceLabel, scores.total);
  }

  function openAuthModal(initialMode = 'login') {
    let mode = initialMode;
    const render = () => {
      openModal(
        mode === 'login' ? '로그인' : '회원가입',
        `
          <div class="auth-tabs">
            <button type="button" class="auth-tab ${mode === 'login' ? 'active' : ''}" id="tabLogin">로그인</button>
            <button type="button" class="auth-tab ${mode === 'signup' ? 'active' : ''}" id="tabSignup">회원가입</button>
          </div>
          <form class="auth-form" id="authForm">
            <div class="auth-field">
              <label for="authUsername">아이디 (닉네임)</label>
              <input type="text" id="authUsername" class="text-input" placeholder="예: 홍길동" required maxlength="20" />
            </div>
            <div class="auth-field">
              <label for="authPassword">비밀번호</label>
              <input type="password" id="authPassword" class="text-input" placeholder="비밀번호 입력" required maxlength="30" />
            </div>
            ${
              mode === 'signup'
                ? `
              <div class="auth-field">
                <label for="authRole">구분</label>
                <select id="authRole" class="select-input">
                  <option value="학생">학생</option>
                  <option value="선생님">선생님</option>
                </select>
              </div>
              <div class="auth-field">
                <label for="authTeacherId">담당 선생님 ID (선택)</label>
                <input type="text" id="authTeacherId" class="text-input" placeholder="예: T001" maxlength="20" />
              </div>
            `
                : ''
            }
            <div id="authErrorMsg" style="color:#dc2626; font-size:12.5px; font-weight:600;"></div>
            <button type="submit" class="btn btn-primary" style="height:40px;">
              ${mode === 'login' ? '로그인하기' : '회원가입 완료'}
            </button>
            <button type="button" class="btn btn-outline" id="demoLoginBtn">
              ⚡ 체험 계정(학생)으로 바로 로그인
            </button>
          </form>
        `
      );

      document.getElementById('tabLogin').addEventListener('click', () => {
        mode = 'login';
        render();
      });
      document.getElementById('tabSignup').addEventListener('click', () => {
        mode = 'signup';
        render();
      });

      document.getElementById('demoLoginBtn').addEventListener('click', () => {
        const users = loadUsers();
        if (!users['체험학생']) {
          users['체험학생'] = { username: '체험학생', password: '123', role: '학생', teacherId: 'T001', history: [] };
          saveUsers(users);
        }
        saveSession({ username: '체험학생', role: '학생', teacherId: 'T001' });
        closeModal();
        showToast('체험학생 계정으로 로그인되었습니다.', 'success');
      });

      document.getElementById('authForm').addEventListener('submit', (e) => {
        e.preventDefault();
        const username = document.getElementById('authUsername').value.trim();
        const password = document.getElementById('authPassword').value;
        const errEl = document.getElementById('authErrorMsg');
        const users = loadUsers();

        if (mode === 'signup') {
          if (users[username]) {
            errEl.textContent = '이미 존재하는 아이디입니다.';
            return;
          }
          const role = document.getElementById('authRole').value;
          const teacherId = document.getElementById('authTeacherId').value.trim();
          users[username] = { username, password, role, teacherId, history: [] };
          saveUsers(users);
          saveSession({ username, role, teacherId });
          closeModal();
          showToast(`${username}님, 회원가입 및 로그인되었습니다!`, 'success');
        } else {
          const found = users[username];
          if (!found || found.password !== password) {
            errEl.textContent = '아이디 또는 비밀번호가 일치하지 않습니다.';
            return;
          }
          saveSession({ username: found.username, role: found.role, teacherId: found.teacherId || '' });
          closeModal();
          showToast(`${found.username}님, 환영합니다!`, 'success');
        }
      });
    };
    render();
  }

  function openHistoryModal() {
    if (!currentUser) return;
    const users = loadUsers();
    const history = (users[currentUser.username] && users[currentUser.username].history) || [];
    const listHtml =
      history.length === 0
        ? '<p style="color:#64748b; text-align:center; padding:20px 0;">아직 채점 기록이 없습니다. 파일을 채점하거나 모의문제를 풀어보세요!</p>'
        : `<table class="quiz-mini-table">
            <thead><tr><th>일시</th><th>유형/회차</th><th>점수</th></tr></thead>
            <tbody>
              ${history
                .map(
                  (h) => `<tr><td>${escapeHtml(h.date)}</td><td>${escapeHtml(h.source)}</td><td><strong>${h.score}점</strong></td></tr>`
                )
                .join('')}
            </tbody>
          </table>`;

    openModal(`${currentUser.username}님의 채점 기록`, listHtml);
  }

  function openMockExamModal() {
    const questions = DATA.MOCK_QUESTIONS || [];
    const qHtml = questions
      .map((q, idx) => {
        const tableHtml = q.tableHeaders
          ? `<table class="quiz-mini-table">
              <thead><tr>${q.tableHeaders.map((h) => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>
              <tbody>${q.tableRows.map((row) => `<tr>${row.map((c) => `<td>${escapeHtml(c)}</td>`).join('')}</tr>`).join('')}</tbody>
            </table>`
          : '';

        const inputHtml =
          q.type === 'choice'
            ? `<div class="quiz-options">
                ${q.options
                  .map(
                    (opt, oIdx) => `
                  <label class="quiz-option-label">
                    <input type="radio" name="mock_${q.id}" value="${oIdx}" />
                    <span>${escapeHtml(opt)}</span>
                  </label>
                `
                  )
                  .join('')}
              </div>`
            : `<input type="text" id="mock_input_${q.id}" class="text-input" placeholder="${escapeHtml(q.placeholder || '수식을 입력하세요')}" />`;

        return `
          <div class="quiz-question-card" id="qcard_${q.id}">
            <div class="quiz-question-head">
              <span class="quiz-part-tag">${escapeHtml(q.partLabel)}</span>
              <span class="quiz-points-tag">${q.points}점</span>
            </div>
            <p class="quiz-question-title">Q${idx + 1}. ${escapeHtml(q.question)}</p>
            ${tableHtml}
            ${inputHtml}
            <div class="quiz-feedback hidden" id="qfb_${q.id}"></div>
          </div>
        `;
      })
      .join('');

    openModal(
      `모의문제 풀기 (${getSelectedRound().name})`,
      `
        <p style="margin-top:0; color:#475569; font-size:13px;">
          컴퓨터활용능력 2급 실기 핵심 문제를 직접 풀고 바로 채점해 보세요. 결과는 메인 점수판에도 반영됩니다!
        </p>
        <div id="mockQuizContainer">${qHtml}</div>
        <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;">
          <button type="button" class="btn btn-primary" id="submitMockExamBtn">✓ 모의문제 채점하기</button>
        </div>
      `
    );

    document.getElementById('submitMockExamBtn').addEventListener('click', () => {
      const scores = { task1: 0, task2: 0, task3: 0, task4: 0, total: 0 };
      const deductions = [];

      questions.forEach((q, idx) => {
        const fbEl = document.getElementById(`qfb_${q.id}`);
        let isCorrect = false;

        if (q.type === 'choice') {
          const checked = document.querySelector(`input[name="mock_${q.id}"]:checked`);
          const val = checked ? Number(checked.value) : -1;
          isCorrect = val === q.answerIndex;
        } else {
          const raw = (document.getElementById(`mock_input_${q.id}`).value || '').trim();
          isCorrect = q.acceptedPatterns.some((re) => re.test(raw));
        }

        fbEl.classList.remove('hidden', 'correct', 'wrong');
        if (isCorrect) {
          scores[q.partKey] += q.points;
          fbEl.classList.add('correct');
          fbEl.textContent = `✓ 정답! (+${q.points}점) — ${q.explanation}`;
        } else {
          fbEl.classList.add('wrong');
          const correctText = q.type === 'choice' ? q.options[q.answerIndex] : q.sampleAnswer;
          fbEl.textContent = `✗ 오답 (-${q.points}점) | 정답: ${correctText} — ${q.explanation}`;
          deductions.push({
            problem: `문항${idx + 1}`,
            task: q.partLabel,
            deduct: -q.points,
            detail: `${q.explanation} (정답: ${correctText})`
          });
        }
      });

      scores.total = scores.task1 + scores.task2 + scores.task3 + scores.task4;
      if (deductions.length === 0) {
        deductions.push({
          problem: '전체',
          task: '감점 없음',
          deduct: 0,
          detail: '모의문제 5문항을 모두 맞혔습니다! (100점 만점)'
        });
      }

      applyGradeResultToDashboard(scores, deductions, `${getSelectedRound().name} (모의문제)`);
      showToast(`모의문제 채점 완료! 총점 ${scores.total}점입니다.`, scores.total >= 70 ? 'success' : 'info');
    });
  }

  openLoginBtn.addEventListener('click', () => openAuthModal('login'));
  logoutBtn.addEventListener('click', () => {
    saveSession(null);
    showToast('로그아웃되었습니다.');
  });
  myHistoryBtn.addEventListener('click', openHistoryModal);
  navMockExam.addEventListener('click', openMockExamModal);
  openMockExamBtn.addEventListener('click', openMockExamModal);

  // Initial render
  initRounds();
  renderStatsHeader();
  renderEmptyDeductionRows();
  renderRankings();
  renderAuthUI();
})();
