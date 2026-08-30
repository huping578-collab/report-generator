(() => {
  const state = {
    template: 'cq',
    running: false,
    desktop: false,
    progress: 0,
    missingTemplates: new Set(),
  };

  const templateButtons = [...document.querySelectorAll('.template-button')];
  const onlyCq = [...document.querySelectorAll('.only-cq')];
  const onlyGd = [...document.querySelectorAll('.only-gd')];
  const fields = [...document.querySelectorAll('.path-field')];
  const chooseButtons = [...document.querySelectorAll('.choose-button')];
  const runButton = document.getElementById('runButton');
  const openButton = document.getElementById('openButton');
  const toast = document.getElementById('toast');
  const logPanel = document.getElementById('logPanel');
  const progressBar = document.getElementById('progressBar');
  const progressTrack = document.querySelector('.progress-track');
  const progressValue = document.getElementById('progressValue');
  const progressLabel = document.getElementById('progressLabel');
  const runStatus = document.getElementById('runStatus');
  const stages = [...document.querySelectorAll('.run-stage')];
  const sideFlow = [...document.querySelectorAll('.flow-step')];
  const environmentStatus = document.getElementById('environmentStatus');

  const stageLabels = ['正在校验资料', '正在识别数据', '正在统计生成', '正在保存结果'];

  function desktopApi() {
    return window.pywebview && window.pywebview.api ? window.pywebview.api : null;
  }

  function showToast(message) {
    toast.textContent = message;
    toast.classList.add('is-visible');
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => toast.classList.remove('is-visible'), 2800);
  }

  function timeNow() {
    return new Intl.DateTimeFormat('zh-CN', {
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    }).format(new Date());
  }

  function addLog(message, type = '') {
    const line = document.createElement('div');
    const time = document.createElement('span');
    const content = document.createElement('span');
    line.className = `log-line ${type}`.trim();
    time.className = 'log-time';
    time.textContent = timeNow();
    content.className = 'log-message';
    content.textContent = String(message);
    line.append(time, content);
    logPanel.appendChild(line);
    logPanel.scrollTop = logPanel.scrollHeight;
  }

  function visibleRequiredFields() {
    return fields.filter((field) => {
      const row = field.closest('.file-row');
      return row && !row.hidden && field.dataset.optional !== 'true';
    });
  }

  function activeTemplateName() {
    return state.template === 'cq' ? '重庆项目报告模板' : '广东项目第五章模板';
  }

  function activeTemplateMissing() {
    // 重庆报告在真实模板缺失时会改用程序化生成；只有广东必须要求模板。
    return state.template === 'gd' && state.missingTemplates.has('广东项目第五章模板');
  }

  function updateActionAvailability() {
    runButton.disabled = state.running || !state.desktop || activeTemplateMissing();
    openButton.disabled = state.running || !state.desktop;
    chooseButtons.forEach((button) => { button.disabled = state.running || !state.desktop; });
    templateButtons.forEach((button) => { button.disabled = state.running; });
  }

  function updateConfigStatus() {
    const required = visibleRequiredFields();
    const filled = required.filter((field) => field.value.trim()).length;
    const status = document.getElementById('configStatus');
    status.classList.remove('is-warning', 'is-error');

    if (!state.desktop) {
      status.textContent = '正在连接桌面服务';
      status.classList.add('is-warning');
    } else if (activeTemplateMissing()) {
      status.textContent = '缺少内置 Word 模板';
      status.classList.add('is-error');
    } else if (filled === required.length) {
      status.textContent = '资料已就绪';
    } else {
      status.textContent = `已配置 ${filled}/${required.length} 项`;
      status.classList.add('is-warning');
    }
    updateActionAvailability();
  }

  function setTemplate(template) {
    if (state.running || !['cq', 'gd'].includes(template)) return;
    state.template = template;
    templateButtons.forEach((button) => {
      button.setAttribute('aria-pressed', String(button.dataset.template === template));
    });
    onlyCq.forEach((element) => { element.hidden = template !== 'cq'; });
    onlyGd.forEach((element) => { element.hidden = template !== 'gd'; });
    document.getElementById('pageTitle').textContent = template === 'cq'
      ? '重庆项目资料配置'
      : '广东项目资料配置';
    document.getElementById('pageDescription').textContent = template === 'cq'
      ? '选择检测资料所在位置。系统将生成统计工作簿和第五部分检测报告。'
      : '按地市扫描标线与护栏数据，生成评价报告、图表工作簿和人工复核对比。';
    document.getElementById('sourceHint').textContent = template === 'cq'
      ? '5 项输入'
      : '4 项必填，2 项可选，3 项阈值';
    addLog(`已切换至${template === 'cq' ? '重庆' : '广东'}项目模板`);
    updateConfigStatus();
  }

  function applyDetectedPaths(detected) {
    Object.entries(detected || {}).forEach(([key, value]) => {
      const field = document.getElementById(key);
      if (field && value) field.value = value;
    });
  }

  async function choosePath(button) {
    if (state.running) return;
    const api = desktopApi();
    if (!api) {
      addLog('当前是浏览器预览，无法访问本地文件选择器', 'error');
      showToast('请从桌面应用中选择本地文件');
      return;
    }

    const target = button.dataset.target;
    const field = document.getElementById(target);
    const isFile = /summary|manual|route/i.test(target);
    const previousText = button.textContent;
    button.disabled = true;
    button.textContent = isFile ? '正在选择…' : '正在识别…';

    try {
      const result = await api.choose_path(isFile ? 'file' : 'folder', target, state.template);
      if (!result || !result.ok) throw new Error(result?.error || '路径选择失败。');
      if (result.cancelled) return;
      field.value = result.path;
      applyDetectedPaths(result.detected);
      if (target === 'projectPath' && Object.keys(result.detected || {}).length > 1) {
        addLog('已选择项目文件夹并自动识别相关资料', 'ok');
      } else {
        addLog(`已选择：${result.path}`);
      }
      updateConfigStatus();
    } catch (error) {
      addLog(error.message, 'error');
      showToast(error.message);
    } finally {
      button.textContent = previousText;
      updateActionAvailability();
    }
  }

  function setProgress(value, label, stageIndex) {
    const safeValue = Math.max(state.progress, Math.min(100, Number(value) || 0));
    const safeStage = Math.max(0, Math.min(3, Number(stageIndex) || 0));
    state.progress = safeValue;
    progressBar.style.clipPath = `inset(0 ${100 - safeValue}% 0 0)`;
    progressTrack.setAttribute('aria-valuenow', String(safeValue));
    progressValue.textContent = `${safeValue}%`;
    progressLabel.textContent = label || stageLabels[safeStage];
    stages.forEach((stage, index) => {
      stage.classList.toggle('is-active', index === safeStage && safeValue < 100);
      stage.classList.toggle('is-done', index < safeStage || safeValue === 100);
    });
    sideFlow.forEach((step, index) => {
      step.classList.toggle('is-current', index === safeStage && safeValue < 100);
    });
  }

  function collectPayload() {
    const values = {};
    fields.forEach((field) => { values[field.id] = field.value.trim(); });
    ['markingThreshold', 'heightThreshold', 'boltThreshold'].forEach((id) => {
      values[id] = document.getElementById(id).value.trim();
    });
    return { template: state.template, values };
  }

  function setRunning(running) {
    state.running = running;
    runButton.textContent = running ? '正在生成…' : (state.progress === 100 ? '再次生成' : '开始生成');
    runStatus.classList.toggle('running', running);
    updateActionAvailability();
  }

  async function startRun() {
    if (state.running) return;
    const missing = visibleRequiredFields().filter((field) => !field.value.trim());
    if (missing.length) {
      const row = missing[0].closest('.file-row');
      const label = row.querySelector('label').textContent;
      addLog(`无法开始：请先选择${label}`, 'error');
      showToast(`请先选择${label}`);
      row.querySelector('.choose-button').focus();
      return;
    }

    if (state.template === 'gd') {
      const invalid = ['markingThreshold', 'heightThreshold', 'boltThreshold']
        .map((id) => document.getElementById(id))
        .find((input) => input.value === '' || Number(input.value) < 0 || Number(input.value) > 100);
      if (invalid) {
        addLog('一致性阈值必须在 0～100 之间', 'error');
        showToast('请检查一致性阈值');
        invalid.focus();
        return;
      }
    }

    const api = desktopApi();
    if (!api) {
      showToast('桌面服务未连接');
      return;
    }

    state.progress = 0;
    setProgress(0, '正在提交任务', 0);
    setRunning(true);
    addLog(`开始生成${state.template === 'cq' ? '重庆' : '广东'}项目报告`);
    try {
      const result = await api.start_run(collectPayload());
      if (!result || !result.ok) throw new Error(result?.error || '任务启动失败。');
    } catch (error) {
      setRunning(false);
      runStatus.textContent = 'FAILED';
      runStatus.classList.add('error');
      addLog(error.message, 'error');
      showToast(error.message);
    }
  }

  async function openOutput() {
    const path = document.getElementById('outputPath').value.trim();
    if (!path) {
      addLog('输出文件夹尚未配置', 'error');
      showToast('请先选择输出文件夹');
      return;
    }
    const api = desktopApi();
    if (!api) {
      showToast('桌面服务未连接');
      return;
    }
    const result = await api.open_output(path);
    if (!result || !result.ok) {
      const message = result?.error || '无法打开输出文件夹。';
      addLog(message, 'error');
      showToast(message);
      return;
    }
    addLog('已在资源管理器中打开输出文件夹', 'ok');
  }

  function handleDesktopEvent(detail) {
    if (!detail) return;
    if (detail.event === 'log') {
      setProgress(detail.progress, stageLabels[detail.stage], detail.stage);
      addLog(detail.message);
      return;
    }
    if (detail.event !== 'run') return;

    if (detail.status === 'running') {
      runStatus.textContent = 'RUNNING';
      runStatus.classList.remove('error');
      setProgress(detail.progress, stageLabels[detail.stage], detail.stage);
      addLog(detail.message);
    } else if (detail.status === 'complete') {
      setProgress(100, '生成完成', 3);
      setRunning(false);
      runStatus.textContent = 'COMPLETE';
      runStatus.classList.remove('running', 'error');
      addLog(detail.message, 'ok');
      showToast('报告生成完成，可打开输出文件夹');
    } else if (detail.status === 'error') {
      setRunning(false);
      runStatus.textContent = 'FAILED';
      runStatus.classList.remove('running');
      runStatus.classList.add('error');
      progressLabel.textContent = '生成失败';
      addLog(detail.message, 'error');
      showToast(detail.message);
    }
  }

  async function initializeDesktop() {
    const api = desktopApi();
    if (!api) return;
    try {
      const environment = await api.get_environment();
      state.desktop = true;
      state.missingTemplates.clear();
      Object.entries(environment.templates || {}).forEach(([name, info]) => {
        if (!info.exists) state.missingTemplates.add(name);
      });
      if (state.missingTemplates.has('广东项目第五章模板')) {
        environmentStatus.innerHTML = '数据仅在本机处理<br><strong>广东模板待补充</strong>';
        addLog('未检测到广东第五章模板：广东报告需要该模板，重庆报告可程序化生成。', 'error');
      } else {
        environmentStatus.innerHTML = '数据仅在本机处理<br>桌面服务已连接';
        addLog('桌面服务已连接，内置模板检查通过', 'ok');
      }
      updateConfigStatus();
    } catch (error) {
      environmentStatus.textContent = '桌面服务连接失败';
      addLog(`桌面服务连接失败：${error.message}`, 'error');
      updateConfigStatus();
    }
  }

  window.desktopEvents = handleDesktopEvent;
  window.addEventListener('pywebviewready', initializeDesktop);
  templateButtons.forEach((button) => {
    button.addEventListener('click', () => setTemplate(button.dataset.template));
  });
  chooseButtons.forEach((button) => {
    button.addEventListener('click', () => choosePath(button));
  });
  fields.forEach((field) => field.addEventListener('change', updateConfigStatus));
  runButton.addEventListener('click', startRun);
  openButton.addEventListener('click', openOutput);

  updateConfigStatus();
  setTimeout(() => {
    if (!state.desktop && !desktopApi()) {
      environmentStatus.innerHTML = '浏览器预览模式<br>本地文件功能仅在桌面版可用';
      addLog('当前为浏览器预览模式', 'error');
      updateConfigStatus();
    }
  }, 900);
})();
