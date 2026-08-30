import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const ROOT = path.resolve(import.meta.dirname, '..');
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const PORT = 9337;
const PROFILE = path.join(process.env.LOCALAPPDATA || 'C:/FakeD/Temp', 'Temp', 'report-ui-cdp-profile');
const SHOTS = path.join(ROOT, 'artifacts', 'screenshots');
const APP_URL = pathToFileURL(path.join(ROOT, 'frontend', 'index.html')).href;

fs.rmSync(PROFILE, { recursive: true, force: true });
fs.mkdirSync(SHOTS, { recursive: true });

const results = [];
const runtimeErrors = [];
function report(name, ok, detail = '') {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${detail ? ` | ${detail}` : ''}`);
}
function sleep(ms) { return new Promise((resolve) => setTimeout(resolve, ms)); }

class CDP {
  constructor(ws) {
    this.ws = ws;
    this.id = 0;
    this.pending = new Map();
    ws.onmessage = (event) => {
      const message = JSON.parse(event.data);
      if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params.exceptionDetails.text);
      if (message.id && this.pending.has(message.id)) {
        const pending = this.pending.get(message.id);
        this.pending.delete(message.id);
        message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result);
      }
    };
  }
  send(method, params = {}) {
    return new Promise((resolve, reject) => {
      const id = ++this.id;
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify({ id, method, params }));
    });
  }
}

let cdp;
async function evaluate(expression) {
  const result = await cdp.send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.text);
  return result.result.value;
}
async function screenshot(name) {
  const result = await cdp.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false });
  const output = path.join(SHOTS, `${name}.png`);
  fs.writeFileSync(output, Buffer.from(result.data, 'base64'));
  console.log(`SHOT ${output}`);
}
async function waitLoaded() {
  for (let i = 0; i < 60; i += 1) {
    try {
      if (await evaluate("document.readyState === 'complete' && !!document.querySelector('.app-shell')")) return true;
    } catch {}
    await sleep(200);
  }
  return false;
}

async function verify() {
  await sleep(1100);
  await evaluate(`document.activeElement && document.activeElement.blur(); true`);
  for (let i = 0; i < 2; i += 1) {
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 });
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 });
  }
  const desktop = await evaluate(`(() => {
    const shell = document.querySelector('.app-shell');
    const focused = document.activeElement;
    const fieldLabels = [...document.querySelectorAll('.path-field')].every((input) =>
      !!document.querySelector('label[for="' + input.id + '"]'));
    return {
      title: document.title,
      cqVisible: [...document.querySelectorAll('.only-cq')].some((el) => !el.hidden),
      gdHidden: [...document.querySelectorAll('.only-gd')].every((el) => el.hidden),
      labels: fieldLabels,
      noOverflow: document.documentElement.scrollWidth <= document.documentElement.clientWidth,
      columns: getComputedStyle(shell).gridTemplateColumns,
      focusOutline: getComputedStyle(focused).outlineStyle,
      focusedClass: focused.className,
      previewStatus: document.getElementById('environmentStatus').textContent,
      runDisabled: document.getElementById('runButton').disabled,
    };
  })()`);
  report('desktop title', desktop.title === '报告生成工具', desktop.title);
  report('default Chongqing surface', desktop.cqVisible && desktop.gdHidden, JSON.stringify(desktop));
  report('all path inputs labelled', desktop.labels);
  report('desktop has no horizontal overflow', desktop.noOverflow, desktop.columns);
  report('keyboard focus is visible', desktop.focusOutline !== 'none', desktop.focusOutline);
  report('browser fallback is explicit', desktop.previewStatus.includes('浏览器预览模式'));
  report('native-only action disabled in browser', desktop.runDisabled);
  await screenshot('desktop-cq');

  const guangdong = await evaluate(`(() => {
    document.querySelector('[data-template="gd"]').click();
    const threshold = document.getElementById('markingThreshold');
    return {
      title: document.getElementById('pageTitle').textContent,
      thresholdVisible: !threshold.closest('.only-gd').hidden,
      hint: document.getElementById('sourceHint').textContent,
      markingOptional: document.getElementById('markingPath').dataset.optional,
    };
  })()`);
  report('Guangdong template switches', guangdong.title.includes('广东') && guangdong.thresholdVisible, JSON.stringify(guangdong));
  report('optional specialized folders preserved', guangdong.markingOptional === 'true' && guangdong.hint.includes('2 项可选'));

  await cdp.send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] });
  const reduced = await evaluate(`getComputedStyle(document.querySelector('.button')).transitionDuration`);
  report('reduced motion honored', reduced === '1e-05s' || reduced === '0.00001s' || reduced === '0s', reduced);

  await cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 375,
    height: 812,
    deviceScaleFactor: 1,
    mobile: true,
  });
  await sleep(250);
  const mobile = await evaluate(`(() => {
    const shell = document.querySelector('.app-shell');
    const fileRow = document.querySelector('.file-row');
    const button = document.querySelector('.button');
    return {
      noOverflow: document.documentElement.scrollWidth <= document.documentElement.clientWidth,
      shellColumns: getComputedStyle(shell).gridTemplateColumns,
      fileColumns: getComputedStyle(fileRow).gridTemplateColumns,
      buttonHeight: button.getBoundingClientRect().height,
      viewport: [innerWidth, innerHeight],
    };
  })()`);
  report('mobile has no horizontal overflow', mobile.noOverflow, JSON.stringify(mobile));
  report('mobile controls meet 44px target', mobile.buttonHeight >= 44, String(mobile.buttonHeight));
  report('mobile collapses to one column', !mobile.shellColumns.includes('232px'), mobile.shellColumns);
  await screenshot('mobile-gd');

  report('no runtime exceptions', runtimeErrors.length === 0, runtimeErrors.join('; '));
}

const chrome = spawn(CHROME, [
  '--headless=new',
  '--disable-gpu',
  '--no-first-run',
  `--remote-debugging-port=${PORT}`,
  `--user-data-dir=${PROFILE}`,
  '--window-size=1440,900',
], { stdio: ['ignore', 'ignore', 'pipe'] });

try {
  let websocketUrl;
  for (let i = 0; i < 50 && !websocketUrl; i += 1) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      websocketUrl = targets.find((target) => target.type === 'page')?.webSocketDebuggerUrl;
    } catch {}
    if (!websocketUrl) await sleep(250);
  }
  if (!websocketUrl) throw new Error('Chrome DevTools endpoint unavailable');
  const ws = new WebSocket(websocketUrl);
  await new Promise((resolve, reject) => {
    ws.onopen = resolve;
    ws.onerror = () => reject(new Error('WebSocket connection failed'));
  });
  cdp = new CDP(ws);
  await cdp.send('Page.enable');
  await cdp.send('Runtime.enable');
  await cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1440,
    height: 900,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await cdp.send('Page.navigate', { url: APP_URL });
  if (!(await waitLoaded())) throw new Error('Page did not render');
  await verify();
} catch (error) {
  console.error(`DRIVER_FAIL ${error.message}`);
  results.push({ name: 'driver', ok: false });
} finally {
  chrome.kill();
  await sleep(400);
  const failed = results.filter((result) => !result.ok).length;
  console.log(`SUMMARY ${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
}
