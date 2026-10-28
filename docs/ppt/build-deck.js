// 生成《皮瓣状态智能辅助监测系统 — 系统架构与开发组织架构设计》评审 PPT
//
// 用法：
//   cd docs/ppt && npm install && npm run build
//
// 全部版式、配色、文字都在本文件中；改内容后重新运行即可生成新版 .pptx。

const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

const OUT = path.join(__dirname, '皮瓣状态智能辅助监测系统-架构与开发组织-v0.1.pptx');

// ───────────────────────── 主题 ─────────────────────────
const THEME = {
  name: 'QingSi Flap Monitor',
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  colors: {
    dk1: '1A2530', // 正文墨色
    lt1: 'FFFFFF',
    dk2: '0E3A47', // 深青（主色，封面/标题）
    lt2: 'EEF3F5', // 浅灰青（卡片底）
    accent1: '0F7C86', // 青绿：平台 / 医护
    accent2: 'D9566A', // 皮瓣玫红：患者 / 重点
    accent3: 'E8962E', // 琥珀：边缘 / 工业 / 临近
    accent4: '2E9E6B', // 绿：正常
    accent5: '5D6F80', // 石板灰：说明文字 / 管理员
    accent6: '6E56B8', // 紫：AI
    hlink: '0F7C86',
    folHlink: '5D6F80',
  },
};
const HEX = THEME.colors;
const MONO = 'Courier New';
const FOOTER = '皮瓣状态智能辅助监测系统 · 系统架构与开发组织架构设计';

const SEC = ['封面与概览', '01 背景与需求', '02 总体架构', '03 关键模块设计', '04 技术与工程', '05 组织与计划'];

let pres;
let C;
let seq = 0;
const nm = (p) => `${p}-${++seq}`;

// ───────────────────────── 图标 ─────────────────────────
const iconCache = {};
async function icon(name, hex) {
  const key = `${name}-${hex}`;
  if (!iconCache[key]) {
    const comp = fa[name];
    if (!comp) throw new Error(`react-icons/fa 中没有图标 ${name}`);
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(comp, { size: 256 })).replace(/currentColor/g, `#${hex}`);
    const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
    iconCache[key] = `image/png;base64,${buf.toString('base64')}`;
  }
  return iconCache[key];
}

// ───────────────────────── 绘图小工具 ─────────────────────────
const softShadow = () => ({ type: 'outer', color: '0E3A47', opacity: 0.14, blur: 8, offset: 2, angle: 90 });

function box(s, x, y, w, h, o = {}) {
  s.addShape(o.rect ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    rectRadius: o.rect ? undefined : (o.r ?? 0.1),
    fill: o.fill === null ? undefined : (o.fill ?? { color: C.background2 }),
    line: o.line,
    shadow: o.shadow ? softShadow() : undefined,
    objectName: nm(o.name || 'card'),
  });
}

function txt(s, text, x, y, w, h, o = {}) {
  s.addText(text, {
    x, y, w, h,
    isTextBox: true,
    margin: o.margin ?? 0,
    fontSize: o.size ?? 14,
    color: o.color ?? C.text1,
    bold: o.bold,
    italic: o.italic,
    align: o.align ?? 'left',
    valign: o.valign ?? 'top',
    fontFace: o.font,
    lineSpacingMultiple: o.lsm,
    paraSpaceAfter: o.psa,
    fit: 'none',
    objectName: nm(o.name || 'text'),
  });
}

// 带项目符号的列表
function bullets(s, items, x, y, w, h, o = {}) {
  const runs = items.map((t, i) => ({
    text: t,
    options: { bullet: { indent: o.indent ?? 14 }, breakLine: i < items.length - 1, paraSpaceAfter: o.psa ?? 6 },
  }));
  s.addText(runs, {
    x, y, w, h, isTextBox: true, margin: 0, fontSize: o.size ?? 14, color: o.color ?? C.text1,
    valign: 'top', fit: 'none', objectName: nm(o.name || 'list'),
  });
}

async function iconCircle(s, iconName, x, y, d, fill, iconHex, o = {}) {
  s.addShape(pres.shapes.OVAL, {
    x, y, w: d, h: d,
    fill: { color: fill, transparency: o.transparency },
    line: o.line,
    objectName: nm('icon-bg'),
  });
  const k = d * (o.scale ?? 0.5);
  s.addImage({ data: await icon(iconName, iconHex), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k, altText: o.alt || iconName, objectName: nm('icon') });
}

function arrow(s, x1, y1, x2, y2, o = {}) {
  s.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line: {
      color: o.color ?? C.accent5,
      width: o.width ?? 1.5,
      dashType: o.dash,
      transparency: o.transparency,
      endArrowType: o.noHead ? undefined : 'triangle',
      beginArrowType: o.both ? 'triangle' : undefined,
    },
    objectName: nm(o.name || 'connector'),
  });
}

// 表格：header 为表头文字数组，rows 中每格可为字符串或 { text, options }
function table(s, header, rows, o) {
  const hdr = header.map((h, i) => ({
    text: h,
    options: { bold: true, color: o.headColor?.[i] ?? C.background1, fill: { color: o.headFill?.[i] ?? C.text2 }, valign: 'middle', align: o.align?.[i] ?? 'left' },
  }));
  const body = rows.map((r, ri) => r.map((cell, ci) => {
    const c = typeof cell === 'string' ? { text: cell } : cell;
    return {
      text: c.text,
      options: {
        color: C.text1, fill: { color: ri % 2 ? C.background2 : C.background1 }, valign: 'middle', align: o.align?.[ci] ?? 'left',
        ...(o.colOpts?.[ci] || {}), ...(c.options || {}),
      },
    };
  }));
  s.addTable([hdr, ...body], {
    x: o.x, y: o.y, w: o.w, colW: o.colW, rowH: o.rowH, fontSize: o.size ?? 14,
    border: { type: 'solid', pt: 0.75, color: 'D5DEE3' }, margin: [0.04, 0.1, 0.04, 0.1],
    objectName: nm('table'),
  });
}

function contentSlide(section, kicker, title) {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: section });
  s.addText(kicker, { placeholder: 'kicker', bullet: false });
  s.addText(title, { placeholder: 'title' });
  return s;
}

// ───────────────────────── 版式 ─────────────────────────
function defineLayouts() {
  pres.defineSlideMaster({
    title: 'COVER',
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 2.3, w: 7.6, h: 1.1, fontSize: 44, bold: true, color: C.background1, align: 'left', valign: 'middle', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.7, y: 3.45, w: 7.6, h: 0.7, fontSize: 26, color: C.background2, valign: 'middle', margin: 0 }, text: '' } },
    ],
  });
  pres.defineSlideMaster({
    title: 'CONTENT',
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: 'kicker', type: 'body', x: 0.6, y: 0.3, w: 9, h: 0.32, fontSize: 12, bold: true, color: C.accent1, valign: 'middle', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.62, w: 12.13, h: 0.78, fontSize: 32, bold: true, color: C.text2, align: 'left', valign: 'middle', margin: 0 }, text: '' } },
      { text: { text: FOOTER, options: { x: 0.6, y: 7.02, w: 8, h: 0.28, fontSize: 10, color: C.accent5, margin: 0 } } },
    ],
    slideNumber: { x: 12.13, y: 7.02, w: 0.6, h: 0.28, fontSize: 10, color: HEX.accent5, align: 'right', margin: 0 },
  });
  pres.defineSlideMaster({
    title: 'CLOSING',
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: 'kicker', type: 'body', x: 0.6, y: 0.3, w: 9, h: 0.32, fontSize: 12, bold: true, color: C.accent3, valign: 'middle', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.62, w: 12.13, h: 0.78, fontSize: 32, bold: true, color: C.background1, align: 'left', valign: 'middle', margin: 0 }, text: '' } },
    ],
  });
}

// ───────────────────────── 幻灯片 ─────────────────────────

async function slideCover() {
  const s = pres.addSlide({ masterName: 'COVER', sectionTitle: SEC[0] });
  s.addText('皮瓣状态智能辅助监测系统', { placeholder: 'title' });
  s.addText('系统架构与开发组织架构设计', { placeholder: 'body', bullet: false });
  txt(s, '开发前架构评审 · v0.1', 0.7, 1.7, 6, 0.4, { size: 14, bold: true, color: C.accent3 });
  txt(s, '微信小程序 × 边缘数采 × 云平台 × AI 预留\n一套架构，覆盖医疗皮瓣监测与烟草在线监测两个场景', 0.7, 4.45, 7.3, 0.9, { size: 16, color: C.background2, lsm: 1.25 });
  txt(s, 'QingSi 项目组  ·  2026 年 10 月', 0.7, 6.4, 6, 0.4, { size: 14, color: C.background2 });

  // 端—边—云—智 网络示意
  const N = { cloud: [10.5, 2.4], end: [9.2, 4.1], edge: [11.8, 4.1], ai: [10.5, 5.6] };
  const d = 1.15;
  const line = { color: C.background2, width: 2, noHead: true, transparency: 55 };
  arrow(s, ...N.end, ...N.cloud, line);
  arrow(s, ...N.edge, ...N.cloud, { ...line });
  arrow(s, ...N.cloud, ...N.ai, { ...line });
  const node = async (k, ic, fill, iconHex, label, labelAbove) => {
    const [cx, cy] = N[k];
    await iconCircle(s, ic, cx - d / 2, cy - d / 2, d, fill, iconHex, { scale: 0.46 });
    txt(s, label, cx - 0.9, labelAbove ? cy - d / 2 - 0.42 : cy + d / 2 + 0.08, 1.8, 0.34, { size: 13, color: C.background2, align: 'center', bold: true });
  };
  await node('cloud', 'FaCloud', C.accent1, 'FFFFFF', '云 · 平台服务', true);
  await node('end', 'FaMobileAlt', C.accent2, 'FFFFFF', '端 · 微信小程序');
  await node('edge', 'FaMicrochip', C.accent3, HEX.dk2, '边 · 数采端');
  await node('ai', 'FaBrain', C.accent6, 'FFFFFF', '智 · AI（预留）');

  s.addNotes('开场：本次汇报目的——在正式编码前，把系统怎么分层、模块怎么切、两个人怎么分工、分几期做讲清楚，请老师评审。核心观点：皮瓣监测和工业边缘-云监测是同一类问题，我们按“平台 + 领域包”来设计，一套架构服务两个场景。');
}

async function slideSummary() {
  const s = contentSlide(SEC[0], '概览', '一页看懂：做什么、怎么做');
  const stats = [
    ['3 类', '用户角色', '患者 · 医护 · 管理员'],
    ['1h → 4h', '采样频率', '术后 72h 内每小时一次，之后每 4 小时一次'],
    ['4 层', '技术架构', '端 · 边 · 云 · 智'],
    ['2 个', '应用场景', '皮瓣监测 → 烟草在线监测'],
  ];
  const w = (12.13 - 3 * 0.3) / 4;
  stats.forEach(([big, label, cap], i) => {
    const x = 0.6 + i * (w + 0.3);
    box(s, x, 1.65, w, 2.05, { name: 'stat-card' });
    txt(s, big, x + 0.3, 1.8, w - 0.6, 0.75, { size: 40, bold: true, color: i === 1 ? C.accent2 : C.accent1, valign: 'middle' });
    txt(s, label, x + 0.3, 2.6, w - 0.6, 0.36, { size: 16, bold: true, color: C.text2 });
    txt(s, cap, x + 0.3, 2.98, w - 0.6, 0.62, { size: 13, color: C.accent5 });
  });

  txt(s, '关键设计决策', 0.6, 3.95, 6, 0.4, { size: 18, bold: true, color: C.text2 });
  const items = [
    ['FaMobileAlt', '一个小程序，三套工作台', '微信小程序按登录角色切换患者 / 医护 / 管理员界面，一套代码维护'],
    ['FaCubes', '平台 + 领域包', '核心平台与业务解耦，场景差异放进配置：医疗、烟草各是一个领域包'],
    ['FaMicrochip', '通用边缘数采端', 'Windows / Linux 同一套代码：断网缓存、断点续传，可本地推理'],
    ['FaPlug', 'AI 接口先留好', '标准化分析插件契约；一期积累医生标注数据，二期直接接入模型'],
  ];
  const cw = (12.13 - 0.3) / 2;
  for (let i = 0; i < items.length; i++) {
    const [ic, t, dsc] = items[i];
    const x = 0.6 + (i % 2) * (cw + 0.3);
    const y = 4.5 + Math.floor(i / 2) * 1.2;
    await iconCircle(s, ic, x, y + 0.04, 0.62, C.accent1, 'FFFFFF');
    txt(s, t, x + 0.85, y, cw - 0.9, 0.38, { size: 16, bold: true, color: C.text2 });
    txt(s, dsc, x + 0.85, y + 0.4, cw - 0.95, 0.66, { size: 14 });
  }
  s.addNotes('用四个数字概括项目：3 类用户、1 小时 / 4 小时的采样节奏、端边云智 4 层架构、2 个应用场景。下面四条是本次评审最希望老师拍板的设计决策。');
}

async function slideAgenda() {
  const s = contentSlide(SEC[0], '目录', '汇报目录');
  const parts = [
    ['01', '背景与需求', 'FaHospital', ['临床背景', '三类用户与功能']],
    ['02', '总体架构', 'FaLayerGroup', ['六条设计原则', '端边云智四层', '与工业监测同构', '通用领域模型']],
    ['03', '关键模块', 'FaCubes', ['小程序三端', '权限模型', '核心流程', '采样策略', '实时监控墙', '数采端', 'AI 预留接口']],
    ['04', '技术与工程', 'FaCode', ['技术选型', '代码仓库', '接口契约', '安全合规', '部署形态']],
    ['05', '组织与计划', 'FaUsers', ['开发组织架构', '分工与协作', '路线图', '风险与应对', '待确认事项']],
  ];
  const w = (12.13 - 4 * 0.25) / 5;
  for (let i = 0; i < parts.length; i++) {
    const [no, t, ic, list] = parts[i];
    const x = 0.6 + i * (w + 0.25);
    box(s, x, 1.65, w, 4.15, { name: 'agenda-card' });
    txt(s, no, x + 0.25, 1.85, 1.2, 0.7, { size: 36, bold: true, color: C.accent1, valign: 'middle' });
    await iconCircle(s, ic, x + w - 0.8, 1.93, 0.55, C.accent1, 'FFFFFF');
    txt(s, t, x + 0.25, 2.7, w - 0.4, 0.45, { size: 20, bold: true, color: C.text2 });
    txt(s, list.join('\n'), x + 0.25, 3.3, w - 0.4, 2.4, { size: 14, color: C.text1, lsm: 1.3 });
  }
  box(s, 0.6, 6.05, 12.13, 0.8, { fill: { color: C.text2 }, name: 'key-message' });
  txt(s, [
    { text: '核心观点：', options: { bold: true, color: C.accent3 } },
    { text: '皮瓣监测只是第一个场景，真正要建的是可复用的“边云协同 + 实时看板 + AI 分析”监测平台', options: { color: C.background1 } },
  ], 0.9, 6.05, 11.6, 0.8, { size: 15, valign: 'middle' });
  s.addNotes('共五部分：先讲为什么做、给谁用；再讲总体架构和“一物多用”的思路；然后逐个模块展开；最后是技术选型、工程规范和我们两个人的分工与计划。');
}

async function slideBackground() {
  const s = contentSlide(SEC[1], '01  背景与需求', '临床背景：术后 72 小时是皮瓣存活的关键窗口');
  s.addText([
    { text: '皮瓣移植', options: { bold: true, color: C.accent2 } },
    { text: '：从身体其他部位（供区）切取带血管的皮肤及皮下组织，转移覆盖创面。术后一旦出现血管危象而未及时处理，皮瓣可能部分或全部坏死。' },
  ], { x: 0.6, y: 1.65, w: 5.75, h: 1.3, isTextBox: true, margin: 0, fontSize: 15, color: C.text1, valign: 'top', lineSpacingMultiple: 1.2, fit: 'none', objectName: nm('intro') });

  txt(s, '需要重点观察的指标', 0.6, 3.05, 5.7, 0.4, { size: 16, bold: true, color: C.text2 });
  const ind = [
    ['FaEye', '颜色', '苍白 → 动脉供血不足；青紫 → 静脉回流障碍', C.accent2],
    ['FaLayerGroup', '纹理', '皮纹、皱褶、水疱与表面光泽变化', C.accent1],
    ['FaExpandArrowsAlt', '肿胀', '肿胀程度与张力，是否进行性加重', C.accent3],
    ['FaMicrochip', '扩展指标', '皮温、毛细血管回流等，后续可接传感器', C.accent5],
  ];
  const cw = (5.75 - 0.25) / 2;
  for (let i = 0; i < ind.length; i++) {
    const [ic, t, dsc, col] = ind[i];
    const x = 0.6 + (i % 2) * (cw + 0.25);
    const y = 3.55 + Math.floor(i / 2) * 1.6;
    box(s, x, y, cw, 1.4, { name: 'indicator' });
    await iconCircle(s, ic, x + 0.2, y + 0.18, 0.48, col, i === 2 ? HEX.dk2 : 'FFFFFF');
    txt(s, t, x + 0.82, y + 0.2, cw - 1.0, 0.44, { size: 15, bold: true, color: C.text2, valign: 'middle' });
    txt(s, dsc, x + 0.2, y + 0.76, cw - 0.4, 0.6, { size: 14 });
  }

  // 右侧：监测频率时间轴 + 痛点
  const X = 6.7, W = 6.03;
  box(s, X, 1.65, W, 5.2, { fill: { color: C.background1 }, shadow: true, line: { color: 'D5DEE3', width: 0.75 }, name: 'timeline-card' });
  txt(s, '监测频率（以手术结束为 T0）', X + 0.3, 1.85, W - 0.6, 0.4, { size: 16, bold: true, color: C.text2 });
  const x0 = X + 0.35, span = W - 0.7, perH = span / 120, barY = 2.85;
  const xA = x0 + 72 * perH;
  s.addShape(pres.shapes.RECTANGLE, { x: x0, y: barY, w: xA - x0, h: 0.14, fill: { color: C.accent2 }, objectName: nm('bar-72h') });
  s.addShape(pres.shapes.RECTANGLE, { x: xA, y: barY, w: x0 + span - xA, h: 0.14, fill: { color: C.accent1 }, objectName: nm('bar-after') });
  for (let h = 0; h <= 72; h++) {
    s.addShape(pres.shapes.LINE, { x: x0 + h * perH, y: barY - 0.28, w: 0, h: 0.26, line: { color: C.accent2, width: 0.5 }, objectName: nm('tick') });
  }
  for (let h = 76; h <= 120; h += 4) {
    s.addShape(pres.shapes.LINE, { x: x0 + h * perH, y: barY - 0.28, w: 0, h: 0.26, line: { color: C.accent1, width: 1.25 }, objectName: nm('tick') });
  }
  [0, 24, 48, 72, 96, 120].forEach((h) => {
    txt(s, h === 0 ? 'T0' : `${h}h`, x0 + h * perH - 0.3, barY + 0.2, 0.6, 0.26, { size: 11, color: C.accent5, align: 'center' });
  });
  const legend = [
    [C.accent2, '0–72 h：每 1 小时上传 1 次，共 72 次'],
    [C.accent1, '72 h 之后：每 4 小时上传 1 次，每天 6 次'],
  ];
  legend.forEach(([col, t], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: x0, y: 3.55 + i * 0.42 + 0.08, w: 0.18, h: 0.18, fill: { color: col }, objectName: nm('legend-key') });
    txt(s, t, x0 + 0.3, 3.55 + i * 0.42, span - 0.3, 0.34, { size: 14, valign: 'middle' });
  });

  txt(s, '当前痛点', X + 0.3, 4.55, W - 0.6, 0.4, { size: 16, bold: true, color: C.text2 });
  const pains = [
    ['FaClock', '夜间人手紧张，巡视频次高'],
    ['FaClipboardList', '纸质记录分散，难以前后对比'],
    ['FaStethoscope', '早期变化细微，依赖个人经验'],
    ['FaVideo', '无法远程查看与会诊'],
  ];
  const pw = (W - 0.6 - 0.2) / 2;
  for (let i = 0; i < pains.length; i++) {
    const [ic, t] = pains[i];
    const x = X + 0.3 + (i % 2) * (pw + 0.2);
    const y = 5.05 + Math.floor(i / 2) * 0.85;
    await iconCircle(s, ic, x, y + 0.05, 0.44, C.background2, HEX.dk2);
    txt(s, t, x + 0.58, y, pw - 0.6, 0.56, { size: 14, valign: 'middle' });
  }
  s.addNotes('皮瓣移植后最危险的是血管危象：动脉供血不足时皮瓣苍白、皮温下降；静脉回流障碍时皮瓣青紫、肿胀。术后前 72 小时最关键，需要每小时观察一次，之后每 4 小时一次。目前靠护士人工巡视，夜间压力大、记录分散。系统要解决的就是“按时采集 + 远程查看 + 留档对比”，并为后续 AI 判读打基础。');
}

async function slideUsers() {
  const s = contentSlide(SEC[1], '01  背景与需求', '三类用户，一套系统');
  const roles = [
    ['FaUserInjured', '患者 / 家属', '数据提供方', C.accent2, '诉求：操作简单，3 步完成上传', [
      '扫描病例二维码，绑定监测任务',
      '按时拍照上传（带拍摄引导）',
      '查看本人上传记录与下次时间',
      '接收提醒，超时可补传',
    ]],
    ['FaUserMd', '医生 / 护士', '数据使用方', C.accent1, '诉求：一眼看出谁有异常', [
      '查看任一负责患者的全部历史图像',
      '实时监控墙：所有患者一屏总览',
      '时间轴对比、评估标注、处置记录',
      '接收超时与（预留）AI 预警',
    ]],
    ['FaUserShield', '管理员', '系统运维方', C.accent5, '诉求：账号与配置可控、可追溯', [
      '新增 / 停用用户，分配角色与科室',
      '配置采样策略模板',
      '登记与管理数采端设备',
      '查看审计日志与运行状态',
    ]],
  ];
  const w = (12.13 - 2 * 0.3) / 3;
  for (let i = 0; i < roles.length; i++) {
    const [ic, name, sub, col, need, list] = roles[i];
    const x = 0.6 + i * (w + 0.3);
    box(s, x, 1.65, w, 4.3, { fill: { color: C.background1 }, shadow: true, line: { color: 'D5DEE3', width: 0.75 }, name: 'role-card' });
    await iconCircle(s, ic, x + 0.3, 1.95, 0.75, col, 'FFFFFF');
    txt(s, name, x + 1.25, 1.93, w - 1.4, 0.45, { size: 22, bold: true, color: C.text2, valign: 'middle' });
    txt(s, sub, x + 1.25, 2.38, w - 1.4, 0.3, { size: 13, color: C.accent5 });
    bullets(s, list, x + 0.35, 3.0, w - 0.6, 2.05, { psa: 14 });
    box(s, x + 0.25, 5.2, w - 0.5, 0.52, { fill: { color: col, transparency: 88 }, r: 0.08, name: 'role-need' });
    txt(s, need, x + 0.4, 5.2, w - 0.8, 0.52, { size: 14, bold: true, color: C.text2, valign: 'middle' });
  }
  box(s, 0.6, 6.17, 12.13, 0.66, { fill: { color: C.accent1, transparency: 88 }, name: 'reserve-band' });
  await iconCircle(s, 'FaPlug', 0.8, 6.29, 0.42, C.accent1, 'FFFFFF');
  txt(s, [
    { text: '预留扩展：', options: { bold: true, color: C.text2 } },
    { text: 'AI 实时分析 · 床旁摄像头实时视频 · 更多角色（护士长 / 科研人员 / 质检员）· 工业场景接入' },
  ], 1.4, 6.17, 11.1, 0.66, { size: 14, valign: 'middle' });
  s.addNotes('三类用户对应老师提出的三个角色：患者负责“产生数据”，医护负责“使用数据”，管理员负责“管理系统”。注意医生端有两个核心功能：一是按患者查看全部历史图像，二是像车辆监控一样的实时监控墙。角色不写死，后续可以加护士长、科研人员，工业场景里就是质检员、工艺员。');
}

async function slidePrinciples() {
  const s = contentSlide(SEC[2], '02  总体架构', '六条设计原则：可拓展、一物多用');
  const ps = [
    ['FaCubes', '平台 + 领域包', '核心平台与业务解耦，医疗、烟草各自是一个可插拔的“领域包”'],
    ['FaSlidersH', '配置驱动', '采样频率、术语、表单、告警规则都写成配置，而不是写死在代码里'],
    ['FaFileContract', '接口先行', '先定 OpenAPI 接口契约，前端、后端、数采端、AI 四方并行开发'],
    ['FaNetworkWired', '边云协同', '边缘就近采集、缓存与预处理；云端集中存储、调度与展示'],
    ['FaBolt', '事件驱动', '“新图片到达”等事件进入消息总线，AI、告警、大屏按需订阅'],
    ['FaShieldAlt', '安全合规', '医疗影像属敏感个人信息：最小权限、全程加密、操作留痕'],
  ];
  const w = (12.13 - 2 * 0.3) / 3, h = 2.45;
  for (let i = 0; i < ps.length; i++) {
    const [ic, t, dsc] = ps[i];
    const x = 0.6 + (i % 3) * (w + 0.3);
    const y = 1.65 + Math.floor(i / 3) * (h + 0.3);
    box(s, x, y, w, h, { name: 'principle' });
    await iconCircle(s, ic, x + 0.3, y + 0.3, 0.65, C.accent1, 'FFFFFF');
    txt(s, t, x + 0.3, y + 1.1, w - 0.6, 0.45, { size: 18, bold: true, color: C.text2 });
    txt(s, dsc, x + 0.3, y + 1.58, w - 0.55, 0.8, { size: 14 });
  }
  s.addNotes('这六条原则是后面所有设计的依据。最重要的是第一条和第二条：平台本身不认识“皮瓣”，皮瓣只是一个领域包；所有会变的东西（频率、术语、表单、规则）都放在配置里。这样换到烟草在线监测时，主要工作是写一个新的领域包，而不是重写系统。');
}

async function slideArchitecture() {
  const s = contentSlide(SEC[2], '02  总体架构', '总体架构：端 · 边 · 云 · 智 四层协同');
  // 左：采集侧
  txt(s, '采集侧 · 端 + 边', 0.6, 1.58, 2.5, 0.3, { size: 13, bold: true, color: C.accent2 });
  const sideBox = async (x, y, ic, col, iconHex, t, dsc, dashed) => {
    box(s, x, y, 2.5, 0.98, { fill: { color: C.background2 }, line: dashed ? { color: C.accent5, width: 1, dashType: 'dash' } : undefined, name: 'node' });
    await iconCircle(s, ic, x + 0.14, y + 0.13, 0.42, col, iconHex);
    txt(s, t, x + 0.66, y + 0.13, 1.8, 0.42, { size: 14, bold: true, color: C.text2, valign: 'middle' });
    txt(s, dsc, x + 0.14, y + 0.6, 2.3, 0.32, { size: 12, color: C.accent5 });
  };
  await sideBox(0.6, 1.95, 'FaMobileAlt', C.accent2, 'FFFFFF', '患者小程序', '拍照上传 · 查看记录');
  await sideBox(0.6, 3.1, 'FaMicrochip', C.accent3, HEX.dk2, '边缘数采端', 'Win / Linux · 摄像头 · 相机');
  await sideBox(0.6, 4.25, 'FaPlug', C.accent5, 'FFFFFF', '第三方（预留）', 'HIS · MES · 设备系统', true);

  // 右：应用侧
  txt(s, '应用侧 · 查看与管理', 10.23, 1.58, 2.5, 0.3, { size: 13, bold: true, color: C.accent1 });
  await sideBox(10.23, 1.95, 'FaUserMd', C.accent1, 'FFFFFF', '医护小程序', '历史图像 · 评估标注');
  await sideBox(10.23, 3.1, 'FaTv', C.accent1, 'FFFFFF', '监控大屏', '实时宫格 · 状态告警');
  await sideBox(10.23, 4.25, 'FaUserCog', C.accent5, 'FFFFFF', '管理后台', '用户 · 策略 · 设备');

  // 中：云平台
  txt(s, '云平台 · 核心服务', 3.5, 1.58, 6.35, 0.3, { size: 13, bold: true, color: C.text2 });
  const CX = 3.5, CW = 6.35, chipX = 4.42, chipW = (5.35 - 3 * 0.1) / 4;
  const layer = (y, h, label, fill) => {
    box(s, CX, y, CW, h, { fill, r: 0.08, name: 'layer' });
    txt(s, label, CX + 0.08, y, 0.8, h, { size: 13, bold: true, color: C.text2, valign: 'middle', align: 'center' });
  };
  const chip = (i, y, h, t, cap) => {
    const x = chipX + i * (chipW + 0.1);
    box(s, x, y, chipW, h, { fill: { color: C.background1 }, r: 0.06, name: 'chip' });
    if (cap) {
      txt(s, t, x + 0.03, y + 0.08, chipW - 0.06, 0.3, { size: 12, bold: true, color: C.text2, align: 'center' });
      txt(s, cap, x + 0.03, y + 0.38, chipW - 0.06, 0.28, { size: 10, color: C.accent5, align: 'center' });
    } else {
      txt(s, t, x + 0.03, y, chipW - 0.06, h, { size: 12, color: C.text2, align: 'center', valign: 'middle' });
    }
  };
  layer(1.95, 0.62, '接入层', { color: C.accent1, transparency: 86 });
  ['API 网关', '统一鉴权', 'WebSocket', 'MQTT（预留）'].forEach((t, i) => chip(i, 2.03, 0.46, t));
  layer(2.72, 1.74, '服务层', { color: C.accent1, transparency: 86 });
  [
    ['用户与权限', 'RBAC · 数据范围'], ['组织与对象', '医院 · 科室 · 患者'], ['监测任务', '建档 · 绑定 · 周期'], ['采样排程', '策略 · 提醒 · 超时'],
  ].forEach(([t, c], i) => chip(i, 2.82, 0.72, t, c));
  [
    ['媒体服务', '直传 · 缩略图'], ['监控墙', '快照聚合 · 推送'], ['告警规则', '阈值 · 升级 · 通知'], ['审计日志', '查看 · 导出留痕'],
  ].forEach(([t, c], i) => chip(i, 3.64, 0.72, t, c));
  layer(4.61, 0.62, '数据层', { color: C.accent5, transparency: 85 });
  ['PostgreSQL', 'Redis 队列', '对象存储', '时序库（预留）'].forEach((t, i) => chip(i, 4.69, 0.46, t));

  // 连线
  [2.44, 3.59, 4.74].forEach((y, i) => arrow(s, 3.12, y, 3.48, y, { color: C.accent5, dash: i === 2 ? 'dash' : undefined }));
  [2.44, 3.59, 4.74].forEach((y) => arrow(s, 9.87, y, 10.21, y, { color: C.accent5 }));

  // 智能层
  box(s, 0.6, 5.5, 12.13, 0.86, { fill: { color: C.accent6, transparency: 86 }, r: 0.08, name: 'ai-layer' });
  txt(s, '智能层\n（预留）', 0.7, 5.5, 1.1, 0.86, { size: 13, bold: true, color: C.accent6, valign: 'middle', align: 'center' });
  const aiChips = ['AI 网关：插件注册与路由', '皮瓣模型：颜色·纹理·肿胀', '烟草缺陷检测模型', '边缘轻量模型（ONNX）'];
  const aw = (12.13 - 1.4 - 3 * 0.15) / 4;
  aiChips.forEach((t, i) => {
    const x = 1.9 + i * (aw + 0.15);
    box(s, x, 5.62, aw, 0.62, { fill: { color: C.background1 }, r: 0.06, name: 'ai-chip' });
    txt(s, t, x + 0.05, 5.62, aw - 0.1, 0.62, { size: 12, color: C.text2, align: 'center', valign: 'middle' });
  });
  arrow(s, 6.67, 5.25, 6.67, 5.6, { color: C.accent6, both: true });
  txt(s, '数据流：采集侧 → 接入层 → 服务层 → 数据层；“新采样”等事件经消息总线驱动监控墙、告警与 AI 分析', 0.6, 6.5, 12.13, 0.32, { size: 12, color: C.accent5 });
  s.addNotes('总体架构从左到右是数据流：左边是采集侧，包括患者手机小程序和边缘数采端；中间是云平台，分接入层、服务层、数据层；右边是查看和管理的应用。底下是预留的智能层，云端和边缘都可以调用模型。患者端和医护端其实是同一个小程序，这里为了表达数据流向分开画。');
}

async function slideAnalogy() {
  const s = contentSlide(SEC[2], '02  总体架构', '与工业边云监测同构：一套架构，两个行业');
  const rows = [
    ['监测对象', '患者的移植皮瓣', '烟支 / 烟包 / 设备工位'],
    ['采集端', '患者手机、床旁摄像头', '产线工业相机、传感器、PLC'],
    ['边缘节点', '病区护士站电脑（Windows）', '工控机（Windows / Linux）'],
    ['采样策略', '术后 0–72 h 每 1 h，之后每 4 h', '按节拍 / 班次 / 批次'],
    ['分析模型', '颜色 · 纹理 · 肿胀 → 坏死风险', '外观缺陷 · 尺寸 · 色差 → 质量判定'],
    ['实时看板', '病区患者监控墙', '车间中控大屏'],
    ['异常处置', '通知医生、记录处置', '报警、剔除、工艺调整'],
    ['组织层级', '医院 → 科室 → 病区 → 床位', '工厂 → 车间 → 产线 → 工位'],
  ].map(([a, b, c]) => [{ text: a, options: { bold: true, color: C.text2 } }, b, c]);
  table(s, ['架构要素', '皮瓣监测（医疗）', '烟草在线监测（工业）'], rows, {
    x: 0.6, y: 1.65, w: 12.13, colW: [2.2, 4.965, 4.965], rowH: 0.5, size: 14,
    headFill: [C.text2, C.accent1, C.accent3], headColor: [C.background1, C.background1, C.text1],
  });
  box(s, 0.6, 6.25, 12.13, 0.6, { fill: { color: C.accent1, transparency: 88 }, name: 'conclusion' });
  await iconCircle(s, 'FaCubes', 0.78, 6.34, 0.42, C.accent1, 'FFFFFF');
  txt(s, [
    { text: '结论：', options: { bold: true, color: C.text2 } },
    { text: '差异只在“领域包”（术语 · 表单 · 策略 · 模型 · 大屏模板），平台核心代码复用，目标复用率 ≥ 80%' },
  ], 1.35, 6.25, 11.2, 0.6, { size: 14, valign: 'middle' });
  s.addNotes('这页回答“一物多用”：把皮瓣监测和烟草在线监测逐项对照，每一行在架构上都是同一个东西，只是名字和参数不同。所以我们做的是一个“边缘采集 + 云端汇聚 + 实时看板 + AI 分析”的通用监测平台，皮瓣是第一个落地场景。80% 是目标值，用来约束我们不要把业务写死在核心代码里。');
}

async function slideDomainModel() {
  const s = contentSlide(SEC[2], '02  总体架构', '通用领域模型：把“患者”抽象为“监测对象”');
  const w = (12.13 - 3 * 0.35) / 4, h = 1.6;
  const rowsDef = [
    ['① 组织与对象：谁被监测', 1.6, [
      ['租户', 'Tenant', '医院', '卷烟厂'],
      ['组织', 'Org', '科室 / 病区', '车间 / 产线'],
      ['监测对象', 'Subject', '患者皮瓣', '机台 / 工位'],
      ['监测任务', 'Task', '术后监测周期', '一个生产批次'],
    ]],
    ['② 数据与智能：监测到了什么', 3.95, [
      ['采集源', 'Source', '手机 / 摄像头', '工业相机'],
      ['采样', 'Sample', '一张创面照片', '一帧图像'],
      ['分析结果', 'Inference', '坏死风险评分', '缺陷类型判定'],
      ['事件告警', 'Alert', '疑似血管危象', '质量异常报警'],
    ]],
  ];
  rowsDef.forEach(([cap, y0, ents]) => {
    txt(s, cap, 0.6, y0, 8, 0.34, { size: 14, bold: true, color: C.accent1 });
    ents.forEach(([cn, en, med, ind], i) => {
      const x = 0.6 + i * (w + 0.35), y = y0 + 0.4;
      box(s, x, y, w, h, { name: 'entity' });
      txt(s, cn, x + 0.22, y + 0.15, w - 0.4, 0.4, { size: 17, bold: true, color: C.text2 });
      txt(s, en, x + 0.22, y + 0.55, w - 0.4, 0.26, { size: 11, color: C.accent5, font: MONO });
      [[C.accent1, `医疗：${med}`], [C.accent3, `烟草：${ind}`]].forEach(([col, t], k) => {
        s.addShape(pres.shapes.OVAL, { x: x + 0.24, y: y + 0.99 + k * 0.36, w: 0.13, h: 0.13, fill: { color: col }, objectName: nm('dot') });
        txt(s, t, x + 0.45, y + 0.9 + k * 0.36, w - 0.55, 0.32, { size: 14, valign: 'middle' });
      });
      if (i < 3) arrow(s, x + w + 0.04, y + h / 2, x + w + 0.31, y + h / 2, { color: C.accent5 });
    });
  });
  box(s, 0.6, 6.33, 12.13, 0.52, { fill: { color: C.accent5, transparency: 88 }, name: 'note' });
  txt(s, [
    { text: '落地方式：', options: { bold: true, color: C.text2 } },
    { text: '通用字段建表固定；领域特有字段存入 JSONB 扩展属性，并由领域包中的表单 Schema 描述 —— 新场景无需改表结构' },
  ], 0.8, 6.33, 11.8, 0.52, { size: 14, valign: 'middle' });
  s.addNotes('数据库和代码里不出现“患者表”“皮瓣表”这样的强业务命名，而是用 8 个通用实体。比如“患者”在系统里是一个监测对象，“一次术后 7 天的观察”是一个监测任务，“一张照片”是一个采样。皮瓣部位、手术方式这些医疗特有的字段放在扩展属性里，由领域包定义表单。');
}

async function slideMiniapp() {
  const s = contentSlide(SEC[3], '03  关键模块设计', '用户端：一个小程序，按角色呈现三套工作台');
  const phones = [
    ['患者端', C.accent2, [['首页', '下次上传倒计时'], ['拍照上传', '引导框 + 色卡，自动质检'], ['我的记录', '按时间查看历史照片'], ['消息', '提醒订阅 · 医护留言']]],
    ['医护端', C.accent1, [['监控墙', '宫格总览，异常置顶'], ['患者列表', '按病区 / 风险筛选'], ['患者详情', '时间轴 · 前后对比'], ['评估标注', '正常 / 可疑 / 异常']]],
    ['管理端', C.accent5, [['用户管理', '新增 · 停用 · 分配角色'], ['组织与床位', '医院 / 科室 / 病区'], ['采样策略', '频率模板配置'], ['设备与审计', '数采端登记 · 操作日志']]],
  ];
  const pw = 2.25;
  phones.forEach(([role, col, rows], i) => {
    const x = 0.6 + i * (pw + 0.3), y = 1.7;
    box(s, x, y, pw, 5.0, { fill: { color: C.background1 }, r: 0.28, line: { color: C.text2, width: 2.25 }, name: 'phone' });
    box(s, x + pw / 2 - 0.3, y + 0.14, 0.6, 0.07, { fill: { color: C.text2 }, r: 0.03, name: 'phone-notch' });
    box(s, x + 0.15, y + 0.35, pw - 0.3, 0.55, { fill: { color: col }, r: 0.08, name: 'phone-header' });
    txt(s, role, x + 0.15, y + 0.35, pw - 0.3, 0.55, { size: 15, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    rows.forEach(([t, c], k) => {
      const ry = y + 1.05 + k * 0.92;
      box(s, x + 0.15, ry, pw - 0.3, 0.8, { r: 0.08, name: 'screen' });
      txt(s, t, x + 0.27, ry + 0.08, pw - 0.5, 0.32, { size: 14, bold: true, color: C.text2 });
      txt(s, c, x + 0.27, ry + 0.4, pw - 0.5, 0.36, { size: 11, color: C.accent5 });
    });
  });
  const notes = [
    ['FaKey', '登录与绑定', '微信授权登录。医护账号由管理员预建，首次登录绑定微信；患者扫描护士出示的二维码绑定病例'],
    ['FaRoute', '按角色路由', '登录后按角色加载对应的菜单与页面；前端只做展示控制，权限由后端统一校验'],
    ['FaLayerGroup', '跨端复用', 'uni-app（Vue 3）一套代码编译为微信小程序与 H5，监控大屏和工业 PC 浏览器可直接复用'],
  ];
  for (let i = 0; i < notes.length; i++) {
    const [ic, t, dsc] = notes[i];
    const x = 8.45, y = 1.75 + i * 1.7;
    await iconCircle(s, ic, x, y, 0.55, C.accent1, 'FFFFFF');
    txt(s, t, x + 0.75, y + 0.04, 3.5, 0.42, { size: 16, bold: true, color: C.text2, valign: 'middle' });
    txt(s, dsc, x + 0.75, y + 0.5, 3.53, 1.0, { size: 14 });
  }
  s.addNotes('三类用户共用一个微信小程序，登录后按角色进入不同的工作台，这样只需要维护一个小程序、一次审核。患者端越简单越好：首页就是倒计时和一个大大的拍照按钮。医护端以监控墙为首页。管理端功能用得少，也可以放到 Web 后台。权限一定以后端校验为准，前端隐藏菜单只是体验优化。');
}

async function slideRbac() {
  const s = contentSlide(SEC[3], '03  关键模块设计', '权限模型：RBAC 角色 + 数据范围双重控制');
  const Y = { text: '✓', options: { color: C.accent4, bold: true } };
  const N = { text: '—', options: { color: C.accent5 } };
  const rows = [
    ['上传图片', '本人', '代拍上传', N],
    ['查看历史图像', '本人', '负责 / 本科室', '只读审计'],
    ['实时监控墙', N, '本科室', '全院'],
    ['评估标注 · 处置记录', N, Y, N],
    ['采样策略配置', N, '本科室调整', '全局模板'],
    ['用户与角色管理', N, N, Y],
    ['数采端设备管理', N, N, Y],
  ].map((r) => r.map((c, i) => (i === 0 ? { text: c, options: { bold: true, color: C.text2 } } : c)));
  table(s, ['功能 \\ 角色', '患者', '医护', '管理员'], rows, {
    x: 0.6, y: 1.65, w: 7.6, colW: [2.8, 1.5, 1.8, 1.5], rowH: 0.52, size: 14,
    headFill: [C.text2, C.accent2, C.accent1, C.accent5], align: ['left', 'center', 'center', 'center'],
  });
  txt(s, '医护可再细分：护士（上传、查看）· 医生（+评估、调策略）· 科主任（+全科统计）', 0.6, 6.0, 7.6, 0.65, { size: 13, color: C.accent5 });

  const cards = [
    ['FaUserCog', '角色可扩展', '新增角色 = 勾选权限点，无需改代码：护士长、科研人员、质检员…'],
    ['FaSitemap', '数据范围分级', '本人 → 本组 → 本科室 → 全院；工业场景对应 工位 → 产线 → 全厂'],
    ['FaKey', '权限点示例', null],
  ];
  for (let i = 0; i < cards.length; i++) {
    const [ic, t, dsc] = cards[i];
    const x = 8.6, y = 1.65 + i * 1.8, w = 4.13;
    box(s, x, y, w, 1.6, { name: 'rbac-card' });
    await iconCircle(s, ic, x + 0.25, y + 0.22, 0.5, C.accent1, 'FFFFFF');
    txt(s, t, x + 0.9, y + 0.22, w - 1.1, 0.5, { size: 16, bold: true, color: C.text2, valign: 'middle' });
    if (dsc) txt(s, dsc, x + 0.25, y + 0.82, w - 0.45, 0.7, { size: 14 });
    else txt(s, 'sample:upload   sample:read\nwall:view   review:write\nuser:manage   edge:manage', x + 0.25, y + 0.82, w - 0.45, 0.72, { size: 12, font: MONO, color: C.text2 });
  }
  s.addNotes('权限由两部分组成：角色决定“能做什么”，数据范围决定“能看谁的”。比如医护都能看历史图像，但只能看自己负责或本科室的患者。我们把每个操作定义成权限点，角色就是权限点的组合，新增角色只需要配置。');
}

async function slideFlow() {
  const s = contentSlide(SEC[3], '03  关键模块设计', '核心流程：从建档到预警的业务闭环');
  const steps = [
    ['建档', '护士创建监测任务，录入手术时间 T0，生成绑定二维码'],
    ['绑定', '患者微信扫码，阅读并同意隐私条款'],
    ['排程', '按采样策略自动生成上传时间表'],
    ['提醒', '到点推送订阅消息；超时未传通知责任护士'],
    ['上传', '引导拍照 → 质量检查 → 压缩并生成缩略图 → 入库'],
    ['查看', '监控墙实时刷新；医生查看、对比并评估'],
    ['AI 分析', '预留：输出风险评分，超过阈值自动告警'],
  ];
  const cw = 12.13 / steps.length, cy = 2.1, d = 0.8;
  arrow(s, 0.6 + cw / 2, cy, 0.6 + cw * 5.5, cy, { color: C.accent1, width: 2, noHead: true, transparency: 40 });
  arrow(s, 0.6 + cw * 5.5, cy, 0.6 + cw * 6.5, cy, { color: C.accent6, width: 2, noHead: true, dash: 'dash' });
  steps.forEach(([t, dsc], i) => {
    const cx = 0.6 + cw * (i + 0.5);
    const ai = i === steps.length - 1;
    s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: ai ? C.accent6 : C.accent1 }, line: { color: C.background1, width: 3 }, objectName: nm('step') });
    txt(s, String(i + 1), cx - d / 2, cy - d / 2, d, d, { size: 22, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    txt(s, t, cx - cw / 2 + 0.05, 2.65, cw - 0.1, 0.4, { size: 16, bold: true, color: ai ? C.accent6 : C.text2, align: 'center' });
    txt(s, dsc, cx - cw / 2 + 0.14, 3.1, cw - 0.28, 1.7, { size: 14 });
  });

  box(s, 0.6, 4.75, 12.13, 2.1, { name: 'flywheel' });
  txt(s, '数据飞轮：一期积累，二期见效', 0.9, 4.92, 8, 0.4, { size: 16, bold: true, color: C.text2 });
  const chips = ['医生评估标注', '带标签数据集', '训练 / 迭代模型', 'AI 辅助提示', '医生复核反馈'];
  const chw = 1.95, gap = (11.53 - chips.length * chw) / (chips.length - 1);
  chips.forEach((t, i) => {
    const x = 0.9 + i * (chw + gap);
    box(s, x, 5.5, chw, 0.66, { fill: { color: i === 3 ? C.accent6 : C.background1 }, r: 0.08, name: 'flywheel-chip' });
    txt(s, t, x, 5.5, chw, 0.66, { size: 14, bold: true, color: i === 3 ? C.background1 : C.text2, align: 'center', valign: 'middle' });
    if (i < chips.length - 1) arrow(s, x + chw + 0.08, 5.83, x + chw + gap - 0.08, 5.83, { color: C.accent1 });
  });
  txt(s, 'AI 只做辅助提示，诊疗判断始终由医生作出；医生的每次复核都会回流为新的标注数据', 0.9, 6.35, 11.5, 0.36, { size: 12, italic: true, color: C.accent5 });
  s.addNotes('一条完整的业务链：护士建档时录入手术结束时间 T0，系统据此排出整张采样时间表；患者按提醒拍照上传；医生在监控墙上看到最新图片并评估。第 7 步 AI 一期只留接口。关键点在下面的数据飞轮：一期医生的每次评估（正常 / 可疑 / 异常）都会存成标注，这就是二期训练模型的数据集。');
}

async function slideSchedule() {
  const s = contentSlide(SEC[3], '03  关键模块设计', '采样策略引擎：频率可配置，不写死在代码里');
  s.addChart(pres.charts.BAR, [{ name: '应上传次数', labels: ['第1天', '第2天', '第3天', '第4天', '第5天', '第6天', '第7天'], values: [24, 24, 24, 6, 6, 6, 6] }], {
    x: 0.6, y: 1.65, w: 6.3, h: 3.3, barDir: 'col', barGapWidthPct: 55,
    chartColors: [HEX.accent2, HEX.accent2, HEX.accent2, HEX.accent1, HEX.accent1, HEX.accent1, HEX.accent1],
    showTitle: true, title: '术后每日应上传次数', titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: '+mn-lt',
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: HEX.dk1, dataLabelFontSize: 12, dataLabelFontFace: '+mn-lt',
    catAxisLabelColor: HEX.accent5, catAxisLabelFontSize: 12, catAxisLabelFontFace: '+mn-lt', catAxisLineShow: false,
    valAxisHidden: true, valAxisMaxVal: 28, valAxisMinVal: 0, valGridLine: { color: 'E3E9EC', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: false, objectName: nm('chart-daily'),
  });
  [[C.accent2, '术后 0–72 h：每 1 小时一次（每天 24 次）'], [C.accent1, '72 h 之后：每 4 小时一次（每天 6 次）']].forEach(([col, t], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.85, y: 5.1 + i * 0.36 + 0.08, w: 0.16, h: 0.16, fill: { color: col }, objectName: nm('legend-key') });
    txt(s, t, 1.12, 5.1 + i * 0.36, 5.6, 0.32, { size: 13, valign: 'middle' });
  });
  await iconCircle(s, 'FaBell', 0.6, 6.0, 0.48, C.accent3, HEX.dk2);
  txt(s, [
    { text: '提醒通道：', options: { bold: true, color: C.text2 } },
    { text: '微信订阅消息为一次性授权 → 每次上传成功后引导订阅下一次；备用短信 / 护士站大屏' },
  ], 1.25, 5.92, 5.65, 0.9, { size: 14, valign: 'middle' });

  const X = 7.25, W = 5.48;
  txt(s, '策略配置示例（领域包内的 JSON）', X, 1.65, W, 0.36, { size: 14, bold: true, color: C.text2 });
  box(s, X, 2.05, W, 2.25, { fill: { color: C.text2 }, r: 0.08, name: 'code' });
  const code = [
    '{',
    '  "policy": "FLAP_DEFAULT",',
    '  "anchor": "surgeryEndTime",',
    '  "rules": [',
    '    { "from": "0h",  "to": "72h", "every": "1h" },',
    '    { "from": "72h", "to": "end", "every": "4h" }',
    '  ],',
    '  "graceMinutes": 15,',
    '  "escalate": ["patient", "nurse", "doctor"]',
    '}',
  ].join('\n');
  txt(s, code, X + 0.2, 2.17, W - 0.4, 2.05, { size: 11, font: MONO, color: C.background2, lsm: 1.12 });
  const rules = [
    ['FaClock', '时间窗与升级：', '允许 ±15 分钟；超时依次提醒患者 → 责任护士 → 主治医生'],
    ['FaSlidersH', '个体化调整：', '发现可疑时，医生可将单个患者临时改为每 30 分钟一次'],
    ['FaIndustry', '工业复用：', '同一引擎按节拍、班次或批次触发采集任务'],
  ];
  for (let i = 0; i < rules.length; i++) {
    const [ic, t, dsc] = rules[i];
    const y = 4.55 + i * 0.76;
    await iconCircle(s, ic, X, y + 0.06, 0.44, C.accent1, 'FFFFFF');
    txt(s, [{ text: t, options: { bold: true, color: C.text2 } }, { text: dsc }], X + 0.6, y, W - 0.6, 0.62, { size: 14, valign: 'middle' });
  }
  s.addNotes('老师给的规则是：术后前 72 小时每小时一次，之后每 4 小时一次。前三天每天 24 次，从第四天起每天 6 次。我们不把这个规则写在代码里，而是写成右边这样的策略配置，医生可以对单个患者临时加密。提醒用微信订阅消息，但它是一次性授权，所以每次上传后要引导患者订阅下一次，超时还要升级通知护士。');
}

async function slideWall() {
  const s = contentSlide(SEC[3], '03  关键模块设计', '实时监控墙：像车辆监控一样一屏看全病区');
  box(s, 0.6, 1.65, 7.5, 5.2, { fill: { color: C.text2 }, r: 0.12, name: 'wall' });
  txt(s, [
    { text: '骨科二病区 · 监控墙', options: { bold: true, color: C.background1, fontSize: 14 } },
    { text: '   在监 12 人', options: { color: C.background2, fontSize: 12 } },
  ], 0.85, 1.8, 4.0, 0.36, { size: 12, valign: 'middle' });
  const ST = {
    ok: { c: C.accent4, n: '正常' },
    near: { c: C.accent3, n: '临近' },
    late: { c: C.accent2, n: '超时' },
    ai: { c: C.accent6, n: 'AI 预警' },
  };
  const legendRuns = [];
  [['ok', 8], ['near', 2], ['late', 1], ['ai', 1]].forEach(([k, n], i) => {
    legendRuns.push({ text: '● ', options: { color: ST[k].c } });
    legendRuns.push({ text: `${ST[k].n} ${n}${i < 3 ? '   ' : ''}`, options: { color: C.background2 } });
  });
  txt(s, legendRuns, 3.55, 1.8, 4.3, 0.36, { size: 12, align: 'right', valign: 'middle' });

  const tiles = [
    ['12床 张*', 'ai', 'AI：疑似静脉危象'], ['7床 李*', 'late', '已超时 18 分钟'], ['3床 王*', 'near', '3 分钟后到期'], ['15床 赵*', 'near', '6 分钟后到期'],
    ['1床 陈*', 'ok', '下次 · 23 分钟后'], ['2床 刘*', 'ok', '下次 · 31 分钟后'], ['5床 杨*', 'ok', '下次 · 40 分钟后'], ['8床 黄*', 'ok', '下次 · 46 分钟后'],
    ['9床 周*', 'ok', '下次 · 1 小时后'], ['10床 吴*', 'ok', '下次 · 2 小时后'], ['11床 徐*', 'ok', '下次 · 3 小时后'], ['16床 孙*', 'ok', '下次 · 3 小时后'],
  ];
  const gx = 0.85, gy = 2.3, tw = (7.0 - 3 * 0.15) / 4, th = (4.35 - 2 * 0.15) / 3;
  for (let i = 0; i < tiles.length; i++) {
    const [bed, st, info] = tiles[i];
    const x = gx + (i % 4) * (tw + 0.15), y = gy + Math.floor(i / 4) * (th + 0.15);
    box(s, x, y, tw, th, { fill: { color: C.background1, transparency: 90 }, r: 0.06, line: { color: ST[st].c, width: st === 'ok' ? 1 : 2 }, name: 'tile' });
    s.addShape(pres.shapes.RECTANGLE, { x: x + 0.08, y: y + 0.08, w: tw - 0.16, h: 0.78, fill: { color: C.accent5, transparency: 50 }, objectName: nm('thumb') });
    s.addImage({ data: await icon('FaCamera', HEX.lt2), x: x + tw / 2 - 0.15, y: y + 0.32, w: 0.3, h: 0.3, altText: '图片缩略图占位', objectName: nm('thumb-icon') });
    s.addShape(pres.shapes.OVAL, { x: x + tw - 0.3, y: y + 0.14, w: 0.15, h: 0.15, fill: { color: ST[st].c }, objectName: nm('status-dot') });
    txt(s, bed, x + 0.1, y + 0.9, tw - 0.2, 0.22, { size: 11, bold: true, color: C.background1 });
    txt(s, info, x + 0.1, y + 1.1, tw - 0.2, 0.2, { size: 10, color: C.background2 });
  }

  const feats = [
    ['FaLayerGroup', '宫格布局', '4 / 9 / 16 宫格切换；异常患者自动置顶并高亮'],
    ['FaBolt', '实时推送', 'WebSocket 推送新图片与状态变化，秒级刷新'],
    ['FaHistory', '一键下钻', '点击宫格进入患者详情：时间轴、前后对比、评估'],
    ['FaVideo', '演进路线', '一期：准实时快照墙；二期：接入床旁摄像头实时视频（RTSP / GB28181 → WebRTC），与工业视频监控共用流媒体服务'],
  ];
  for (let i = 0; i < feats.length; i++) {
    const [ic, t, dsc] = feats[i];
    const x = 8.45, y = 1.65 + i * 1.22;
    await iconCircle(s, ic, x, y + 0.02, 0.5, i === 3 ? C.accent6 : C.accent1, 'FFFFFF');
    txt(s, t, x + 0.7, y, 3.55, 0.4, { size: 16, bold: true, color: C.text2, valign: 'middle' });
    txt(s, dsc, x + 0.7, y + 0.42, 3.58, i === 3 ? 1.35 : 0.75, { size: 14 });
  }
  s.addNotes('这就是老师说的“像车辆监控系统一样”的页面：一屏显示病区所有在监患者，每个格子是最新一张照片的缩略图和状态。颜色代表状态：绿色正常、黄色快到上传时间、红色已超时、紫色是二期 AI 给出的预警。异常的自动排到前面。一期是“准实时”的照片墙；二期如果接床旁摄像头，就升级为真正的实时视频，这部分的流媒体服务和工业视频监控是同一套。');
}

async function slideEdge() {
  const s = contentSlide(SEC[3], '03  关键模块设计', '数采端（边缘代理）：Windows / Linux 通用');
  const stages = [
    ['FaCamera', '采集驱动', ['USB 摄像头', '网络摄像头 RTSP', '工业相机 SDK', '文件夹监听', 'PLC · 传感器']],
    ['FaSlidersH', '预处理', ['模糊 / 曝光检测', '重复帧去重', '压缩 · 缩略图', '时间戳水印', '隐私区域遮挡']],
    ['FaBrain', '本地推理', ['可选模块', 'ONNX Runtime', 'OpenVINO', '弱网就地预警']],
    ['FaDatabase', '本地缓存', ['SQLite 任务队列', '本地文件暂存', '断网自动续传', '容量滚动清理']],
    ['FaCloud', '上行通道', ['HTTPS 批量上传', 'MQTT 消息', '请求签名鉴权', '失败重试退避']],
  ];
  const bw = (12.13 - 4 * 0.35) / 5, y = 1.65, bh = 2.65;
  for (let i = 0; i < stages.length; i++) {
    const [ic, t, list] = stages[i];
    const x = 0.6 + i * (bw + 0.35);
    box(s, x, y, bw, bh, { name: 'stage' });
    await iconCircle(s, ic, x + 0.2, y + 0.2, 0.55, i === 2 ? C.accent6 : C.accent3, i === 2 ? 'FFFFFF' : HEX.dk2);
    txt(s, t, x + 0.2, y + 0.85, bw - 0.4, 0.4, { size: 16, bold: true, color: C.text2 });
    txt(s, list.join('\n'), x + 0.2, y + 1.3, bw - 0.3, 1.3, { size: 13, lsm: 1.0, color: i === 2 ? C.text1 : C.text1 });
    if (i < stages.length - 1) arrow(s, x + bw + 0.04, y + bh / 2, x + bw + 0.31, y + bh / 2, { color: C.accent5 });
  }
  txt(s, '实现：Node.js 主程序 + 驱动插件（厂商 SDK 以子进程 / gRPC 封装）；部署为 Windows 服务、Linux systemd 或 Docker', 0.6, 4.4, 12.13, 0.32, { size: 12, color: C.accent5 });
  box(s, 0.6, 4.85, 12.13, 0.6, { fill: { color: C.accent1, transparency: 88 }, name: 'mgmt-band' });
  await iconCircle(s, 'FaSyncAlt', 0.78, 4.94, 0.42, C.accent1, 'FFFFFF');
  txt(s, [
    { text: '远程管理（下行）：', options: { bold: true, color: C.text2 } },
    { text: '设备注册 · 心跳上报 · 配置下发 · 远程升级（OTA）· 日志回传' },
  ], 1.35, 4.85, 11.2, 0.6, { size: 14, valign: 'middle' });
  const sc = [
    ['FaHospital', '医疗场景', '病区护士站电脑 + 床旁摄像头，服务 ICU 或行动不便的患者；与手机上传共用同一接口', C.accent1],
    ['FaIndustry', '工业场景', '产线工控机 + 工业相机 / 传感器，断网自治、本地推理，结果汇聚到中心云平台', C.accent3],
  ];
  const sw = (12.13 - 0.3) / 2;
  for (let i = 0; i < sc.length; i++) {
    const [ic, t, dsc, col] = sc[i];
    const x = 0.6 + i * (sw + 0.3);
    box(s, x, 5.65, sw, 1.2, { name: 'scenario' });
    await iconCircle(s, ic, x + 0.25, 5.85, 0.55, col, i === 1 ? HEX.dk2 : 'FFFFFF');
    txt(s, t, x + 1.0, 5.75, sw - 1.2, 0.36, { size: 16, bold: true, color: C.text2, valign: 'middle' });
    txt(s, dsc, x + 1.0, 6.12, sw - 1.2, 0.66, { size: 14 });
  }
  s.addNotes('数采端是“一物多用”的关键。它是一个跑在 Windows 或 Linux 上的常驻程序，分五段流水线：驱动层负责接各种设备，预处理做质检和压缩，可选的本地推理，本地缓存保证断网不丢数据，最后上传到云端。医院里它可以接床旁摄像头，工厂里接工业相机和 PLC，代码是同一份，只是驱动插件不同。');
}

async function slideAi() {
  const s = contentSlide(SEC[3], '03  关键模块设计', 'AI 预留接口：分析器插件化，模型可热插拔');
  const flow = [
    ['采样入库', '图片 + 元数据', false], ['发布事件', 'sample.created', false], ['AI 网关', '按领域 / 模型路由', true],
    ['模型服务', 'HTTP / gRPC', true], ['结果回写', 'inference 记录', false], ['规则引擎', '告警 · 监控墙', false],
  ];
  const fw = (12.13 - 5 * 0.3) / 6;
  flow.forEach(([t, cap, ai], i) => {
    const x = 0.6 + i * (fw + 0.3);
    box(s, x, 1.65, fw, 0.82, { fill: ai ? { color: C.accent6, transparency: 82 } : { color: C.background2 }, r: 0.08, name: 'ai-flow' });
    txt(s, t, x + 0.05, 1.72, fw - 0.1, 0.36, { size: 14, bold: true, color: ai ? C.accent6 : C.text2, align: 'center' });
    txt(s, cap, x + 0.05, 2.08, fw - 0.1, 0.3, { size: 11, color: C.accent5, align: 'center', font: i === 1 ? MONO : undefined });
    if (i < flow.length - 1) arrow(s, x + fw + 0.04, 2.06, x + fw + 0.26, 2.06, { color: C.accent5 });
  });

  box(s, 0.6, 2.75, 6.55, 3.65, { fill: { color: C.text2 }, r: 0.08, name: 'code' });
  const code = [
    '// 分析请求',
    'POST /ai/v1/analyze',
    '{',
    '  "sampleId": "S20261028-0012",',
    '  "domain":   "medical.flap",',
    '  "model":    "flap-risk@1.0",',
    '  "media":    { "type": "image", "url": "<签名URL>" },',
    '  "context":  { "hoursAfterSurgery": 26 }',
    '}',
    '// 响应（同步返回，或异步回调 /api/v1/inferences）',
    '{',
    '  "riskScore": 0.82, "level": "high",',
    '  "labels":  [{ "code": "VENOUS_CONGESTION", "conf": 0.79 }],',
    '  "metrics": { "color": 0.71, "texture": 0.40, "swelling": 0.66 },',
    '  "modelVersion": "1.0.3"',
    '}',
  ].join('\n');
  txt(s, code, 0.8, 2.88, 6.2, 3.45, { size: 11, font: MONO, color: C.background2, lsm: 1.08 });
  txt(s, 'AI 仅提供辅助提示，诊疗决策始终由医生作出', 0.6, 6.5, 6.55, 0.32, { size: 12, italic: true, color: C.accent5 });

  const cards = [
    ['FaPuzzlePiece', '统一分析器契约', '输入采样与上下文，输出风险分、标签、指标、区域与模型版本；皮瓣与烟草共用'],
    ['FaProjectDiagram', '同步 / 异步两种模式', '轻量模型同步返回；重模型经队列异步回调；边缘端可运行同一 ONNX 模型'],
    ['FaCamera', '拍摄规范决定上限', '固定距离与角度、引导框、标准色卡做颜色校正——一期就开始规范采集'],
  ];
  for (let i = 0; i < cards.length; i++) {
    const [ic, t, dsc] = cards[i];
    const x = 7.45, y = 2.75 + i * 1.27, w = 5.28;
    box(s, x, y, w, 1.15, { name: 'ai-card' });
    await iconCircle(s, ic, x + 0.2, y + 0.2, 0.5, C.accent6, 'FFFFFF');
    txt(s, t, x + 0.85, y + 0.14, w - 1.0, 0.38, { size: 15, bold: true, color: C.text2, valign: 'middle' });
    txt(s, dsc, x + 0.85, y + 0.52, w - 1.0, 0.6, { size: 14 });
  }
  s.addNotes('AI 一期不做模型，但接口现在就定下来。每张图片入库后发布一个 sample.created 事件，AI 网关按领域和模型把它路由到对应的模型服务，结果统一回写，再由规则引擎决定是否告警。输入输出格式是通用的：皮瓣模型输出颜色、纹理、肿胀等指标和风险分；烟草缺陷检测也用同一个格式。另外，模型效果的上限取决于照片质量，所以拍摄规范（固定距离、引导框、标准色卡）一期就要做。');
}

async function slideTech() {
  const s = contentSlide(SEC[4], '04  技术与工程', '技术选型：JavaScript 为主线，全栈统一');
  const rows = [
    ['用户端', 'uni-app（Vue 3）+ uni-ui', 'JavaScript 技术栈；一套代码编译微信小程序 + H5'],
    ['后台 / 大屏', 'Vue 3 + Vite + Element Plus + ECharts', '与小程序同属 Vue 生态，组件与经验复用'],
    ['业务后端', 'Node.js 20 LTS + NestJS', '模块化 + 依赖注入；守卫实现 RBAC，便于插件化'],
    ['数采端', 'Node.js + 驱动插件', 'Windows / Linux 同一份代码，可注册为系统服务'],
    ['数据库', 'PostgreSQL 16', 'JSONB 存领域扩展字段；可加 TimescaleDB 存工业时序'],
    ['缓存 / 队列', 'Redis + BullMQ', '定时提醒、超时检测、异步任务'],
    ['对象存储', 'MinIO / 腾讯云 COS', 'S3 兼容，私有化与公有云切换不改代码'],
    ['物联 / 流媒体', 'EMQX（MQTT）· ZLMediaKit', '预留：设备消息；RTSP / GB28181 转 WebRTC'],
    ['AI 服务', 'Python + FastAPI + ONNX Runtime', '预留：模型生态在 Python，经 HTTP 与主体解耦'],
    ['工程化', 'Docker Compose · GitHub Actions', '单机一键部署起步，CI 自动检查'],
  ].map(([a, b, c]) => [{ text: a, options: { bold: true, color: C.text2 } }, b, c]);
  table(s, ['层次', '选型', '理由'], rows, { x: 0.6, y: 1.65, w: 12.13, colW: [1.9, 4.2, 6.03], rowH: 0.43, size: 13 });
  txt(s, '原则：主体 JavaScript（可渐进引入 TypeScript）；只有 AI 模型服务使用 Python，并通过 HTTP 接口与主体隔离', 0.6, 6.5, 12.13, 0.32, { size: 12, color: C.accent5 });
  s.addNotes('按老师要求主体用 JavaScript。前端用 uni-app 写小程序，原因是一套代码还能出 H5，大屏和工厂里的电脑浏览器都能用；后台和大屏用 Vue 3，与小程序同一套思路。后端用 Node.js + NestJS，它天然按模块组织，正好对应我们的模块划分。数据库选 PostgreSQL 是因为 JSONB 适合领域扩展字段，将来工业时序数据也能用 TimescaleDB 扩展。AI 模型只能用 Python，但通过 HTTP 接口隔离，不影响主体。');
}

async function slideRepo() {
  const s = contentSlide(SEC[4], '04  技术与工程', '代码仓库：Monorepo 统一管理，模块边界清晰');
  box(s, 0.6, 1.65, 6.35, 5.2, { name: 'tree-card' });
  const tree = [
    ['QingSi/', ''],
    ['├─ apps/', ''],
    ['│  ├─ miniapp/', '微信小程序（三角色）'],
    ['│  ├─ console-web/', '管理后台 + 监控大屏'],
    ['│  └─ edge-agent/', '数采端（Win / Linux）'],
    ['├─ services/', ''],
    ['│  ├─ api-server/', '核心业务服务'],
    ['│  └─ ai-gateway/', 'AI 网关（预留）'],
    ['├─ ai-models/', 'Python 模型服务（二期）'],
    ['├─ packages/', ''],
    ['│  ├─ shared/', '常量 · 角色 · 错误码'],
    ['│  └─ sdk-js/', '统一 API SDK'],
    ['├─ domain-packs/', ''],
    ['│  ├─ medical-flap/', '皮瓣监测领域包'],
    ['│  └─ tobacco-online/', '烟草在线监测（三期）'],
    ['├─ deploy/', 'Docker Compose · Nginx'],
    ['└─ docs/', '架构 · 接口 · PPT'],
  ];
  const runs = [];
  tree.forEach(([p, c], i) => {
    const last = i === tree.length - 1;
    runs.push({ text: c ? p.padEnd(23) : p, options: { fontFace: MONO, color: C.text2, bold: !c, breakLine: !c && !last } });
    if (c) runs.push({ text: c, options: { color: C.accent5, breakLine: !last } });
  });
  s.addText(runs, { x: 0.9, y: 1.85, w: 5.9, h: 4.85, isTextBox: true, margin: 0, fontSize: 14, valign: 'top', lineSpacingMultiple: 1.1, fit: 'none', objectName: nm('tree') });

  const X = 7.35, W = 5.38;
  txt(s, 'api-server 模块划分（NestJS Module）', X, 1.65, W, 0.4, { size: 16, bold: true, color: C.text2 });
  const mods = [
    ['auth', '认证'], ['user', '用户'], ['rbac', '权限'], ['org', '组织'], ['subject', '监测对象'],
    ['task', '监测任务'], ['schedule', '采样排程'], ['media', '媒体'], ['sample', '采样'], ['wall', '监控墙'],
    ['alert', '告警'], ['ai-adapter', 'AI 适配'], ['edge', '设备'], ['audit', '审计'], ['domain-pack', '领域包'],
  ];
  const mw = (W - 2 * 0.15) / 3;
  mods.forEach(([en, cn], i) => {
    const x = X + (i % 3) * (mw + 0.15), y = 2.15 + Math.floor(i / 3) * 0.65;
    box(s, x, y, mw, 0.55, { fill: { color: i === 14 ? C.accent1 : C.background2 }, r: 0.06, name: 'module' });
    txt(s, [
      { text: en, options: { fontFace: MONO, bold: true, color: i === 14 ? C.background1 : C.text2, fontSize: 12, breakLine: true } },
      { text: cn, options: { color: i === 14 ? C.background1 : C.accent5, fontSize: 11 } },
    ], x + 0.05, y + 0.02, mw - 0.1, 0.51, { size: 12, align: 'center', valign: 'middle' });
  });
  txt(s, '领域包 = 场景差异的全部', X, 5.45, W, 0.4, { size: 16, bold: true, color: C.text2 });
  ['术语字典', '表单 Schema', '采样策略', '告警规则', 'AI 模型绑定', '大屏模板'].forEach((t, i) => {
    const x = X + (i % 3) * (mw + 0.15), y = 5.95 + Math.floor(i / 3) * 0.47;
    box(s, x, y, mw, 0.4, { fill: { color: C.accent1, transparency: 85 }, r: 0.06, name: 'pack-tag' });
    txt(s, t, x, y, mw, 0.4, { size: 12, color: C.text2, align: 'center', valign: 'middle' });
  });
  s.addNotes('建议用一个 Monorepo 管理所有代码：apps 放三个终端程序，services 放后端服务，packages 放前后端共享的常量和 SDK，domain-packs 放领域包。后端按右边 15 个模块组织，其中 domain-pack 模块负责加载领域包。新增烟草场景时，主要工作是在 domain-packs 下加一个目录。');
}

async function slideApi() {
  const s = contentSlide(SEC[4], '04  技术与工程', '接口契约（v1 草案）：先定接口，再并行开发');
  const M = (m) => ({ POST: C.accent2, GET: C.accent1, WS: C.accent6 }[m]);
  const ep = (parts) => ({
    text: parts.flatMap(([m, p], i) => [
      ...(i ? [{ text: ' + ', options: { fontFace: MONO, color: C.accent5 } }] : []),
      { text: `${m} `, options: { fontFace: MONO, bold: true, color: M(m) } },
      { text: p, options: { fontFace: MONO, color: C.text1 } },
    ]),
  });
  const rows = [
    ['认证', ep([['POST', '/api/v1/auth/wx-login']]), '微信 code 换取访问令牌', '全部角色'],
    ['监测任务', ep([['POST', '/api/v1/tasks']]), '建档，生成绑定二维码', '医护'],
    ['绑定', ep([['POST', '/api/v1/tasks/{id}/bind']]), '患者扫码绑定', '患者'],
    ['上传凭证', ep([['POST', '/api/v1/media/upload-token']]), '获取对象存储直传凭证', '患者 / 数采端'],
    ['采样', ep([['POST', '/api/v1/samples']]), '登记一次采样记录', '患者 / 数采端'],
    ['历史', ep([['GET', '/api/v1/subjects/{id}/samples']]), '按时间段查询历史图像', '医护'],
    ['监控墙', ep([['GET', '/api/v1/wall'], ['WS', '/ws/wall']]), '快照 + 实时事件推送', '医护 / 大屏'],
    ['评估', ep([['POST', '/api/v1/samples/{id}/reviews']]), '医生评估与标注', '医护'],
    ['用户', ep([['POST', '/api/v1/admin/users']]), '新增用户并分配角色', '管理员'],
    ['设备', ep([['POST', '/api/v1/edge/heartbeat']]), '心跳上报、拉取配置', '数采端'],
    ['AI 结果', ep([['POST', '/api/v1/inferences']]), '回写 AI 分析结果（预留）', 'AI 网关'],
  ].map(([a, b, c, d]) => [{ text: a, options: { bold: true, color: C.text2 } }, b, c, d]);
  table(s, ['模块', '方法与路径', '说明', '调用方'], rows, { x: 0.6, y: 1.65, w: 12.13, colW: [1.3, 4.55, 4.08, 2.2], rowH: 0.4, size: 13, colOpts: { 1: { fontSize: 12 } } });
  txt(s, [
    { text: '事件总线（预留）：', options: { bold: true, color: C.text2 } },
    { text: 'sample.created · task.overdue · inference.done · alert.raised', options: { fontFace: MONO } },
    { text: '  —— AI、告警、大屏按需订阅' },
  ], 0.6, 6.58, 12.13, 0.3, { size: 12, color: C.accent5 });
  s.addNotes('这是一期接口的草案，评审通过后会写成 OpenAPI 文档放在 docs 目录，前后端以文档为准并行开发。图片不经过业务服务器，而是先申请上传凭证，再直传对象存储，最后登记采样记录，这样服务器压力小、数采端和小程序走同一套接口。');
}

async function slideSecurity() {
  const s = contentSlide(SEC[4], '04  技术与工程', '数据安全与合规：医疗影像按敏感个人信息管理');
  const cards = [
    ['FaBalanceScale', '合规前置', C.accent2, ['医疗健康信息属敏感个人信息，需单独同意', '配置小程序《用户隐私保护指引》，数据境内存储', '临床使用前完成伦理审批与知情同意']],
    ['FaUserShield', '访问控制', C.accent1, ['最小权限 + 数据范围隔离，后端统一鉴权', '图片私有读，按需签发短时效签名 URL', '查看页叠加用户水印，防止截屏外传']],
    ['FaLock', '传输与存储', C.accent1, ['全链路 HTTPS；数采端上行请求签名', '对象存储加密；姓名、手机号脱敏展示', '拍摄引导只拍创面，避免拍到面部']],
    ['FaClipboardList', '审计与留存', C.accent5, ['查看、下载、导出、权限变更全部留痕', '明确数据保留期限与销毁流程', '每日备份；后续对标网络安全等级保护']],
  ];
  const w = (12.13 - 0.3) / 2, h = 2.45;
  for (let i = 0; i < cards.length; i++) {
    const [ic, t, col, list] = cards[i];
    const x = 0.6 + (i % 2) * (w + 0.3), y = 1.65 + Math.floor(i / 2) * (h + 0.3);
    box(s, x, y, w, h, { name: 'security-card' });
    await iconCircle(s, ic, x + 0.3, y + 0.28, 0.6, col, 'FFFFFF');
    txt(s, t, x + 1.1, y + 0.28, w - 1.4, 0.6, { size: 18, bold: true, color: C.text2, valign: 'middle' });
    bullets(s, list, x + 0.35, y + 1.08, w - 0.65, 1.3, { psa: 8 });
  }
  s.addNotes('患者创面照片属于医疗健康信息，按《个人信息保护法》是敏感个人信息，必须单独取得同意。技术上做到四点：合规前置、访问控制、加密传输存储、全程审计。还有一个容易忽略的细节：拍照引导框只框创面，避免拍到患者面部。');
}

async function slideDeploy() {
  const s = contentSlide(SEC[4], '04  技术与工程', '部署形态：单机起步，云边按需扩展');
  const cols = [
    ['FaCloud', '云端单机', '开发 / 试点', C.accent1, ['Nginx（HTTPS）', 'api-server（Node.js）', 'PostgreSQL · Redis', 'MinIO 对象存储', 'Docker Compose 编排'], ['一台云服务器即可运行', '一条命令启动全部服务', '适合课题验证与小规模试点']],
    ['FaHospital', '医院私有化', '正式上线', C.accent2, ['院内应用服务器', '院内数据库与存储', '病区边缘节点', '安全网关（小程序接入）', 'HIS 对接（预留）'], ['数据不出医院内网', '满足医院信息安全要求', '小程序经安全网关访问']],
    ['FaIndustry', '工业现场', '烟草复用', C.accent3, ['中心云平台', 'MQTT Broker（EMQX）', '厂区边缘工控机集群', '工业相机 / PLC / 传感器', 'MES 对接（预留）'], ['边缘断网自治、本地推理', '样本与结果汇聚到中心', '多厂区多租户统一管理']],
  ];
  const w = (12.13 - 2 * 0.3) / 3;
  for (let i = 0; i < cols.length; i++) {
    const [ic, t, tag, col, chips, list] = cols[i];
    const x = 0.6 + i * (w + 0.3), y = 1.65;
    box(s, x, y, w, 5.2, { name: 'deploy-card' });
    await iconCircle(s, ic, x + 0.3, y + 0.3, 0.6, col, i === 2 ? HEX.dk2 : 'FFFFFF');
    txt(s, t, x + 1.08, y + 0.27, w - 1.3, 0.4, { size: 18, bold: true, color: C.text2 });
    txt(s, tag, x + 1.08, y + 0.67, w - 1.3, 0.3, { size: 12, color: C.accent5 });
    chips.forEach((c, k) => {
      const cy = y + 1.2 + k * 0.42;
      const last = k === chips.length - 1;
      box(s, x + 0.3, cy, w - 0.6, 0.35, { fill: last ? null : { color: C.background1 }, line: last ? { color: C.accent5, width: 1, dashType: 'dash' } : undefined, r: 0.06, name: 'deploy-chip' });
      txt(s, c, x + 0.3, cy, w - 0.6, 0.35, { size: 12, color: C.text2, align: 'center', valign: 'middle' });
    });
    bullets(s, list, x + 0.35, y + 3.45, w - 0.6, 1.6, { psa: 8 });
  }
  s.addNotes('部署分三种形态，代码是同一套：课题阶段一台云服务器用 Docker Compose 跑起来即可；正式在医院用时，数据要求不出院，就私有化部署到医院内网；到了工厂，就是边缘工控机集群加中心云平台。');
}

async function slideOrg() {
  const s = contentSlide(SEC[5], '05  组织与计划', '开发组织架构：小团队，角色清晰');
  // 顶层
  box(s, 4.67, 1.65, 4.0, 1.15, { fill: { color: C.text2 }, name: 'org-lead' });
  await iconCircle(s, 'FaChalkboardTeacher', 4.9, 1.92, 0.6, C.accent1, 'FFFFFF');
  txt(s, '指导老师', 5.7, 1.78, 2.85, 0.4, { size: 18, bold: true, color: C.background1 });
  txt(s, '项目负责人：方向把控、需求确认、阶段验收', 5.7, 2.18, 2.85, 0.55, { size: 13, color: C.background2 });
  // 两侧顾问
  const side = async (x, ic, col, t, dsc) => {
    box(s, x, 1.65, 3.6, 1.15, { fill: { color: C.background1 }, line: { color: col, width: 1.25, dashType: 'dash' }, name: 'org-advisor' });
    await iconCircle(s, ic, x + 0.22, 1.92, 0.6, col, 'FFFFFF');
    txt(s, t, x + 1.0, 1.78, 2.5, 0.4, { size: 16, bold: true, color: C.text2 });
    txt(s, dsc, x + 1.0, 2.18, 2.5, 0.58, { size: 13, color: C.accent5 });
  };
  await side(0.6, 'FaBrain', C.accent6, '算法同学（二期）', '模型训练、评估与部署');
  await side(9.13, 'FaStethoscope', C.accent2, '临床顾问', '合作医生 / 护士：需求、拍摄规范、评估标准与标注');
  arrow(s, 4.2, 2.225, 4.67, 2.225, { color: C.accent6, dash: 'dash', noHead: true });
  arrow(s, 8.67, 2.225, 9.13, 2.225, { color: C.accent2, dash: 'dash', noHead: true });
  // 树形连线
  const w = (12.13 - 0.3) / 2, cA = 0.6 + w / 2, cB = 0.6 + w + 0.3 + w / 2;
  arrow(s, 6.67, 2.8, 6.67, 3.1, { color: C.text2, noHead: true });
  arrow(s, cA, 3.1, cB, 3.1, { color: C.text2, noHead: true });
  arrow(s, cA, 3.1, cA, 3.38, { color: C.text2 });
  arrow(s, cB, 3.1, cB, 3.38, { color: C.text2 });
  const devs = [
    ['FaLaptopCode', '开发 A｜前端与交互负责人', 'uni-app · Vue 3 · ECharts', C.accent2, ['微信小程序三端（患者 / 医护 / 管理员）', '监控墙、Web 管理后台与大屏', '交互原型、UI 规范、拍摄引导', '前端联调与体验测试']],
    ['FaServer', '开发 B｜后端与平台负责人', 'Node.js · NestJS · PostgreSQL', C.accent1, ['核心服务、数据库与接口契约', '采样排程、提醒、权限与审计', '数采端（边缘代理）与 AI 接口预留', '部署运维、CI 与安全']],
  ];
  for (let i = 0; i < devs.length; i++) {
    const [ic, t, sub, col, list] = devs[i];
    const x = 0.6 + i * (w + 0.3), y = 3.4;
    box(s, x, y, w, 3.0, { name: 'org-dev' });
    await iconCircle(s, ic, x + 0.3, y + 0.3, 0.65, col, 'FFFFFF');
    txt(s, t, x + 1.15, y + 0.26, w - 1.4, 0.42, { size: 18, bold: true, color: C.text2 });
    txt(s, sub, x + 1.15, y + 0.68, w - 1.4, 0.3, { size: 12, color: C.accent5, font: MONO });
    bullets(s, list, x + 0.35, y + 1.2, w - 0.65, 1.7, { psa: 8 });
  }
  box(s, 3.17, 6.5, 7.0, 0.36, { fill: { color: C.accent1, transparency: 85 }, r: 0.18, name: 'shared-duty' });
  txt(s, '双人共同：需求评审 · 代码互审 · 交叉测试 · 文档维护', 3.17, 6.5, 7.0, 0.36, { size: 13, bold: true, color: C.text2, align: 'center', valign: 'middle' });
  s.addNotes('团队是“1 + 2 + 2”：指导老师把方向和验收；我们两个开发者按前端 / 后端切分，接口契约是两人之间的边界；临床顾问提供业务知识和标注标准；二期做模型时再引入算法同学。两人都要参与需求评审、互相审代码和交叉测试，避免单点。');
}

async function slideRaci() {
  const s = contentSlide(SEC[5], '05  组织与计划', '分工矩阵与协作规范');
  const rows = [
    ['需求梳理与原型', '开发 A', '临床顾问', '指导老师'],
    ['接口契约（OpenAPI）', '开发 B', '开发 A', '双人互审'],
    ['小程序三端 · 监控墙', '开发 A', '开发 B', '临床顾问'],
    ['后端服务 · 数据库', '开发 B', '开发 A', '指导老师'],
    ['数采端 · AI 接口预留', '开发 B', '开发 A', '指导老师'],
    ['拍摄规范 · 评估标准', '临床顾问', '开发 A', '指导老师'],
    ['测试 · 部署 · 文档', '双人', '—', '指导老师'],
    ['阶段验收', '指导老师', '双人', '—'],
  ].map(([a, b, c, d]) => [{ text: a, options: { bold: true, color: C.text2 } }, { text: b, options: { bold: true } }, c, d]);
  table(s, ['工作项', '主责', '协作', '评审'], rows, { x: 0.6, y: 1.65, w: 7.3, colW: [2.9, 1.4, 1.5, 1.5], rowH: 0.5, size: 14 });
  txt(s, '主责 = 对结果负责；协作 = 参与实现；评审 = 把关验收', 0.6, 6.3, 7.3, 0.32, { size: 12, color: C.accent5 });

  const X = 8.25, W = 4.48;
  txt(s, '协作规范', X, 1.65, W, 0.4, { size: 16, bold: true, color: C.text2 });
  const rules = [
    ['FaCodeBranch', 'Git 分支流程', 'feature 分支开发，PR 互审后合并'],
    ['FaSyncAlt', '两周一个迭代', '每周例会同步进度、风险与下周计划'],
    ['FaFileContract', '文档先行', '接口变更先改 docs，再改代码'],
    ['FaCheck', '质量门禁', 'ESLint + Prettier、核心单测、CI 检查'],
    ['FaTasks', '任务看板', 'GitHub Issues / Projects 跟踪任务'],
  ];
  for (let i = 0; i < rules.length; i++) {
    const [ic, t, dsc] = rules[i];
    const y = 2.15 + i * 0.93;
    await iconCircle(s, ic, X, y + 0.05, 0.48, C.accent1, 'FFFFFF');
    txt(s, t, X + 0.7, y, W - 0.7, 0.36, { size: 15, bold: true, color: C.text2 });
    txt(s, dsc, X + 0.7, y + 0.37, W - 0.7, 0.42, { size: 14 });
  }
  s.addNotes('左边是分工矩阵，每项工作只有一个主责人，避免“两个人都管等于没人管”。右边是协作规范：代码全部走 Pull Request 互审，两周一个迭代，接口变更先改文档。');
}

async function slideRoadmap() {
  const s = contentSlide(SEC[5], '05  组织与计划', '路线图：三期演进，一期约 8 周交付 MVP');
  const L = 3.3, R = 12.73, wk = (R - L) / 8;
  for (let i = 0; i < 8; i++) {
    if (i % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: L + i * wk, y: 1.65, w: wk, h: 3.75, fill: { color: C.background2 }, objectName: nm('week-band') });
    txt(s, `W${i + 1}`, L + i * wk, 1.68, wk, 0.32, { size: 12, bold: true, color: C.accent5, align: 'center', valign: 'middle' });
  }
  const tasks = [
    ['架构评审 · 原型 · 接口契约', 1, 2],
    ['工程脚手架 · 登录与角色', 2, 3],
    ['建档绑定 · 拍照上传', 3, 5],
    ['历史查看 · 监控墙', 4, 6],
    ['采样提醒 · 管理后台 · 审计', 5, 7],
    ['部署联调 · 科室试点', 7, 8],
  ];
  tasks.forEach(([t, a, b], i) => {
    const y = 2.1 + i * 0.52;
    txt(s, t, 0.6, y, 2.62, 0.42, { size: 14, valign: 'middle' });
    box(s, L + (a - 1) * wk + 0.08, y + 0.06, (b - a + 1) * wk - 0.16, 0.3, { fill: { color: i === 0 ? C.accent2 : C.accent1 }, r: 0.15, name: 'gantt-bar' });
  });
  const ms = [['M1 架构冻结', 2], ['M2 核心流程跑通', 5], ['M3 科室试点', 8]];
  ms.forEach(([t, wkEnd]) => {
    const cx = Math.min(L + wkEnd * wk, R - 0.15);
    s.addShape(pres.shapes.DIAMOND, { x: cx - 0.12, y: 5.18, w: 0.24, h: 0.24, fill: { color: C.accent2 }, objectName: nm('milestone') });
    txt(s, t, cx - 1.95, 5.15, 1.75, 0.3, { size: 12, bold: true, color: C.text2, align: 'right', valign: 'middle' });
  });
  const phases = [
    ['一期 · MVP（约 8 周）', '三角色小程序 · 监控墙 · 采样提醒', C.accent1, C.background1, C.background2],
    ['二期 · 智能化', '标注数据集 · AI 模型 · 实时视频', null, C.accent6, C.text1],
    ['三期 · 平台化复用', '抽取领域包 · 烟草在线监测试点', null, C.text2, C.text1],
  ];
  const pw = (12.13 - 2 * 0.3) / 3;
  phases.forEach(([t, dsc, fill, tc, dc], i) => {
    const x = 0.6 + i * (pw + 0.3);
    box(s, x, 5.75, pw, 1.1, { fill: fill ? { color: fill } : { color: i === 1 ? C.accent6 : C.accent3, transparency: 82 }, name: 'phase' });
    txt(s, t, x + 0.25, 5.85, pw - 0.5, 0.4, { size: 16, bold: true, color: tc });
    txt(s, dsc, x + 0.25, 6.27, pw - 0.5, 0.45, { size: 14, color: dc });
  });
  s.addNotes('一期按 8 周规划，目标是一个能在科室试用的 MVP：三角色小程序、监控墙、采样提醒。第 1–2 周就是现在这一步：架构评审、原型和接口契约。二期在积累了足够标注数据后做 AI 和实时视频；三期把领域包抽出来，到烟草在线监测场景验证复用。二期和三期的时间视数据积累和老师安排再定。');
}

async function slideRisks() {
  const s = contentSlide(SEC[5], '05  组织与计划', '主要风险与应对');
  const lv = (l) => ({ text: l, options: { bold: true, align: 'center', color: l === '高' ? C.background1 : C.text1, fill: { color: l === '高' ? C.accent2 : C.accent3 } } });
  const rows = [
    [lv('高'), '小程序医疗类目需资质；实时音视频组件仅对部分类目开放', '正式版无法上线', '课题期先用体验版；由医院主体申请类目；实时视频先放 Web 端'],
    [lv('中'), '订阅消息需用户逐次授权', '提醒漏发', '每次上传后引导订阅下次；超时升级至护士站；短信兜底'],
    [lv('高'), '照片光照、角度、距离不一致', '人工判读与 AI 精度下降', '引导框 + 标准色卡 + 自动质量检测'],
    [lv('高'), '敏感数据泄露', '合规与伦理风险', '权限隔离、加密、审计；伦理审批与知情同意'],
    [lv('中'), '两人团队人力有限', '进度延误', '严控 MVP 范围；接口先行并行开发；复用成熟组件'],
    [lv('中'), '需求变化、跨场景扩展', '返工', '领域包 + 配置驱动，变化收敛在配置层'],
  ].map((r) => r.map((c, i) => (i === 1 ? { text: c, options: { bold: true, color: C.text2 } } : c)));
  table(s, ['等级', '风险', '影响', '应对'], rows, { x: 0.6, y: 1.65, w: 12.13, colW: [0.8, 3.6, 2.2, 5.53], rowH: [0.45, 0.72, 0.72, 0.72, 0.72, 0.72, 0.72], size: 14, align: ['center', 'left', 'left', 'left'] });
  txt(s, '说明：小程序类目、订阅消息与音视频组件的开放范围以微信官方最新规则为准，立项时需再次核实', 0.6, 6.55, 12.13, 0.3, { size: 12, color: C.accent5 });
  s.addNotes('最大的风险不是技术，而是小程序的类目资质：医疗类小程序需要医疗机构资质，实时音视频组件也只对部分类目开放。所以课题阶段建议先用体验版在小范围使用，正式上线由医院作为主体申请。第二个高风险是照片质量，这直接决定医生能不能看准、模型能不能训好。');
}

async function slideClosing() {
  const s = pres.addSlide({ masterName: 'CLOSING', sectionTitle: SEC[5] });
  s.addText('结语', { placeholder: 'kicker', bullet: false });
  s.addText('待确认事项与下一步', { placeholder: 'title' });
  box(s, 0.6, 1.65, 6.55, 4.75, { fill: { color: C.background1, transparency: 92 }, name: 'confirm-card' });
  txt(s, '需要老师确认的 5 件事', 0.9, 1.85, 6, 0.45, { size: 18, bold: true, color: C.accent3 });
  const qs = [
    '小程序注册主体（个人 / 学校 / 医院）与上线方式',
    '试点科室、床位规模与参与医护人数',
    '一期是否需要床旁摄像头，还是仅患者手机拍照',
    '数据存放：公有云，还是医院内网',
    '伦理审批与患者知情同意由谁牵头',
  ];
  qs.forEach((q, i) => {
    const y = 2.5 + i * 0.75;
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.04, w: 0.42, h: 0.42, fill: { color: C.accent3 }, objectName: nm('q-no') });
    txt(s, String(i + 1), 0.9, y + 0.04, 0.42, 0.42, { size: 14, bold: true, color: C.text2, align: 'center', valign: 'middle' });
    txt(s, q, 1.5, y, 5.5, 0.5, { size: 15, color: C.background1, valign: 'middle' });
  });
  box(s, 7.45, 1.65, 5.28, 4.75, { fill: { color: C.background1, transparency: 92 }, name: 'next-card' });
  txt(s, '下一步（评审通过后 2 周内）', 7.75, 1.85, 4.8, 0.45, { size: 18, bold: true, color: C.accent3 });
  const nx = ['冻结一期范围与验收标准', '输出交互原型与 OpenAPI v1 接口文档', '搭建 Monorepo、开发环境与 CI', '与临床顾问确定拍摄规范与评估量表'];
  for (let i = 0; i < nx.length; i++) {
    const y = 2.5 + i * 0.9;
    await iconCircle(s, 'FaCheck', 7.75, y + 0.1, 0.42, C.accent1, 'FFFFFF', { scale: 0.45 });
    txt(s, nx[i], 8.35, y, 4.2, 0.62, { size: 15, color: C.background1, valign: 'middle' });
  }
  txt(s, '谢谢！欢迎老师批评指正', 0.6, 6.6, 8, 0.4, { size: 16, color: C.background2 });
  s.addNotes('最后有 5 个问题需要老师拍板，尤其是小程序用谁的主体注册、数据放在哪里，这两件事决定部署方式。评审通过后两周内，我们交付原型图、接口文档和可运行的工程框架。');
}

// ───────────────────────── 主题写入 ─────────────────────────
// pptxgenjs 只写入主题的西文字体并使用 Office 默认配色；这里把配色与中文字体写进 theme1.xml。
async function applyTheme(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const part = 'ppt/theme/theme1.xml';
  const slots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const scheme = `<a:clrScheme name="${THEME.name}">${slots.map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join('')}</a:clrScheme>`;
  let xml = await zip.file(part).async('string');
  xml = xml
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${THEME.name}"`)
    .replace(/<a:ea typeface="[^"]*"\/>/g, `<a:ea typeface="${THEME.bodyFontFace}"/>`)
    .replace(/<a:font script="Hans" typeface="[^"]*"\/>/g, `<a:font script="Hans" typeface="${THEME.bodyFontFace}"/>`);
  zip.file(part, xml);
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith('.xml')) continue;
    const bad = (await zip.file(name).async('string')).match(/<a:srgbClr val="((?![0-9A-Fa-f]{6}")[^"]*)"/);
    if (bad) throw new Error(`${name} 中出现非法颜色值 "${bad[1]}"（可能把主题色传给了只接受十六进制的选项）`);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

async function main() {
  pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.author = 'QingSi 项目组';
  pres.company = 'QingSi';
  pres.title = '皮瓣状态智能辅助监测系统 — 系统架构与开发组织架构设计';
  pres.subject = '开发前架构评审';
  C = pres.SchemeColor;
  defineLayouts();

  const builders = [
    slideCover, slideSummary, slideAgenda,
    slideBackground, slideUsers,
    slidePrinciples, slideArchitecture, slideAnalogy, slideDomainModel,
    slideMiniapp, slideRbac, slideFlow, slideSchedule, slideWall, slideEdge, slideAi,
    slideTech, slideRepo, slideApi, slideSecurity, slideDeploy,
    slideOrg, slideRaci, slideRoadmap, slideRisks, slideClosing,
  ];
  let current = null;
  for (const b of builders) {
    const sec = SECTION_OF[b.name];
    if (sec !== current) {
      pres.addSection({ title: sec });
      current = sec;
    }
    await b();
  }
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT);
  console.log(`已生成：${OUT}`);
}

const SECTION_OF = {
  slideCover: SEC[0], slideSummary: SEC[0], slideAgenda: SEC[0],
  slideBackground: SEC[1], slideUsers: SEC[1],
  slidePrinciples: SEC[2], slideArchitecture: SEC[2], slideAnalogy: SEC[2], slideDomainModel: SEC[2],
  slideMiniapp: SEC[3], slideRbac: SEC[3], slideFlow: SEC[3], slideSchedule: SEC[3], slideWall: SEC[3], slideEdge: SEC[3], slideAi: SEC[3],
  slideTech: SEC[4], slideRepo: SEC[4], slideApi: SEC[4], slideSecurity: SEC[4], slideDeploy: SEC[4],
  slideOrg: SEC[5], slideRaci: SEC[5], slideRoadmap: SEC[5], slideRisks: SEC[5], slideClosing: SEC[5],
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
