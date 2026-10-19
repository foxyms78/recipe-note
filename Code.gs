/**
 * 레시피 노트 - Google Apps Script 백엔드
 *
 * 하는 일
 *  - 단축어/웹앱에서 보낸 링크(인스타·네이버 블로그) + 캡션을 받아 재료/양념/순서로 정리해 'recipes' 탭에 저장
 *  - 네이버 블로그는 링크만으로 본문을 가져와 정리
 *  - 레시피 내용이 없는 링크는 'links' 탭(링크 모음)에 따로 기록
 *  - 레시피에도 출처 링크(url)와 출처 종류(source)를 함께 남김
 *  - 같은 게시물(링크)은 중복 저장하지 않음
 *  - 대표 사진 1장: 인스타/블로그 썸네일을 자동으로 가져오거나 직접 올린 사진을 드라이브 '레시피 노트 사진' 폴더에 보관
 *  - 스크립트 속성에 CLAUDE_API_KEY 가 있으면 AI로 정리, 없으면 규칙 기반(무료)으로 정리
 *
 * 처음 한 번: setup() 실행 → TOKEN 확인 (실행 로그 또는 시트의 '설정' 탭 B2) → 웹앱/단축어에 입력
 */

const SHEETS = {
  recipe: { name: 'recipes', headers: ['id', 'createdAt', 'url', 'title', 'category', 'ingredients', 'sauce', 'steps', 'tips',
    'caption', 'rating', 'cooked', 'memo', 'method', 'updatedAt', 'source', 'image'] },
  link: { name: 'links', headers: ['id', 'createdAt', 'url', 'title', 'category', 'memo', 'source', 'updatedAt', 'image'] }
};
const IMAGE_FOLDER = '레시피 노트 사진';
const MOBILE_UA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1';
const SOURCE_LABEL = { instagram: '인스타그램', naver: '네이버 블로그', web: '웹' };
const LIST_FIELDS = ['ingredients', 'sauce', 'steps', 'tips'];
const CATEGORIES = ['반찬', '국·찌개', '밥', '면', '제과·제빵', '디저트·간식', '다이어트', '기타'];
const AI_MODEL = 'claude-haiku-5-5'; // AI 정리에 쓸 모델 (필요하면 바꾸세요)
const TZ = 'Asia/Seoul';

/* ───────────── 초기 설정 ───────────── */

/** 시트를 열면 상단에 '🍳 레시피 노트' 메뉴가 생겨요 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍳 레시피 노트')
    .addItem('초기 설정 / 토큰 보기', 'setup')
    .addToUi();
}

/**
 * 처음 한 번 실행: recipes / links 탭을 만들고 토큰을 발급해요.
 * 토큰은 ① 실행 로그 ② '설정' 탭 B2 칸 ③ (시트 메뉴에서 실행 시) 알림창 에서 확인할 수 있어요.
 */
function setup() {
  Object.keys(SHEETS).forEach(kind => {
    const sh = sheet_(kind);
    const h = SHEETS[kind].headers;
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  }
  const token = props.getProperty('TOKEN');
  imageFolder_(); // 사진 폴더 준비 (드라이브 권한 요청)
  const ai = props.getProperty('CLAUDE_API_KEY') ? '켜짐 (AI 정리)' : '꺼짐 (무료 규칙 정리)';

  // '설정' 탭에 토큰 적어두기
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ss.getSheetByName('설정') || ss.insertSheet('설정');
  st.clear();
  st.getRange(1, 1, 5, 2).setValues([
    ['항목', '값'],
    ['TOKEN', token],
    ['AI 정리', ai],
    ['안내', '이 TOKEN 값을 웹앱 설정과 단축어의 token 칸에 똑같이 넣어주세요.'],
    ['웹 앱 URL', '배포 → 배포 관리에서 확인 (…/exec 로 끝나는 주소)']
  ]);
  st.getRange(1, 1, 1, 2).setFontWeight('bold');
  st.getRange(2, 2).setFontWeight('bold').setFontSize(14).setBackground('#FFF1C7');
  st.setColumnWidth(1, 110); st.setColumnWidth(2, 420);

  Logger.log('TOKEN: ' + token);
  Logger.log('AI 정리: ' + ai);
  Logger.log("시트의 '설정' 탭 B2 칸에서도 TOKEN을 볼 수 있어요.");
  try {
    SpreadsheetApp.getUi().alert('레시피 노트 설정 완료', 'TOKEN: ' + token + "\n\n'설정' 탭 B2 칸에도 적어뒀어요.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) { /* 편집기에서 실행하면 알림창 대신 실행 로그로 확인 */ }
}

/** 캡션 정리가 어떻게 되는지 미리 확인해 보는 함수 (실행 → 로그 확인) */
function testParse() {
  const sample = [
    '✨초간단 두부조림 레시피✨',
    '밥도둑 반찬 하나 알려드릴게요!',
    '',
    '📌재료',
    '두부 1모, 대파 1/2대, 양파 1/4개',
    '',
    '📌양념장',
    '간장 3큰술 / 설탕 1큰술 / 고춧가루 1큰술',
    '다진마늘 0.5큰술, 물 100ml',
    '',
    '📌만드는 법',
    '1. 두부는 1cm 두께로 썰어 키친타월로 물기를 빼요.',
    '2. 팬에 기름을 두르고 두부를 노릇하게 구워주세요.',
    '3. 양념장을 붓고 중약불에서 5분 졸이면 끝!',
    '',
    'tip. 두부를 먼저 구우면 부서지지 않아요',
    '',
    '저장해두고 꼭 해보세요🙏 팔로우는 사랑입니다',
    '#두부조림 #반찬 #집밥'
  ].join('\n');
  Logger.log(JSON.stringify(organize_(sample), null, 2));
}

/* ───────────── 웹 요청 ───────────── */

function doGet() {
  return json_({ ok: true, message: '레시피 노트 서버가 동작 중이에요.' });
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, message: '요청 형식이 올바르지 않아요.' });
  }
  if (!checkToken_(body.token)) return json_({ ok: false, message: '토큰이 맞지 않아요.' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'list': return json_({ ok: true, recipes: readAll_('recipe').reverse(), links: readAll_('link').reverse() });
      case 'save': return json_(saveRecipe_(body.url, body.caption, body.title, body.category));
      case 'update': return json_(updateItem_(body.kind === 'link' ? 'link' : 'recipe', body.id, body.fields || {}));
      case 'reparse': return json_(reparseRecipe_(body.id, body.caption));
      case 'promote': return json_(promoteLink_(body.id, body.caption));
      case 'delete': return json_(deleteItem_(body.kind === 'link' ? 'link' : 'recipe', body.id));
      case 'setImage': return json_(setImage_(body.kind === 'link' ? 'link' : 'recipe', body.id, body.data, body.mime));
      default: return json_({ ok: false, message: '알 수 없는 요청이에요.' });
    }
  } catch (err) {
    return json_({ ok: false, message: '오류가 났어요: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── 저장/수정/삭제 ───────────── */

function saveRecipe_(rawUrl, rawCaption, rawTitle, rawCategory) {
  const userTitle = String(rawTitle || '').trim().slice(0, 40);
  const userCat = CATEGORIES.indexOf(String(rawCategory || '').trim()) >= 0 ? String(rawCategory).trim() : '';
  const url = normalizeUrl_(resolveUrl_(rawUrl));
  const source = sourceOf_(url);
  const recipes = readAll_('recipe');
  const links = readAll_('link');

  if (url) {
    const key = urlKey_(url);
    const dupR = recipes.find(r => urlKey_(r.url) === key);
    if (dupR) return { ok: true, duplicate: true, kind: 'recipe', id: dupR.id, message: '이미 저장된 레시피예요: ' + dupR.title };
    const dupL = links.find(r => urlKey_(r.url) === key);
    if (dupL) return { ok: true, duplicate: true, kind: 'link', id: dupL.id, message: '이미 링크 모음에 있어요: ' + dupL.title };
  }

  let caption = cleanCaptionInput_(rawCaption);
  // 클립보드에 예전 캡션이 남아 있던 경우는 무시
  if (caption && recipes.some(r => r.caption && r.caption === caption)) caption = '';
  let fetchedTitle = '';
  let imageSrc = '';
  if (source === 'naver') {
    const blog = fetchNaverBlog_(url); // 블로그는 본문을 직접 가져옴
    if (blog.text) { caption = blog.text; fetchedTitle = blog.title; }
    imageSrc = blog.image || '';
  } else if (source === 'instagram') {
    const meta = fetchInstagramMeta_(url);
    if (!caption) caption = meta.caption;
    imageSrc = meta.image;
  }
  const image = imageFromUrl_(imageSrc, source === 'naver' ? 'https://blog.naver.com/' : '', 'recipe-' + Date.now());

  if (!url && !caption) return { ok: false, message: '링크나 캡션이 없어요.' };
  const now = nowStr_();

  // 레시피 내용이 없으면 링크 모음에 따로 기록
  if (!caption) {
    const t = userTitle || cleanTitle_(fetchedTitle) || '이름 없는 링크';
    const item = { id: newId_(), createdAt: now, url: url, title: t, category: userCat || guessCategory_(t, []),
      memo: '', source: source, updatedAt: now, image: image };
    appendItem_('link', item);
    return { ok: true, kind: 'link', id: item.id, item: item, message: '링크 모음에 저장했어요: ' + t };
  }

  const parsed = organize_(caption);
  const title = userTitle || (source === 'naver' && fetchedTitle ? cleanTitle_(fetchedTitle) : '') || parsed.title || '이름 없는 레시피';
  const category = userCat || (parsed.method === 'ai' ? parsed.category : guessCategory_(title, parsed.ingredients));
  const rec = {
    id: newId_(), createdAt: now, url: url, title: title, category: category,
    ingredients: parsed.ingredients, sauce: parsed.sauce, steps: parsed.steps, tips: parsed.tips,
    caption: caption, rating: '', cooked: false, memo: '', method: parsed.method, updatedAt: now, source: source, image: image
  };
  appendItem_('recipe', rec);
  const from = url ? ' · 출처: ' + (SOURCE_LABEL[source] || '웹') : '';
  return { ok: true, kind: 'recipe', id: rec.id, recipe: rec,
    message: '저장했어요: ' + title + ' (재료 ' + (rec.ingredients.length + rec.sauce.length) + '개 · ' + rec.steps.length + '단계)' + from };
}

function updateItem_(kind, id, fields) {
  const found = findRow_(kind, id);
  if (!found) return { ok: false, message: '항목을 찾지 못했어요.' };
  const allowed = kind === 'link' ? ['title', 'category', 'memo', 'url']
    : ['title', 'category', 'ingredients', 'sauce', 'steps', 'tips', 'rating', 'cooked', 'memo', 'url', 'caption'];
  const rec = found.rec;
  allowed.forEach(k => { if (k in fields) rec[k] = fields[k]; });
  if ('url' in fields) { rec.url = normalizeUrl_(rec.url); rec.source = sourceOf_(rec.url); }
  rec.updatedAt = nowStr_();
  writeRow_(kind, found.row, rec);
  return { ok: true, recipe: rec, item: rec };
}

function reparseRecipe_(id, caption) {
  const found = findRow_('recipe', id);
  if (!found) return { ok: false, message: '레시피를 찾지 못했어요.' };
  const rec = found.rec;
  if (typeof caption === 'string' && caption.trim()) rec.caption = caption.trim();
  if (!rec.caption) return { ok: false, message: '정리할 캡션이 없어요.' };
  const p = organize_(rec.caption);
  rec.ingredients = p.ingredients; rec.sauce = p.sauce; rec.steps = p.steps; rec.tips = p.tips;
  rec.method = p.method;
  rec.updatedAt = nowStr_();
  writeRow_('recipe', found.row, rec);
  return { ok: true, recipe: rec };
}

/** 링크 모음 항목을 레시피로 옮기기 (캡션을 붙여넣거나, 블로그면 본문을 다시 가져옴) */
function promoteLink_(id, caption) {
  const found = findRow_('link', id);
  if (!found) return { ok: false, message: '링크를 찾지 못했어요.' };
  const link = found.rec;
  let text = cleanCaptionInput_(caption);
  if (!text && link.source === 'naver') text = fetchNaverBlog_(link.url).text || '';
  if (!text && link.source === 'instagram') text = fetchInstagramMeta_(link.url).caption;
  if (!text) return { ok: false, message: '정리할 캡션을 붙여넣어 주세요.' };
  const p = organize_(text);
  const now = nowStr_();
  const title = link.title && link.title !== '이름 없는 링크' ? link.title : (p.title || '이름 없는 레시피');
  const rec = {
    id: newId_(), createdAt: link.createdAt || now, url: link.url, title: title, category: link.category || p.category,
    ingredients: p.ingredients, sauce: p.sauce, steps: p.steps, tips: p.tips, caption: text,
    rating: '', cooked: false, memo: link.memo || '', method: p.method, updatedAt: now, source: link.source || sourceOf_(link.url),
    image: link.image || ''
  };
  appendItem_('recipe', rec);
  found.sheet.deleteRow(found.row);
  return { ok: true, recipe: rec, message: '레시피로 정리했어요: ' + title };
}

function deleteItem_(kind, id) {
  const found = findRow_(kind, id);
  if (!found) return { ok: false, message: '항목을 찾지 못했어요.' };
  trashImage_(found.rec.image);
  found.sheet.deleteRow(found.row);
  return { ok: true };
}

/* ───────────── 정리 (AI 또는 규칙) ───────────── */

function organize_(caption) {
  const empty = { title: '', category: '기타', ingredients: [], sauce: [], steps: [], tips: [], method: 'rule' };
  if (!caption) return empty;
  const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (key) {
    try { return aiOrganize_(caption, key); } catch (err) { Logger.log('AI 정리 실패, 규칙으로 대체: ' + err.message); }
  }
  const p = parseCaption_(caption);
  p.category = guessCategory_(p.title, p.ingredients);
  p.method = 'rule';
  return p;
}

function aiOrganize_(caption, key) {
  const prompt = '다음은 인스타그램 레시피 게시물의 캡션이야. 레시피 내용만 뽑아서 JSON 하나로만 답해.\n' +
    '형식: {"title":"메뉴 이름(짧게)","category":"' + CATEGORIES.join('|') + ' 중 하나",' +
    '"ingredients":["재료와 분량"],"sauce":["양념/소스와 분량"],"steps":["조리 순서 한 단계씩"],"tips":["팁"]}\n' +
    '규칙: 캡션에 없는 내용은 지어내지 마. 분량은 캡션에 적힌 그대로 써. 해시태그, 인사말, 팔로우·저장 권유 같은 홍보 문구는 빼. ' +
    '양념이 따로 구분되지 않으면 sauce는 빈 배열로 둬. 레시피가 없으면 배열을 비워 둬.\n\n캡션:\n' + caption;
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: AI_MODEL, max_tokens: 2000, messages: [{ role: 'user', content: prompt }] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('AI 응답 ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  const data = JSON.parse(res.getContentText());
  const text = (data.content || []).map(c => c.text || '').join('');
  const m = text.match(/\{[\s\S]*\}/);
  if (!m) throw new Error('AI 응답에서 JSON을 찾지 못함');
  const o = JSON.parse(m[0]);
  const arr = v => (Array.isArray(v) ? v : []).map(s => String(s).trim()).filter(Boolean);
  return {
    title: String(o.title || '').trim().slice(0, 40),
    category: CATEGORIES.indexOf(o.category) >= 0 ? o.category : '기타',
    ingredients: arr(o.ingredients), sauce: arr(o.sauce), steps: arr(o.steps), tips: arr(o.tips),
    method: 'ai'
  };
}

/* 규칙 기반 캡션 정리 */

const EMOJI_RE = /[\p{Extended_Pictographic}\uFE0F\u200D\u20E3]/gu;
const QTY_RE = /((\d+(\.\d+)?|\d+\/\d+|한|두|세|반)\s*(큰술|작은술|스푼|숟가락|숟갈|티스푼|종이컵|컵|국자|kg|g|ml|L|cc|개|모|봉지|팩|줌|꼬집|장|알|쪽|톨|대|단|포기|캔|tbsp|tsp|ts|T|t|인분|조각|마리|공기|덩이|줄기))|약간|적당량|조금|취향껏|\d\s*[:：]\s*\d/i;
const STEP_RE = /^\s*(?:(\d{1,2})\s*(?:[.)\]:]|번\b|단계)|[①-⑳]|\d\uFE0F?\u20E3|step\s*\d+\s*[.:)]?|\[\d+\]|\(\d+\))\s*/i;
const PROMO_RE = /(팔로우|좋아요|저장해|저장하고|저장 꾹|댓글|DM|디엠|공유해|구독|문의|협찬|광고|프로필 링크|링크는|맞팔|이벤트|공구)/i;
const SECTIONS = [
  ['tips', /^(꿀?팁|tip|tips|포인트|참고)$/i],
  ['sauce', /(양념장?|소스|드레싱|양념\s*재료|소스\s*재료|양념\s*비율|소스\s*비율|비율)/],
  ['ingredients', /(재료|준비물|필요한\s*것|ingredients?)/i],
  ['steps', /(만드는\s*법|만드는\s*방법|만들기|조리\s*법|조리\s*방법|조리\s*순서|요리\s*순서|순서|레시피|방법|과정|how\s*to)/i]
];

function parseCaption_(text) {
  const out = { title: '', ingredients: [], sauce: [], steps: [], tips: [] };
  if (!text) return out;

  const lines = String(text).replace(/\r/g, '').replace(/\u2028/g, '\n').split('\n')
    .map(l => l.replace(/#[^\s#]+/g, '').trim())
    .filter(l => l && /[\p{L}\p{N}]/u.test(l));

  let section = null;
  let titleDone = false;

  lines.forEach((line, idx) => {
    const plain = line.replace(EMOJI_RE, '').trim();

    // 제목: 첫 번째 의미 있는 줄
    if (!titleDone) {
      titleDone = true;
      const header = detectHeader_(plain);
      if (!header || header.section === 'steps' && !header.inline && plain.length > 6) {
        if (!PROMO_RE.test(plain)) { out.title = cleanTitle_(plain); return; }
      }
    }

    if (PROMO_RE.test(plain) && !QTY_RE.test(plain)) return;

    // 팁 (줄 앞에 tip/팁)
    const tipM = plain.match(/^[^\p{L}\p{N}]*(꿀?팁|tip)\s*[.:：)\-]?\s*(.+)$/iu);
    if (tipM && tipM[2].length > 2) { out.tips.push(cleanItem_(tipM[2])); return; }

    // 섹션 제목
    const header = detectHeader_(plain);
    if (header) {
      section = header.section;
      if (header.inline) addItems_(out, section, header.inline);
      return;
    }

    // '~가 포인트!' 같은 줄은 팁으로
    if (/(포인트|꿀팁|비법)\s*[!.~♡♥]*$/.test(plain) && plain.length > 6) { out.tips.push(cleanItem_(plain)); return; }

    // 번호가 붙은 줄 (1️⃣ 같은 숫자 이모지도 인식하도록 원문으로 검사)
    const stepM = line.match(STEP_RE);
    if (stepM) {
      const body = cleanItem_(line.slice(stepM[0].length));
      if (!body) return;
      if ((section === 'ingredients' || section === 'sauce') && !looksLikeStep_(body)) {
        addItems_(out, section, body);
      } else {
        out.steps.push(body);
        section = 'steps';
      }
      return;
    }

    const item = cleanItem_(plain);
    if (!item) return;
    if (section === 'ingredients' || section === 'sauce') {
      if (looksLikeStep_(item) && !QTY_RE.test(item)) { out.steps.push(item); section = 'steps'; }
      else addItems_(out, section, item);
    } else if (section === 'steps') {
      if (item.length >= 3) out.steps.push(item);
    } else if (section === 'tips') {
      out.tips.push(item);
    } else {
      // 섹션 없이 나온 줄: 분량이 있으면 재료, 조리 문장이면 순서, 나머지(소개글)는 버림
      if (QTY_RE.test(item) && item.length <= 40 && !looksLikeStep_(item)) addItems_(out, 'ingredients', item);
      else if (looksLikeStep_(item) && out.ingredients.length) out.steps.push(item);
    }
  });

  ['ingredients', 'sauce', 'tips'].forEach(k => { out[k] = uniq_(out[k]); });
  return out;
}

function detectHeader_(plain) {
  let head = plain;
  let inline = '';
  const colon = plain.search(/[:：]/);
  if (colon > 0 && colon <= 14 && !/\d\s*[:：]\s*\d/.test(plain.slice(Math.max(0, colon - 2), colon + 3))) {
    head = plain.slice(0, colon);
    inline = plain.slice(colon + 1).trim();
  }
  const core = head.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '').replace(/[\[\]【】()（）<>「」『』]/g, '').trim();
  if (!core || core.length > 14) return null;
  if (!inline && (QTY_RE.test(core) || STEP_RE.test(plain) || looksLikeStep_(core))) return null;
  for (const [name, re] of SECTIONS) {
    if (re.test(core)) return { section: name, inline: inline };
  }
  return null;
}

function addItems_(out, section, text) {
  const target = section === 'sauce' ? out.sauce : section === 'tips' ? out.tips : section === 'steps' ? out.steps : out.ingredients;
  if (section === 'steps' || section === 'tips') { const t = cleanItem_(text); if (t) target.push(t); return; }
  text.split(/,|，|、|·|ㆍ|\s\/\s|(?<!\d)\/(?!\d)|\s\+\s/).forEach(part => {
    const t = cleanItem_(part);
    if (t) target.push(t);
  });
}

function looksLikeStep_(s) {
  return s.length > 28 ||
    /(다|요|기|세요|줘요|줍니다|니다|해요|하기|끝)[.!~♡♥\s]*$/.test(s) ||
    /(넣어|넣고|볶아|볶고|끓여|끓이|구워|굽고|썰어|썰고|섞어|섞고|올려|부어|버무려|졸여|데쳐|삶아|익혀|담아|뿌려)/.test(s);
}

function cleanItem_(s) {
  return String(s).replace(EMOJI_RE, '')
    .replace(/^[\s\-–—•·▪▫◾◽■□●○◆◇▶▷►✓✔︎*+>~]+/, '')
    .replace(/[\s,，.。]+$/, '')
    .trim();
}

function cleanTitle_(s) {
  let t = s.replace(EMOJI_RE, '').replace(/[\[\]【】「」『』<>]/g, ' ').replace(/\s+/g, ' ').trim();
  t = t.replace(/\s+(레시피|만드는\s*법|만들기)\s*[!.~]*$/, '').trim() || t;
  return t.length > 40 ? t.slice(0, 40) + '…' : t;
}

function guessCategory_(title, ingredients) {
  const t = (title || '') + ' ' + (ingredients || []).slice(0, 3).join(' ');
  const rules = [
    ['다이어트', /(다이어트|저탄|키토|저칼로리|샐러드|닭가슴살|오트밀|두부면)/],
    ['제과·제빵', /(빵|케이크|쿠키|머핀|브라우니|스콘|파이|타르트|마들렌|휘낭시에|베이글|식빵|치아바타|포카치아|카스텔라|롤케이크|파운드|마카롱|크로플|와플|에그타르트|박력분|강력분|중력분|이스트|베이킹)/],
    ['디저트·간식', /(디저트|간식|아이스|푸딩|떡|토스트|라떼|에이드|젤리|양갱|청|그래놀라)/],
    ['면', /(면|국수|파스타|우동|라면|짜장|짬뽕|소바|쌀국수)/],
    ['밥', /(밥|덮밥|김밥|주먹밥|리조또|죽|오므라이스|카레)/],
    ['국·찌개', /(국|찌개|탕|전골|수프|스프)(\s|$)/],
    ['반찬', /(조림|무침|볶음|나물|장아찌|전|구이|김치|반찬|계란말이|장조림|겉절이|찜)/]
  ];
  for (const [name, re] of rules) if (re.test(t)) return name;
  return '기타';
}

/* ───────────── 인스타 링크/캡션 ───────────── */

function normalizeUrl_(raw) {
  if (!raw) return '';
  const m = String(raw).match(/https?:\/\/[^\s"'<>]+/);
  if (!m) return '';
  const u = m[0];
  const ig = u.match(/instagram\.com\/(?:[^\/]+\/)?(p|reel|reels|tv)\/([A-Za-z0-9_-]+)/);
  if (ig) return 'https://www.instagram.com/' + (ig[1] === 'reels' ? 'reel' : ig[1]) + '/' + ig[2] + '/';
  const nb = naverIds_(u);
  if (nb) return 'https://blog.naver.com/' + nb.blogId + '/' + nb.logNo;
  return u.split('#')[0];
}

function naverIds_(u) {
  if (!/blog\.naver\.com/.test(u)) return null;
  const q = (name) => { const m = u.match(new RegExp('[?&]' + name + '=([^&#]+)')); return m ? decodeURIComponent(m[1]) : ''; };
  if (/PostView|PostList/i.test(u) && q('blogId') && q('logNo')) return { blogId: q('blogId'), logNo: q('logNo') };
  const m = u.match(/blog\.naver\.com\/([A-Za-z0-9_-]+)\/(\d+)/);
  return m ? { blogId: m[1], logNo: m[2] } : null;
}

function urlKey_(url) {
  if (!url) return '';
  const ig = String(url).match(/instagram\.com\/(?:[^\/]+\/)?(?:p|reel|reels|tv)\/([A-Za-z0-9_-]+)/);
  if (ig) return 'ig:' + ig[1];
  const nb = naverIds_(String(url));
  if (nb) return 'nb:' + nb.blogId + '/' + nb.logNo;
  return String(url).replace(/[?#].*$/, '').replace(/\/$/, '');
}

function sourceOf_(url) {
  if (/instagram\.com/.test(url || '')) return 'instagram';
  if (/blog\.naver\.com/.test(url || '')) return 'naver';
  return url ? 'web' : '';
}

/** naver.me 같은 단축 링크, 인스타 /share/ 링크를 실제 주소로 풀기 */
function resolveUrl_(raw) {
  const m = String(raw || '').match(/https?:\/\/[^\s"'<>]+/);
  if (!m) return '';
  let u = m[0];
  if (!/(naver\.me\/|instagram\.com\/share\/)/.test(u)) return u;
  try {
    for (let i = 0; i < 5; i++) {
      const res = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148' } });
      const code = res.getResponseCode();
      const h = res.getHeaders();
      const loc = h['Location'] || h['location'];
      if (code < 300 || code >= 400 || !loc) break;
      u = /^https?:/.test(loc) ? loc : u.replace(/^(https?:\/\/[^\/]+).*$/, '$1') + loc;
      if (/blog\.naver\.com|instagram\.com\/(p|reel|tv)\//.test(u)) break;
    }
  } catch (err) { /* 풀지 못하면 원래 주소 사용 */ }
  return u;
}

/** 네이버 블로그 본문 가져오기 → { title, text } */
function fetchNaverBlog_(url) {
  const ids = naverIds_(url);
  if (!ids) return {};
  try {
    const res = UrlFetchApp.fetch('https://m.blog.naver.com/PostView.naver?blogId=' + encodeURIComponent(ids.blogId) + '&logNo=' + ids.logNo, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1',
        'Accept-Language': 'ko-KR,ko;q=0.9' }
    });
    if (res.getResponseCode() !== 200) return {};
    return parseNaverHtml_(res.getContentText());
  } catch (err) {
    return {};
  }
}

function parseNaverHtml_(html) {
  const meta = (prop) => {
    const m = html.match(new RegExp('<meta[^>]+property="' + prop + '"[^>]+content="([^"]*)"', 'i')) ||
      html.match(new RegExp('<meta[^>]+content="([^"]*)"[^>]+property="' + prop + '"', 'i'));
    return m ? decodeEntities_(m[1]).trim() : '';
  };
  let title = meta('og:title');
  const tm = html.match(/class="[^"]*se-title-text[^"]*"[\s\S]*?<\/div>/);
  if (tm) title = htmlToText_(tm[0].slice(tm[0].indexOf('>') + 1)).split('\n').map(s => s.trim()).filter(Boolean).join(' ') || title;

  let body = '';
  const start = html.indexOf('se-main-container');
  if (start >= 0) {
    // 스마트에디터 ONE: 문단(p.se-text-paragraph) 단위로 추출 (표 안의 재료 칸도 포함)
    const part = html.slice(start);
    const paras = [];
    const re = /<p[^>]*class="[^"]*se-text-paragraph[^"]*"[^>]*>([\s\S]*?)<\/p>/g;
    let m;
    while ((m = re.exec(part))) paras.push(htmlToText_(m[1]).replace(/\n+/g, ' ').trim());
    body = paras.join('\n');
  } else {
    // 예전 에디터
    const old = html.match(/<div[^>]+(?:id="postViewArea"|class="[^"]*post_ct[^"]*")[^>]*>([\s\S]*?)<\/div>\s*<\/div>/);
    if (old) body = htmlToText_(old[1]);
  }
  body = body.replace(/[​ ]/g, ' ').split('\n').map(s => s.replace(/\s+/g, ' ').trim())
    .filter((s, i, a) => s || (a[i - 1] && a[i - 1].trim())).join('\n').trim();
  if (body.length > 20000) body = body.slice(0, 20000);
  return { title: title, text: body.length >= 15 ? body : '', image: meta('og:image') };
}

function htmlToText_(h) {
  return decodeEntities_(String(h)
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|tr|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, ''));
}

function cleanCaptionInput_(raw) {
  const s = String(raw || '').trim();
  if (s.length < 15) return '';
  if (/^https?:\/\/\S+$/.test(s)) return '';
  return s;
}

/** 인스타 게시물 미리보기 정보(og:description / og:image) 가져오기 — 인스타가 막으면 빈 값 */
function fetchInstagramMeta_(url) {
  const out = { caption: '', image: '' };
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'facebookexternalhit/1.1', 'Accept-Language': 'ko-KR,ko;q=0.9' }
    });
    if (res.getResponseCode() !== 200) return out;
    const html = res.getContentText();
    let d = metaContent_(html, 'og:description');
    const q = d.indexOf(': "');
    if (q >= 0) d = d.slice(q + 3).replace(/"\s*\.?\s*$/, '');
    d = d.trim();
    out.caption = d.length >= 15 ? d : '';
    out.image = metaContent_(html, 'og:image');
  } catch (err) { /* 무시 */ }
  return out;
}

function metaContent_(html, prop) {
  const m = html.match(new RegExp('<meta[^>]+property="' + prop + '"[^>]+content="([^"]*)"', 'i')) ||
    html.match(new RegExp('<meta[^>]+content="([^"]*)"[^>]+property="' + prop + '"', 'i'));
  return m ? decodeEntities_(m[1]).trim() : '';
}

/* ───────────── 대표 사진 (구글 드라이브) ───────────── */

function imageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('IMAGE_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) { /* 지워졌으면 새로 만듦 */ } }
  const it = DriveApp.getFoldersByName(IMAGE_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(IMAGE_FOLDER);
  props.setProperty('IMAGE_FOLDER_ID', folder.getId());
  return folder;
}

/** 사진을 드라이브에 저장하고, 웹앱에서 보여줄 주소를 돌려줌 */
function storeImage_(blob, name) {
  const file = imageFolder_().createFile(blob.setName(name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // 링크가 있어야 웹앱에서 보여요
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

/** 인스타·블로그 썸네일 주소에서 사진을 받아 드라이브에 저장 (실패하면 빈 값) */
function imageFromUrl_(imgUrl, referer, name) {
  if (!imgUrl) return '';
  try {
    const headers = { 'User-Agent': MOBILE_UA };
    if (referer) headers.Referer = referer;
    const res = UrlFetchApp.fetch(imgUrl, { muteHttpExceptions: true, followRedirects: true, headers: headers });
    if (res.getResponseCode() !== 200) return '';
    const blob = res.getBlob();
    if (!/^image\//.test(blob.getContentType() || '')) return '';
    if (blob.getBytes().length > 8 * 1024 * 1024) return '';
    const ext = (blob.getContentType().split('/')[1] || 'jpg').replace('jpeg', 'jpg');
    return storeImage_(blob, name + '.' + ext);
  } catch (err) {
    return '';
  }
}

function trashImage_(imageUrl) {
  const m = String(imageUrl || '').match(/[?&]id=([^&]+)/);
  if (!m) return;
  try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (err) { /* 이미 없음 */ }
}

/** 웹앱에서 직접 올린 사진 저장 (data 가 비어 있으면 사진 삭제) */
function setImage_(kind, id, data, mime) {
  const found = findRow_(kind, id);
  if (!found) return { ok: false, message: '항목을 찾지 못했어요.' };
  let url = '';
  if (data) {
    const bytes = Utilities.base64Decode(String(data));
    if (bytes.length > 6 * 1024 * 1024) return { ok: false, message: '사진이 너무 커요.' };
    url = storeImage_(Utilities.newBlob(bytes, mime || 'image/jpeg'), kind + '-' + id + '-' + Date.now() + '.jpg');
  }
  trashImage_(found.rec.image);
  found.rec.image = url;
  found.rec.updatedAt = nowStr_();
  writeRow_(kind, found.row, found.rec);
  return { ok: true, recipe: found.rec, item: found.rec, message: url ? '사진을 저장했어요' : '사진을 지웠어요' };
}

function decodeEntities_(s) {
  return s.replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

/* ───────────── 시트 입출력 ───────────── */

function sheet_(kind) {
  const def = SHEETS[kind || 'recipe'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_(kind) {
  const sh = sheet_(kind);
  const h = SHEETS[kind].headers;
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, h.length).getValues().map(row => fromRow_(h, row)).filter(r => r.id);
}

function findRow_(kind, id) {
  if (!id) return null;
  const sh = sheet_(kind);
  const h = SHEETS[kind].headers;
  const n = sh.getLastRow();
  if (n < 2) return null;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      return { sheet: sh, row: i + 2, rec: fromRow_(h, sh.getRange(i + 2, 1, 1, h.length).getValues()[0]) };
    }
  }
  return null;
}

function appendItem_(kind, rec) {
  sheet_(kind).appendRow(toRow_(SHEETS[kind].headers, rec));
}

function writeRow_(kind, row, rec) {
  const h = SHEETS[kind].headers;
  sheet_(kind).getRange(row, 1, 1, h.length).setValues([toRow_(h, rec)]);
}

function toRow_(headers, rec) {
  return headers.map(h => {
    const v = rec[h];
    if (LIST_FIELDS.indexOf(h) >= 0) return (Array.isArray(v) ? v : String(v || '').split('\n')).filter(Boolean).join('\n');
    if (h === 'cooked') return v === true || v === 'TRUE' || v === 'true';
    if (h === 'createdAt' || h === 'updatedAt') return "'" + String(v || '');
    return v === undefined || v === null ? '' : v;
  });
}

function fromRow_(headers, row) {
  const rec = {};
  headers.forEach((h, i) => {
    let v = row[i];
    if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
    if (LIST_FIELDS.indexOf(h) >= 0) v = String(v || '').split('\n').map(s => s.trim()).filter(Boolean);
    else if (h === 'cooked') v = v === true || v === 'TRUE';
    else if (h === 'rating') v = v === '' ? '' : Number(v);
    else v = String(v === null || v === undefined ? '' : v);
    rec[h] = v;
  });
  if ('source' in rec && !rec.source) rec.source = sourceOf_(rec.url);
  return rec;
}

/* ───────────── 공통 ───────────── */

function checkToken_(token) {
  const t = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !!t && token === t;
}

function newId_() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function nowStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

function uniq_(arr) {
  const seen = {};
  return arr.filter(x => { const k = x.replace(/\s/g, ''); if (seen[k]) return false; seen[k] = 1; return true; });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
