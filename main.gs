/************** 設定 **************/
const SUBJECT_DEBIT = '【デビットカード】ご利用のお知らせ(住信SBIネット銀行)';
const SUBJECT_DEBIT_OUT = '出金のお知らせ';
const PAYPAY_SERVICE_KEYWORD = 'サービス名：ＰａｙＰａｙ';
const TIMEZONE = 'Asia/Tokyo';

// 突合許容（ミリ秒）
const MATCH_WINDOW_MS = 1 * 60 * 1000; // ±2分

/************** 初期化 **************/
function initIfNeeded() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SPREADSHEET_ID');

  if (!sheetId) {
    const ss = SpreadsheetApp.create('kakei-bot-data');
    sheetId = ss.getId();
    props.setProperty('SPREADSHEET_ID', sheetId);

    // transactions
    const transactions = ss.getActiveSheet();
    transactions.setName('transactions');
    transactions.appendRow(['ts','merchant','amount','currency','approval','message_id','source']);

    // settings
    const settings = ss.insertSheet('settings');
    settings.appendRow(['key','value']);
    settings.appendRow(['current_balance','0']);
    settings.appendRow(['monthly_budget','0']); // 任意：月予算
    settings.appendRow(['goal_title','']);
    settings.appendRow(['goal_amount','0']);

    // paypay_events
    const paypay = ss.insertSheet('paypay_events');
    paypay.appendRow(['ts','amount','raw_text','status']);

    // bank_events
    const bank = ss.insertSheet('bank_events');
    bank.appendRow(['ts','type','message_id','status','amount','raw_text']);

    Logger.log('Spreadsheet created: ' + ss.getUrl());
  } else {
    // 既存シートが欠けてる場合に補完
    const ss = SpreadsheetApp.openById(sheetId);

    if (!ss.getSheetByName('transactions')) {
      const sh = ss.insertSheet('transactions');
      sh.appendRow(['ts','merchant','amount','currency','approval','message_id','source']);
    }
    if (!ss.getSheetByName('settings')) {
      const sh = ss.insertSheet('settings');
      sh.appendRow(['key','value']);
      sh.appendRow(['current_balance','0']);
      sh.appendRow(['monthly_budget','0']);
      sh.appendRow(['goal_title','']);
      sh.appendRow(['goal_amount','0']);
    }
    if (!ss.getSheetByName('paypay_events')) {
      const sh = ss.insertSheet('paypay_events');
      sh.appendRow(['ts','amount','raw_text','status']);
    }
    if (!ss.getSheetByName('bank_events')) {
      const sh = ss.insertSheet('bank_events');
      sh.appendRow(['ts','type','message_id','status','amount','raw_text']);
    }
  }
}

/************** Webhook (LINE + MacroDroid PayPay) **************/
function doPost(e) {
  initIfNeeded();

  // ✅ LINEのWebhook検証など「空POST」で落ちないためのガード
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return textOK_();
    }

    const raw = e.postData.contents;

    let body;
    try {
      body = JSON.parse(raw);
    } catch (jsonErr) {
      Logger.log('JSON parse error: ' + jsonErr + ' raw=' + raw);
      return textOK_(); // 必ず200
    }

    // --- PayPay Webhook (MacroDroid) ---
    // body: { source:"paypay", text:"...", subtext:"...", ts:"..." }
    if (body && body.source === 'paypay') {
      handlePayPayWebhook(body);
      return textOK_();
    }

    // --- LINE Webhook ---
    if (body && body.events && Array.isArray(body.events)) {
      body.events.forEach(event => {
        if (event.type === "message" && event.message && event.message.type === "text") {
          const userId = (event.source && event.source.userId) ? event.source.userId : 'unknown';
          const groupId = (event.source && event.source.groupId) ? event.source.groupId : '';
          handleCommand(event.message.text, event.replyToken, userId, groupId);
        }
      });
    }

  } catch (err) {
    Logger.log('doPost error: ' + err);
  }

  return textOK_();
}

function textOK_() {
  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}

/************** PayPay通知受付 (MacroDroid) **************/
function handlePayPayWebhook(payload) {
  Logger.log("TS raw = " + payload.ts + " / type = " + typeof payload.ts);
  const sheet = getSheet('paypay_events');

  const ts = payload.ts
  ? new Date(Number(payload.ts))
  : new Date();
  const raw = String(payload.text || '');
  const amount = extractPayPayAmount(raw);

  sheet.appendRow([formatTs(ts), amount || '', raw, 'pending']);

  // 突合実行（pendingが溜まっててもここで処理）
  matchPayPayCharges();
}

function extractPayPayAmount(text) {
  // 例: "PayPay残高を1,234円チャージしました"
  const m = text.match(/([\d,]+)\s*円/);
  if (!m) return null;
  const n = parseInt(m[1].replace(/,/g,''), 10);
  return isFinite(n) ? n : null;
}

/************** Gmail監視（トリガーで定期実行） **************/
function pollGmail() {
  initIfNeeded();

  const props = PropertiesService.getScriptProperties();
  const lastCheck = props.getProperty("lastGmailCheck");

  let debitQuery = `subject:"${SUBJECT_DEBIT}"`;
  let outQuery   = `subject:"${SUBJECT_DEBIT_OUT}"`;

  if (lastCheck) {
    const after = Math.floor(new Date(lastCheck).getTime() / 1000) + 1;
    debitQuery += ` after:${after}`;
    outQuery   += ` after:${after}`;
  } else {
    const now = Math.floor(Date.now() / 1000);
    debitQuery += ` after:${now}`;
    outQuery   += ` after:${now}`;
  }

  const debitThreads = GmailApp.search(debitQuery);
  processDebitThreads(debitThreads);

  const outThreads = GmailApp.search(outQuery);
  processPayPayOutThreads(outThreads);

  matchPayPayCharges();

  // ★ 最後に必ず保存
  props.setProperty("lastGmailCheck", new Date().toISOString());
}

/************** デビット処理 **************/
function processDebitThreads(threads) {
  const sheet = getSheet('transactions');

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const messageId = msg.getHeader('Message-ID') || ('MSG_' + msg.getId());
      if (isAlreadyProcessed(messageId)) return;

      const body = msg.getPlainBody();

      const approval = extract(body, /承認番号\s*：\s*(\d+)/);
      const datetimeStr = extract(body, /利用日時\s*：\s*([^\n]+)/);
      const merchant = extract(body, /利用加盟店\s*：\s*([^\n]+)/);
      const amountStr = extract(body, /引落金額\s*：\s*([\d,]+(?:\.\d+)?)/);

      if (!amountStr) return;

      const amount = parseFloat(amountStr.replace(/,/g,''));
      if (!isFinite(amount)) return;

      const ts = parseSbiDatetime_(datetimeStr) || msg.getDate();

      sheet.appendRow([formatTs(ts), merchant, amount, 'JPY', approval, messageId, 'debit']);

      // 残高更新（支出はマイナス）
      updateBalance(-amount);

      sendLinePush(
        `💳 デビット利用\n` +
        `妻がお買い物したよ！いつもご苦労様🧀🧀\n` +
        `${merchant}\n${amount.toLocaleString()}円\n` +
        `残高: ${Math.round(getBalance()).toLocaleString()}円`
      );
    });
  });
}

/************** PayPay出金メール処理 **************/
function processPayPayOutThreads(threads) {
  const sheet = getSheet('bank_events');

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const messageId = msg.getHeader('Message-ID') || ('MSG_' + msg.getId());
      if (isBankMailProcessed(messageId)) return;

      const body = msg.getPlainBody();
      if (!body.includes(PAYPAY_SERVICE_KEYWORD)) return;

      const ts = extractBankPayPayTime(body) || msg.getDate();

      // 金額が取れそうなら取る（取れなくてもOK、後で時刻で突合する）
      const amount = extractPayPayOutAmount_(body);

      sheet.appendRow([
        formatTs(ts),
        'paypay_debit',
        messageId,
        'pending',
        amount || '',
        body.slice(0, 2000)
      ]);
    });
  });
}

function extractPayPayOutAmount_(body) {
  // 住信メールの文面が環境で違うので、汎用的に「xxx円」っぽいものを拾う
  // ※誤爆回避のため「金額」「出金」「引落」近辺の行を優先するのが理想だが、まずは緩めで
  const m = body.match(/([\d,]+)\s*円/);
  if (!m) return null;
  const n = parseInt(m[1].replace(/,/g,''), 10);
  return isFinite(n) ? n : null;
}

function extractBankPayPayTime(body) {
  // 例: 2026年2月21日 午後11時59分
  const m = body.match(/(\d{4})年(\d{1,2})月(\d{1,2})日.*?(午前|午後)(\d{1,2})時(\d{1,2})分/);
  if (!m) return null;

  const y = parseInt(m[1],10);
  const mo = parseInt(m[2],10);
  const d = parseInt(m[3],10);
  let h = parseInt(m[5],10);
  const mi = parseInt(m[6],10);

  if (m[4] === '午後' && h < 12) h += 12;
  if (m[4] === '午前' && h === 12) h = 0;

  return new Date(y, mo-1, d, h, mi, 0);
}

/************** 突合ロジック（±1分） **************/
function matchPayPayCharges() {
  const paypay = getSheet('paypay_events');
  const bank = getSheet('bank_events');

  const pVals = paypay.getDataRange().getValues();
  const bVals = bank.getDataRange().getValues();

  // 既存pendingが無ければ即終了
  if (pVals.length <= 1 || bVals.length <= 1) return;

  for (let i=1; i<pVals.length; i++) {
    if (String(pVals[i][3]) !== 'pending') continue;

    const pTs = parseTs_(pVals[i][0]);
    const amount = parseInt(pVals[i][1] || 0, 10);

    // ✅ 金額取れない通知は突合対象にしない（誤爆防止）
    if (!pTs || !amount || !isFinite(amount)) continue;

    for (let j=1; j<bVals.length; j++) {
      if (String(bVals[j][3]) !== 'pending') continue;
      if (String(bVals[j][1]) !== 'paypay_debit') continue;

      const bTs = parseTs_(bVals[j][0]);
      if (!bTs) continue;

      const diff = Math.abs(pTs.getTime() - bTs.getTime());

      if (diff <= MATCH_WINDOW_MS) {
        // matched
        paypay.getRange(i+1, 4).setValue('matched');
        bank.getRange(j+1, 4).setValue('matched');

        // transactions へ記録（共通口座チャージ扱い）
        const tx = getSheet('transactions');
        const msgId = 'PAYPAY_' + new Date().getTime();

        tx.appendRow([
          formatTs(pTs),
          'PayPayチャージ(共通口座)',
          amount,
          'JPY',
          '',
          msgId,
          'paypay_charge_common'
        ]);

        updateBalance(-amount);

        sendLinePush(
          `💸 共通口座 → PayPayへ出金\n` +
          `金額: ${amount.toLocaleString()}円\n` +
          `共通口座残高: ${Math.round(getBalance()).toLocaleString()}円\n` +
          `夫がPayPayにチャージしたみたい。妻と何買ったの？🙃`
        );

        break;
      }
    }
  }
}

/************** コマンド（LINE） **************/
function handleCommand(text, replyToken, userId, groupId) {
  text = String(text || '').trim();

  // セッション入力（更新/入金/目標）
  const pending = getSession_(userId);

  // 1) メニュー
  if (text === 'メニュー' || text === 'menu') {
    return replyLine(replyToken, menuText_());
  }

  // 2) コマンド一覧
  if (text === 'コマンド' || text === 'help' || text === '？' || text === '?') {
    return replyLine(replyToken, menuText_());
  }

  // pending入力の処理（数字）
  if (pending && pending.type) {

  // =========================
  // 数字入力が必要なタイプ
  // =========================
  if (
    pending.type === 'update_balance' ||
    pending.type === 'deposit' ||
    pending.type === 'set_goal_amount'
  ) {

    const n = parseMoney_(text);
    if (!n) {
      return replyLine(replyToken, `数字（例: 12000 / 12,000 / １２０００）で入力してね🙏`);
    }

    if (pending.type === 'update_balance') {
      setBalance(n);
      clearSession_(userId);
      return replyLine(replyToken,
        `✅ 残高を更新したよ\n💰 ${Math.round(getBalance()).toLocaleString()}円`);
    }

    if (pending.type === 'deposit') {
      updateBalance(n);
      clearSession_(userId);
      return replyLine(replyToken,
        `✅ 入金を反映したよ\n➕ ${n.toLocaleString()}円\n💰 残高: ${Math.round(getBalance()).toLocaleString()}円`);
    }

    if (pending.type === 'set_goal_amount') {
      if (n < 1500000) {
        return replyLine(replyToken,
          `150万円未満…？\n` +
          `それ、本気？🔥\n` +
          `未来の子供が泣いている。\n` +
          `もう一度、夢の金額をどうぞ。`
        );
      }

      setSetting_('goal_amount', String(n));
      clearSession_(userId);
      const title = getSetting_('goal_title') || '目標';

      return replyLine(replyToken,
        `🎯 目標「${title}」を ${n.toLocaleString()}円 に設定した。\n` +
        `いいね、その覚悟だ。`);
    }
  }

  // =========================
  // タイトル入力（数字不要）
  // =========================
  if (pending.type === 'set_goal_title') {
    setSetting_('goal_title', text);
    setSession_(userId, { type: 'set_goal_amount' });
    return replyLine(replyToken,
      `OK！「${text}」だね。\n次に目標金額（数字）を入力してね`);
  }

  // =========================
  // 削除確認（数字不要）
  // =========================
  if (pending.type === 'confirm_goal_delete') {
    if (text === 'はい') {
      setSetting_('goal_title', '');
      setSetting_('goal_amount', '');
      clearSession_(userId);
      return replyLine(replyToken,
        `🗑 目標を削除したよ。次はもっと大きい夢を見よう。`);
    } else {
      clearSession_(userId);
      return replyLine(replyToken,
        `安心した。夢は消えなかった。`);
    }
  }

  // =========================
  // 目標モード
  // =========================
  if (pending.type === 'goal_mode') {

    if (text === '終了') {
      clearSession_(userId);
      return replyLine(replyToken,
        `目標モードを終了したよ。わからなければ「メニュー」と送信してね。`);
    }

    if (text === '追加') {
      setSession_(userId, { type: 'set_goal_title' });
      return replyLine(replyToken,
        `目標名を入力してね。`);
    }

    if (text === '確認') {

      const title = getSetting_('goal_title');
      const amount = parseInt(getSetting_('goal_amount') || '0', 10);
      const balance = Math.round(getBalance());

      if (!title || !amount) {
        return replyLine(replyToken, `まだ目標はないよ。`);
      }

      const remain = Math.max(amount - balance, 0);
      const monthlySpend = calcThisMonthSpend_();

      const hint = (monthlySpend > 0)
        ? `今月ペース（支出 ${monthlySpend.toLocaleString()}円）だと、節約で加速できるかも🧠`
        : `まだ今月の支出データが少ない。まずは記録を積み上げよう。`;

      return replyLine(replyToken,
        `🎯 目標「${title}」\n` +
        `目標: ${amount.toLocaleString()}円\n` +
        `現在: ${balance.toLocaleString()}円\n` +
        `残り: ${remain.toLocaleString()}円\n\n` +
        `${hint}`);
    }

    if (text === '削除') {

      const title = getSetting_('goal_title');
      const amount = parseInt(getSetting_('goal_amount') || '0', 10);

      if (!title || !amount) {
        return replyLine(replyToken,
          `削除できる目標はないよ。`);
      }

      setSession_(userId, { type: 'confirm_goal_delete' });

      return replyLine(replyToken,
        `⚠ 「${title}」を削除する？\n\nはい / いいえ`);
    }

    return replyLine(replyToken,
      `目標モード中だよ。\n\n追加 / 確認 / 削除 / 終了\nのどれかを送ってね。`);
  }
}

  // 3) 残高
  if (text === '残高') {
    return replyLine(replyToken, `💰 現在残高\n${Math.round(getBalance()).toLocaleString()}円`);
  }

  // 4) 更新（残高を上書き）
  if (text === '更新') {
  setSession_(userId, { type: 'update_balance' });
  return replyLine(replyToken,
    `さあ、現実と向き合う時間だ。\n\n` +
    `今の貯金額を数字で入力してね。\n` +
    `例）120000 / 120,000 / １２００００\n\n` +
    `私は味方だ。たぶん。`);
  }

  // 5) 入金（残高に加算）
  if (text === '入金') {
    setSession_(userId, { type: 'deposit' });
    return replyLine(replyToken, `今月もお仕事お疲れ様😍\n入金額を数字で入力してね💪`);
  }

  // 6) 今月（集計）
  if (text === '今月') {
    const rep = buildMonthlyReport_();
    return replyLine(replyToken, rep);
  }

  // 7) 目標メニュー
  if (text === '目標') {
    setSession_(userId, { type: 'goal_mode' });
    return replyLine(replyToken,
      `🎯 目標モードに入ったよ\n\n` +
      `追加\n` +
      `一覧\n` +
      `削除\n` +
      `終了\n\n` +
      `やりたい操作をそのまま送ってね。`
    );
  }

  // デフォルト（何も一致しない）
  return replyLine(replyToken, `「メニュー」と送るとこのグループでできること一覧が出るよ🙂`);
}

function menuText_() {
  return (
`📌 メニュー / コマンド一覧
✅ 残高：いまの残高を見る
✅ 更新：残高を上書きする（数字入力）
✅ 入金：入金を反映する（数字入力）
✅ 今月：今月の利用額まとめ
✅ 目標：目標と残り金額

👀 まずは「更新」で残高を入れてね`
  );
}

/************** 月次集計 **************/
function buildMonthlyReport_() {
  const ym = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM');
  const tx = getSheet('transactions');
  const vals = tx.getDataRange().getValues();

  let total = 0;
  const byMerchant = {};

  for (let i=1; i<vals.length; i++) {
    const ts = parseTs_(vals[i][0]);
    if (!ts) continue;

    const rowYm = Utilities.formatDate(ts, TIMEZONE, 'yyyy/MM');
    if (rowYm !== ym) continue;

    const merchant = String(vals[i][1] || '');
    const amount = Number(vals[i][2] || 0);
    if (!isFinite(amount) || amount <= 0) continue;

    total += amount;
    byMerchant[merchant] = (byMerchant[merchant] || 0) + amount;
  }

  const items = Object.entries(byMerchant)
    .sort((a,b) => b[1]-a[1])
    .slice(0, 8)
    .map(([m,a]) => `・${m}: ${Math.round(a).toLocaleString()}円`)
    .join('\n');

  const day = Utilities.formatDate(new Date(), TIMEZONE, 'd');

  return (
`📅 今月の利用まとめ（${ym}/1〜${ym}/${day}）
合計: ${Math.round(total).toLocaleString()}円

🏪 利用店（多い順・上位）
${items || '・データなし'}

💰 残高: ${Math.round(getBalance()).toLocaleString()}円`
  );
}

function calcThisMonthSpend_() {
  const ym = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM');
  const tx = getSheet('transactions');
  const vals = tx.getDataRange().getValues();
  let total = 0;

  for (let i=1; i<vals.length; i++) {
    const ts = parseTs_(vals[i][0]);
    if (!ts) continue;
    const rowYm = Utilities.formatDate(ts, TIMEZONE, 'yyyy/MM');
    if (rowYm !== ym) continue;

    const amount = Number(vals[i][2] || 0);
    if (isFinite(amount) && amount > 0) total += amount;
  }
  return Math.round(total);
}

/************** 残高管理 **************/
function updateBalance(diff) {
  setBalance(getBalance() + diff);
}

function getBalance() {
  const v = parseFloat(getSetting_('current_balance') || '0');
  return isFinite(v) ? v : 0;
}

function setBalance(val) {
  setSetting_('current_balance', String(val));
}

function getSetting_(key) {
  const sheet = getSheet('settings');
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === key) {
      return String(data[i][1] ?? '');
    }
  }
  return '';
}

function setSetting_(key, value) {
  const sheet = getSheet('settings');
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i+1, 2).setValue(value);
      return;
    }
  }
  // なければ追加
  sheet.appendRow([key, value]);
}

/************** LINE送信 **************/
function sendLinePush(message) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const groupId = PropertiesService.getScriptProperties().getProperty('GROUP_ID');

  if (!token || !groupId) {
    Logger.log('LINE config missing. Set LINE_CHANNEL_ACCESS_TOKEN and GROUP_ID in Script Properties.');
    return;
  }

  const url = 'https://api.line.me/v2/bot/message/push';

  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      to: groupId,
      messages: [{ type: 'text', text: message }]
    }),
    muteHttpExceptions: true
  });
}

function replyLine(replyToken, message) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) {
    Logger.log('LINE token missing.');
    return;
  }

  const url = 'https://api.line.me/v2/bot/message/reply';

  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: message }]
    }),
    muteHttpExceptions: true
  });
}

/************** セッション（更新/入金/目標の入力待ち） **************/
function getSession_(userId) {
  if (!userId || userId === 'unknown') return null;
  const props = PropertiesService.getScriptProperties();
  const v = props.getProperty('SESSION_' + userId);
  if (!v) return null;
  try { return JSON.parse(v); } catch (_) { return null; }
}

function setSession_(userId, obj) {
  if (!userId || userId === 'unknown') return;
  PropertiesService.getScriptProperties().setProperty('SESSION_' + userId, JSON.stringify(obj));
}

function clearSession_(userId) {
  if (!userId || userId === 'unknown') return;
  PropertiesService.getScriptProperties().deleteProperty('SESSION_' + userId);
}

/************** 共通 **************/
function getSheet(name) {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id).getSheetByName(name);
}

function extract(text, regex) {
  const match = String(text || '').match(regex);
  return match ? String(match[1]).trim() : '';
}

function isAlreadyProcessed(messageId) {
  const sheet = getSheet('transactions');
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const ids = sheet.getRange(2, 6, last-1, 1).getValues().flat().map(String);
  return ids.includes(String(messageId));
}

function isBankMailProcessed(messageId) {
  const sheet = getSheet('bank_events');
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const ids = sheet.getRange(2, 3, last-1, 1).getValues().flat().map(String);
  return ids.includes(String(messageId));
}

function formatTs(d) {
  return Utilities.formatDate(new Date(d), TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
}

function parseTs_(v) {
  // Date
  if (v instanceof Date) return v;
  // number (timestamp)
  if (typeof v === 'number') return new Date(v);
  // string "yyyy/MM/dd HH:mm:ss"
  const s = String(v || '').trim();
  if (!s) return null;

  // ISO等は Date で拾える場合がある
  const d1 = new Date(s);
  if (!isNaN(d1.getTime())) return d1;

  const m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (!m) return null;
  const y = parseInt(m[1],10);
  const mo = parseInt(m[2],10);
  const da = parseInt(m[3],10);
  const hh = parseInt(m[4],10);
  const mi = parseInt(m[5],10);
  const ss = parseInt(m[6] || '0',10);
  return new Date(y, mo-1, da, hh, mi, ss);
}

function parseSbiDatetime_(s) {
  // 住信の「利用日時」表記が環境差あるので、拾えたら拾う（拾えなければnull）
  const str = String(s || '').trim();
  if (!str) return null;

  // 例: "2026/02/21 23:59" を想定
  const m1 = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(\d{1,2}):(\d{1,2})/);
  if (m1) {
    return new Date(
      parseInt(m1[1],10),
      parseInt(m1[2],10)-1,
      parseInt(m1[3],10),
      parseInt(m1[4],10),
      parseInt(m1[5],10),
      0
    );
  }
  return null;
}

function parseMoney_(text) {
  // 半角/全角カンマ/円を許容
  const s = String(text || '')
    .replace(/[，,]/g, '')
    .replace(/円/g, '')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .trim();

  if (!/^\d+$/.test(s)) return null;
  const n = parseInt(s, 10);
  return isFinite(n) ? n : null;
}

function buildHumor_(title, main, tail) {
  // “ユーモアたっぷり”はここで統一管理（文言の気分転換もしやすい）
  const pep = [
    '家計の神は見ている…👁️',
    '未来の自分が拍手してる👏',
    '貯金は裏切らない（たぶん）🪙',
    '今日もエラい、我々🫡'
  ];
  const pick = pep[Math.floor(Math.random() * pep.length)];
  return `${title}\n${main}\n${tail}\n${pick}`;
}

/************** 今月利用額リセット **************/
function resetThisMonthOnly() {

  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('transactions');

  const vals = sheet.getDataRange().getValues();
  const tz = 'Asia/Tokyo';
  const thisYm = Utilities.formatDate(new Date(), tz, 'yyyy/MM');

  for (let i = vals.length - 1; i >= 1; i--) {
    const ts = parseTs_(vals[i][0]);
    if (!ts) continue;

    const rowYm = Utilities.formatDate(ts, tz, 'yyyy/MM');
    if (rowYm === thisYm) {
      sheet.deleteRow(i + 1);
    }
  }

  Logger.log('今月分の利用データを削除しました');
}

/************** 完全初期化 **************/
function resetAllFinanceData() {

  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = SpreadsheetApp.openById(id);

  // transactions
  const tx = ss.getSheetByName('transactions');
  if (tx.getLastRow() > 1) {
    tx.deleteRows(2, tx.getLastRow() - 1);
  }

  // paypay_events
  const pay = ss.getSheetByName('paypay_events');
  if (pay.getLastRow() > 1) {
    pay.deleteRows(2, pay.getLastRow() - 1);
  }

  // bank_events
  const bank = ss.getSheetByName('bank_events');
  if (bank.getLastRow() > 1) {
    bank.deleteRows(2, bank.getLastRow() - 1);
  }

  // settings 初期化
  setBalance(0);
  setSetting_('goal_title', '');
  setSetting_('goal_amount', '0');

  // Gmailチェック時刻削除
  PropertiesService.getScriptProperties().deleteProperty("lastGmailCheck");

  Logger.log('完全初期化完了');
}












