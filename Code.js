/** HARIBOW単独公演 興味関心アンケート — GAS Web App (HtmlService)
 *  team-builder/gas/Code.js の設計(doGet page分岐＋google.script.runでSheetsへappend)を参考にした独立プロジェクト。
 *  オーディション/team-builderとは無関係・完全に別のGASプロジェクト。
 *  SOLO_SS_ID は Config.js で定義（.gitignore対象・このリポジトリには含まれません。setup手順はREADME参照）。
 */

function doGet(e){
  var page=(e&&e.parameter&&e.parameter.page)||'';
  if(page==='en'){
    return HtmlService.createHtmlOutputFromFile('solo_interest_en')
      .setTitle('HARIBOW Solo Show Interest Survey')
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('solo_interest')
    .setTitle('HARIBOW単独公演 興味関心アンケート')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

/** GitHub Pages(静的サイト)からのフォーム送信を受ける口。GAS直URLのアクセスエラーを避けるため、
 *  外部配布用のUIはGitHub Pagesで配信し、送信だけここにfetch()でPOSTする。 */
function doPost(e){
  try{
    var payload=JSON.parse(e.postData.contents);
    var adminResult;
    if(payload.survey==='adminmail'){
      if(payload.secret!==ADMIN_SECRET) throw new Error('unauthorized');
      if(payload.action==='testdraft'){ testCreateDraftToJunya(); }
      else if(payload.action==='syncmaster'){ adminResult = syncMailToMaster(); }
      else if(payload.action==='draftreply'){ adminResult = createDraftReplyForThread_(payload.params.threadId, payload.params.body, payload.params.to, payload.params.subject); }
      else if(payload.action==='removetriggers'){ adminResult = listAndRemoveMailTriggers(); }
      else if(payload.action==='backfilllabels'){ adminResult = backfillMailMasterLabels(); }
      else if(payload.action==='setupsync'){ adminResult = setupSyncMasterTrigger(); }
      else { sendVenueInquiryEmailTest(); }
      return ContentService.createTextOutput(JSON.stringify({ok:true, result:adminResult})).setMimeType(ContentService.MimeType.JSON);
    }
    else if(payload.survey==='en'){ submitSoloInterestEn(payload); }
    else if(payload.survey==='member'){ submitMemberEstimate(payload); }
    else if(payload.survey==='circle'){ submitCircleInterest(payload); }
    else if(payload.survey==='group'){ submitGroupInterest(payload); }
    else if(payload.survey==='stream'){ submitStreamInterest(payload); }
    else { submitSoloInterest(payload); }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function soloSs_(){ return SpreadsheetApp.openById(SOLO_SS_ID); }

/** データブロック(startRow〜endRow)内でA列が空の最初の行を返す。ブロック下の集計式を上書きしないためのガード。 */
function firstEmptyRowInBlock_(sheet, startRow, endRow){
  var vals = sheet.getRange(startRow, 1, endRow-startRow+1, 1).getValues();
  for (var i=0; i<vals.length; i++){
    if (String(vals[i][0]).trim()==='') return startRow+i;
  }
  throw new Error('回答欄が満杯です。運営に連絡してください。');
}

/** 国内向け 興味関心アンケート回答保存 */
function submitSoloInterest(payload){
  var sh = soloSs_().getSheetByName('興味関心_国内');
  var row = firstEmptyRowInBlock_(sh, 6, 105);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange(row,1,1,12).setValues([[
    payload.name||'', payload.source||'', payload.aware||'', payload.interest||'',
    payload.price||'', payload.student||'', payload.wish||'', payload.location||'',
    payload.stream||'', payload.startTime||'', payload.childPlan||'', ts
  ]]);
  return 'ok';
}

/** 大学ダブルダッチサークル(現役生)代表による団体ヒアリング回答保存 */
function submitCircleInterest(payload){
  var sh = soloSs_().getSheetByName('興味関心_現役生サークル');
  var row = firstEmptyRowInBlock_(sh, 6, 15);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange(row,1,1,8).setValues([[
    payload.circleName||'', payload.repName||'', payload.contact||'', payload.headcount||'',
    payload.priceForVolume||'', payload.startTime||'', payload.wish||'', ts
  ]]);
  return 'ok';
}

/** 非DDサークル団体(学校/企業/ダンス・チアスタジオ/地域団体等)からの団体来場相談 回答保存 */
function submitGroupInterest(payload){
  var sh = soloSs_().getSheetByName('興味関心_団体営業');
  var row = firstEmptyRowInBlock_(sh, 6, 105);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange(row,1,1,7).setValues([[
    payload.groupName||'', payload.groupType||'', payload.repName||'', payload.contact||'',
    payload.headcount||'', payload.note||'', ts
  ]]);
  return 'ok';
}

/** 全国DDスクール・大学サークルからの配信視聴意向アンケート 回答保存 */
function submitStreamInterest(payload){
  var sh = soloSs_().getSheetByName('興味関心_配信全国DD団体');
  var row = firstEmptyRowInBlock_(sh, 6, 105);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange(row,1,1,9).setValues([[
    payload.groupName||'', payload.area||'', payload.repName||'', payload.contact||'',
    payload.wantWatch||'', payload.watchCount||'', payload.ticketCount||'', payload.price||'', payload.note||''
  ]]);
  sh.getRange(row,10).setValue(ts);
  return 'ok';
}

/** 海外向け(英語) 興味関心アンケート回答保存 */
function submitSoloInterestEn(payload){
  var sh = soloSs_().getSheetByName('興味関心_海外EN');
  var row = firstEmptyRowInBlock_(sh, 6, 55);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange(row,1,1,6).setValues([[
    payload.name||'', payload.heard||'', payload.watch||'', payload.watchRecorded||'',
    payload.location||'', ts
  ]]);
  return 'ok';
}

/** 動員見込み_出演者シートで氏名(A列・5〜35行目)が一致する行を探す。見つからなければ-1。 */
function findMemberRow_(sheet, name){
  var names = sheet.getRange(5,1,31,1).getValues();
  for (var i=0; i<names.length; i++){
    if (String(names[i][0]).trim() === String(name).trim()) return 5+i;
  }
  return -1;
}

/** 出演者本人による動員見込み回答。既存行(氏名でマッチ)にupsert。2回目送信でも上書きされるだけ。 */
function submitMemberEstimate(payload){
  var sh = soloSs_().getSheetByName('動員見込み_出演者');
  var row = findMemberRow_(sh, payload.name);
  if(row<0) throw new Error('名前が見つかりません: '+payload.name);
  sh.getRange(row,3,1,5).setValues([[
    payload.familySure||'', payload.familyMaybe||'', payload.friendsSure||'', payload.friendsMaybe||'', payload.otherCount||''
  ]]);
  return 'ok';
}

/** clasp run専用・doPostには繋がず公開エンドポイント化しない。会場問い合わせメール送信用の一時関数。 */
function sendVenueInquiryEmailTest(){
  GmailApp.sendEmail(
    'junyasankura@gmail.com',
    '【テスト】9/29大ホールご利用に関する舞台設備のお問い合わせ（HARIBOW）',
    'めぐろパーシモンホール　舞台技術ご担当者様\n\n'+
    'お世話になっております。\n'+
    '2026年9月29日に貴ホール大ホールにてイベントの開催を予定しております、HARIBOW（ダブルダッチパフォーマンスチーム）と申します。\n\n'+
    '演出内容の検討にあたり、舞台設備についていくつか確認させていただきたく、ご連絡いたしました。お手数をおかけいたしますが、ご回答いただけますと幸いです。\n\n'+
    '1. プロジェクターについて\n'+
    '　貴ホールの備品としてプロジェクターの貸出は可能でしょうか。\n'+
    '　可能な場合、輝度（ルーメン）・解像度・入力端子の種類を教えてください。\n'+
    '　舞台間口18mの幅に投影するのに十分な明るさがあるかどうかも併せてご教示いただけますと幸いです。\n\n'+
    '2. 完全暗転について\n'+
    '　消防・非常灯の規定上、完全暗転が可能な最大秒数を教えてください。\n\n'+
    '3. レーザー演出について\n'+
    '　舞台上でのレーザー演出は可能でしょうか。可能な場合、事前申請の要否・手続きを教えてください。\n\n'+
    '4. スモーク・ドライアイスの使用について\n'+
    '　備品一覧に記載のスモークマシン（ROSCO1200）・ドライアイスマシーン（A型）の使用にあたり、\n'+
    '　感知器の一時停止など事前申請の手続き・必要書類・許可にかかる日数・追加費用・申請締切を教えてください。\n\n'+
    '5. 音量の規制について\n'+
    '　本番中の音量上限に関する規制があれば教えてください。\n\n'+
    '6. 舞台監督・映像オペレーターの手配について\n'+
    '　貴ホール側で舞台監督・映像オペレーターの手配は可能でしょうか。可能な場合の費用感も教えていただけますと幸いです。\n\n'+
    'なお、紗幕・電源容量・バトン耐荷重については、貴ホールWebサイトに掲載の資料（大ホール備品一覧表・照明設備概要）にて確認できましたので、上記のみで結構です。\n\n'+
    'お忙しいところ恐れ入りますが、ご確認のほどよろしくお願いいたします。\n\n'+
    'HARIBOW\n'+
    '連絡先：info@haribow.com',
    {name: 'HARIBOW'}
  );
}

/** ===== メールマスタ同期システム(2026-07-09構築・2026-07-09 Gemini除去・案件横断基盤に再編) =====
 *  info@haribow.comの受信箱の生データを、案件横断の「HARIBOWメールマスタ」SSにコピーするだけ。
 *  Gemini等の外部AI呼び出しは一切しない。迷惑メール除外はGmail自身のカテゴリ分類(プロモーション/ソーシャル/更新情報/フォーラム)を使う。
 *  スコープ判定(単独公演の業者かどうか等)は、このマスタを読んだ上で別途(Claude等が)行う。
 *  「メール対応リスト」「除外ログ」の列定義は維持(将来、案件側の判定結果を書き込む先として使用)。 */
var MAIL_SETTINGS_SHEET_ = 'メール対応設定';
var MAIL_LIST_SHEET_ = 'メール対応リスト';
var MAIL_LIST_FIRST_DATA_ROW_ = 3;
var MAIL_LIST_LAST_DATA_ROW_ = 502;
var EXCLUDE_LOG_SHEET_ = '除外ログ';
var EXCLUDE_LOG_FIRST_DATA_ROW_ = 3;
var EXCLUDE_LOG_LAST_DATA_ROW_ = 1002;
var MAIL_SYNCED_LABEL_ = 'hb-mail-synced';
var MAIL_MASTER_SHEET_ = 'メールマスタ';
var MAIL_MASTER_FIRST_DATA_ROW_ = 3;
var MAIL_MASTER_LAST_DATA_ROW_ = 5002;

function mailMasterSs_(){ return SpreadsheetApp.openById(MAIL_MASTER_SS_ID); }

/** 手動実行専用(常時トリガーなし)。info@haribow.comの受信箱→HARIBOWメールマスタへ生データを転記するだけ。
 *  迷惑メール除外はGmail自身のカテゴリ分類を利用。AI判定・外部API呼び出しは一切なし。
 *  マスタ列: No./受信日時/送信者(メアド)/送信者名/件名/本文抜粋/Gmailラベル/スレッドID/メッセージID/処理状態
 *  既にGmail上でラベル付け(手動仕分け)されているスレッドは、そのラベル名も転記して分類の手がかりにする。 */
function syncMailToMaster(){
  var masterSh = mailMasterSs_().getSheetByName(MAIL_MASTER_SHEET_);
  var label = GmailApp.getUserLabelByName(MAIL_SYNCED_LABEL_) || GmailApp.createLabel(MAIL_SYNCED_LABEL_);
  var threads = GmailApp.search('in:inbox -category:promotions -category:social -category:updates -category:forums -label:'+MAIL_SYNCED_LABEL_, 0, 30);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var count = 0;

  threads.forEach(function(thread){
    var messages = thread.getMessages();
    var msg = messages[messages.length-1];
    var from = msg.getFrom();
    var fromMatch = from.match(/<(.+)>/);
    var fromAddr = fromMatch ? fromMatch[1] : from;
    var fromName = from.replace(/<.+>/,'').replace(/"/g,'').trim();
    var subject = msg.getSubject();
    var snippet = msg.getPlainBody().substring(0,400);
    var gmailLabels = thread.getLabels().map(function(l){ return l.getName(); })
      .filter(function(n){ return n!==MAIL_SYNCED_LABEL_; }).join(', ');

    var row = firstEmptyRowInBlock_(masterSh, MAIL_MASTER_FIRST_DATA_ROW_, MAIL_MASTER_LAST_DATA_ROW_);
    var no = row - MAIL_MASTER_FIRST_DATA_ROW_ + 1;
    masterSh.getRange(row,1,1,10).setValues([[no, ts, fromAddr, fromName, subject, snippet, gmailLabels, thread.getId(), msg.getId(), '']]);

    thread.addLabel(label);
    count++;
  });

  return count+'件をメールマスタへ転記';
}

/** 一時関数: 既存行のうちGmailラベル列(G)が空の行だけ、スレッドIDから現在のラベルを引いて補完する。
 *  Gmail側のラベル操作は行わない・重複行も作らない安全な後付け処理。 */
function backfillMailMasterLabels(){
  var masterSh = mailMasterSs_().getSheetByName(MAIL_MASTER_SHEET_);
  var lastRow = masterSh.getLastRow();
  if(lastRow < MAIL_MASTER_FIRST_DATA_ROW_) return 'no data';
  var numRows = lastRow - MAIL_MASTER_FIRST_DATA_ROW_ + 1;
  var values = masterSh.getRange(MAIL_MASTER_FIRST_DATA_ROW_, 1, numRows, 10).getValues();
  var count = 0;
  values.forEach(function(r, i){
    var row = MAIL_MASTER_FIRST_DATA_ROW_ + i;
    var existingLabel = r[6]; // G列: Gmailラベル
    var threadId = r[7];      // H列: スレッドID
    if(!threadId || existingLabel) return;
    try{
      var thread = GmailApp.getThreadById(threadId);
      if(thread){
        var labels = thread.getLabels().map(function(l){ return l.getName(); })
          .filter(function(n){ return n!==MAIL_SYNCED_LABEL_; }).join(', ');
        masterSh.getRange(row,7).setValue(labels);
        count++;
      }
    }catch(e){}
  });
  return count+'件のGmailラベルを補完';
}

/** Claudeがスレッドを読んで返信文を作った直後に、その場でGmail下書きを作る口。実送信はしない・承認チェックボックス等の中間ステップは無し。
 *  threadIdが見つかればそのスレッドへの返信下書き、無ければto/subject指定で新規下書きを作成する。 */
function createDraftReplyForThread_(threadId, body, to, subject){
  var thread = threadId ? GmailApp.getThreadById(threadId) : null;
  if(thread){
    thread.createDraftReply(body);
    return '下書き作成: '+thread.getFirstMessageSubject();
  }
  if(!to) throw new Error('threadIdが見つからず、toも未指定です');
  GmailApp.createDraft(to, 'Re: '+(subject||''), body);
  return '下書き作成(新規): '+to;
}

/** syncMailToMaster専用の30分おきトリガーを設定する。Gemini等の外部AI呼び出しはこの関数にはない(単純転記のみ)。
 *  一度だけ手動実行すればOK(既存トリガーがあれば重複作成しない)。 */
function setupSyncMasterTrigger(){
  var existing = ScriptApp.getProjectTriggers().filter(function(t){
    return t.getHandlerFunction()==='syncMailToMaster';
  });
  if(existing.length>0) return 'already exists';
  ScriptApp.newTrigger('syncMailToMaster').timeBased().everyMinutes(30).create();
  return 'created';
}

/** syncMailToMaster向けの既存トリガーを確認し、あれば削除する(常時自動化はしない方針のため)。 */
function listAndRemoveMailTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  var targets = triggers.filter(function(t){
    var fn = t.getHandlerFunction();
    return fn==='checkNewMailForScope' || fn==='syncMailToMaster';
  });
  var removed = targets.map(function(t){ return t.getHandlerFunction(); });
  targets.forEach(function(t){ ScriptApp.deleteTrigger(t); });
  return 'found/removed: '+(removed.length ? removed.join(',') : 'none');
}

/** テスト用・一時関数。junyasankura@gmail.com宛の下書きをinfo@haribow.com側で作成できるか確認する。送信はしない。 */
function testCreateDraftToJunya(){
  GmailApp.createDraft(
    'junyasankura@gmail.com',
    'Re: 【テスト】メール対応システムのドラフト作成テスト',
    'これはメール対応スコープ判定システムの、返信ドラフト作成機能のテストです。\n\n'+
    'このドラフトが正しく作成されていれば、実際の運用でも「判定→ログ記録→下書き作成→人が確認して送信」という流れが機能する見込みです。\n\n'+
    '（このメール自体はテストであり、送信はされていません。下書きフォルダに保存されているだけです）'
  );
}

