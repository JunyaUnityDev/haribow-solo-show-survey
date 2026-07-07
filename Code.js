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
    if(payload.survey==='en'){ submitSoloInterestEn(payload); }
    else if(payload.survey==='member'){ submitMemberEstimate(payload); }
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
  sh.getRange(row,1,1,10).setValues([[
    payload.name||'', payload.source||'', payload.aware||'', payload.interest||'',
    payload.price||'', payload.student||'', payload.wish||'', payload.location||'',
    payload.stream||'', ts
  ]]);
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

