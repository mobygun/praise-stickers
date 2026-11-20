/**
 * 칭찬스티커 판 - 구글 시트 저장용 Apps Script
 *
 * 처음 한 번: 위 메뉴에서 setup 함수를 선택하고 ▶실행 → 권한 허용
 * 그 다음: 배포 → 새 배포 → 유형 "웹 앱", 실행 계정 "나", 액세스 "모든 사용자"
 * PIN 바꾸기: changePin 함수 안의 값을 바꾸고 ▶실행
 */

var DEFAULT_PIN = 'CHANGE_ME'; // 실제 PIN은 changePin으로 설정 (공개 저장소에 적지 않음)
// PIN 없이 할 수 있는 작업 (스티커 붙이기, 이미지 올리기). 나머지 수정은 PIN 필요
var OPEN_ACTIONS = ['addSticker', 'addImage'];
var DEFAULT_PEOPLE = ['장경화', '정지현', '문남현', '김건'];
var COLORS = ['#ff8a5c', '#5cb8ff', '#6cc58b', '#b48cf0', '#f06c9b', '#f2b33d', '#3fbfb0', '#8f9bff'];

var SHEETS = {
  people:   { name: '사람',   head: ['id', '이름', '색상', '완성판수', '삭제', '추가시각'] },
  stickers: { name: '스티커', head: ['id', '사람id', '사람이름', '날짜', '사유', '붙인사람', '스티커', '판번호', '기록시각'] },
  images:   { name: '이미지', head: ['id', '이름', '데이터', '추가시각'] },
  settings: { name: '설정',   head: ['항목', '값'] }
};

// ---------- 처음 설정 ----------
function setup() {
  var ss = getSS_();
  Logger.log('구글 시트 주소: ' + ss.getUrl());
  Logger.log('현재 PIN: ' + getPin_());
}

function changePin() {
  var NEW_PIN = '새PIN'; // ← 여기 값을 바꾸고 실행
  PropertiesService.getScriptProperties().setProperty('PIN', String(NEW_PIN));
  Logger.log('PIN 변경 완료: ' + NEW_PIN);
}

function getPin_() {
  var props = PropertiesService.getScriptProperties();
  var pin = props.getProperty('PIN');
  if (!pin) { pin = DEFAULT_PIN; props.setProperty('PIN', pin); }
  return pin;
}

function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SSID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 지워졌으면 새로 만듦 */ }
  }
  var ss = SpreadsheetApp.create('칭찬스티커 기록');
  Object.keys(SHEETS).forEach(function (k, i) {
    var sh = i === 0 ? ss.getSheets()[0].setName(SHEETS[k].name) : ss.insertSheet(SHEETS[k].name);
    sh.getRange(1, 1, 1, SHEETS[k].head.length).setValues([SHEETS[k].head]).setFontWeight('bold').setBackground('#fff1c9');
    sh.setFrozenRows(1);
  });
  var now = new Date();
  var ppl = sheet_(ss, 'people');
  DEFAULT_PEOPLE.forEach(function (n, i) { ppl.appendRow([newId_(), n, COLORS[i % COLORS.length], 0, '', now]); });
  sheet_(ss, 'settings').appendRow(['goal', 20]);
  props.setProperty('SSID', ss.getId());
  getPin_();
  return ss;
}

// ---------- 도우미 ----------
function sheet_(ss, key) { return ss.getSheetByName(SHEETS[key].name); }
function newId_() { return Utilities.getUuid().slice(0, 8); }
function rows_(ss, key) {
  var sh = sheet_(ss, key);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, SHEETS[key].head.length).getValues();
}
function findRow_(ss, key, id) {
  var r = rows_(ss, key);
  for (var i = 0; i < r.length; i++) if (String(r[i][0]) === String(id)) return { row: i + 2, values: r[i] };
  return null;
}
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  return String(v || '');
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readAll_(ss) {
  var people = rows_(ss, 'people').filter(function (r) { return !r[4]; }).map(function (r) {
    return { id: String(r[0]), name: String(r[1]), color: String(r[2]), boards: Number(r[3]) || 0 };
  });
  var stickers = rows_(ss, 'stickers').map(function (r) {
    return { id: String(r[0]), personId: String(r[1]), date: fmtDate_(r[3]), reason: String(r[4]),
             giver: String(r[5]), sticker: String(r[6]), board: Number(r[7]) || 0 };
  });
  var images = rows_(ss, 'images').map(function (r) { return { id: String(r[0]), name: String(r[1]), src: String(r[2]) }; });
  var settings = {};
  rows_(ss, 'settings').forEach(function (r) { settings[r[0]] = r[1]; });
  return { people: people, stickers: stickers, images: images, goal: Number(settings.goal) || 20 };
}

// ---------- 웹 API ----------
function doGet(e) {
  var ss = getSS_();
  return out_({ ok: true, data: readAll_(ss) });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  if (OPEN_ACTIONS.indexOf(body.action) < 0 && String(body.pin) !== getPin_()) return out_({ ok: false, error: 'pin' });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSS_();
    var now = new Date();
    var p, row;
    switch (body.action) {
      case 'checkPin':
        break;

      case 'addSticker':
        p = findRow_(ss, 'people', body.personId);
        if (!p) return out_({ ok: false, error: '사람을 찾을 수 없어요' });
        sheet_(ss, 'stickers').appendRow([
          newId_(), body.personId, p.values[1], String(body.date || fmtDate_(now)),
          String(body.reason || '').slice(0, 200), String(body.giver || '').slice(0, 30),
          String(body.sticker || 'c:#f2b33d'), Number(p.values[3]) || 0, now
        ]);
        break;

      case 'deleteSticker':
        row = findRow_(ss, 'stickers', body.stickerId);
        if (row) sheet_(ss, 'stickers').deleteRow(row.row);
        break;

      case 'addPerson':
        var name = String(body.name || '').trim().slice(0, 20);
        if (!name) return out_({ ok: false, error: '이름이 비었어요' });
        sheet_(ss, 'people').appendRow([newId_(), name, String(body.color || COLORS[0]), 0, '', now]);
        break;

      case 'renamePerson':
        p = findRow_(ss, 'people', body.personId);
        if (p) sheet_(ss, 'people').getRange(p.row, 2).setValue(String(body.name || '').trim().slice(0, 20));
        break;

      case 'deletePerson':
        p = findRow_(ss, 'people', body.personId);
        if (p) sheet_(ss, 'people').getRange(p.row, 5).setValue('Y'); // 기록은 남기고 숨김
        break;

      case 'newBoard':
        p = findRow_(ss, 'people', body.personId);
        if (p) sheet_(ss, 'people').getRange(p.row, 4).setValue((Number(p.values[3]) || 0) + 1);
        break;

      case 'setGoal':
        var g = Math.max(5, Math.min(50, Number(body.goal) || 20));
        var s = findRow_(ss, 'settings', 'goal');
        if (s) sheet_(ss, 'settings').getRange(s.row, 2).setValue(g);
        else sheet_(ss, 'settings').appendRow(['goal', g]);
        break;

      case 'addImage':
        var src = String(body.src || '');
        if (src.indexOf('data:image/') !== 0 || src.length > 48000) return out_({ ok: false, error: '이미지가 너무 커요' });
        sheet_(ss, 'images').appendRow([newId_(), String(body.name || '').slice(0, 40), src, now]);
        break;

      case 'deleteImage':
        row = findRow_(ss, 'images', body.imageId);
        if (row) sheet_(ss, 'images').deleteRow(row.row);
        break;

      default:
        return out_({ ok: false, error: 'unknown_action' });
    }
    SpreadsheetApp.flush();
    return out_({ ok: true, data: readAll_(ss) });
  } finally {
    lock.releaseLock();
  }
}
