/******************************************************
 * Daythree × Taobao Malaysia New Hire Portal
 * Enterprise V2 - TWO FILE VERSION
 * Files: Code.gs + Index.html
 ******************************************************/

const APP = {
  ROOT_FOLDER_ID: '13adNWrDF9FMQ0NKlnKhgWzZS-F1sNery8z4JEj2s9lnK1bKOzY0ga15DedDzuGnLi4OyrqAg',
  DEFAULT_ADMIN_EMAIL: 'jasonpoen@gmail.com',
  DEFAULT_ADMIN_PASSWORD: '000206Abc#',
  SHEETS: {
    ADMINS: 'Admins',
    MASTER: 'Master',
    ATTACHMENTS: 'Attachments',
    NICKNAME: 'NicknameList',
    AUDIT: 'Audit Log',
    SETTINGS: 'Settings'
  }
};

const MASTER_HEADERS = [
  '提交时间','编号','状态','英文全名','中文全名','身份证号码','手机号码','出生日期','国籍','首次工作日期','美国绿卡','最高学历','工作花名【中文】','工作花名【英文】',
  '想进的组别','备注','资料夹链接','CreatedBy','UpdatedAt','Agent Email'
];
const ADMIN_HEADERS = ['Email','Name','Password','Salt','Role','Status','CreatedAt','LastPasswordChanged','Approval Status'];
const ATT_HEADERS = ['RecordID','FileType','FileName','FileUrl','FileId','UploadedAt'];
const NICK_HEADERS = ['中文花名','英文花名','Status','CreatedBy','CreatedAt'];
const AUDIT_HEADERS = ['Time','User','Action','Target','Details'];
const SETTING_HEADERS = ['Key','Value'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Daythree New Hire Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('请从 Google Sheet 的 Extensions > Apps Script 打开此项目。');
  return ss;
}
function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
function tz_() { return Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur'; }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function dateOnly_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return ymd_(value);
  return value == null ? '' : String(value);
}
function uuid_() { return Utilities.getUuid().replace(/-/g, ''); }
function salt_() { return uuid_() + uuid_(); }
function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + String(salt), Utilities.Charset.UTF_8);
  return bytes.map(function(b){ return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
}
function ensureHeader_(sheetName, headers) {
  const sh = sheet_(sheetName);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  const current = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
  headers.forEach(function(h, i){ current[i] = h; });
  sh.getRange(1, 1, 1, current.length).setValues([current]);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#fff0e0').setFontColor('#111827');
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, headers.length); } catch(e) {}
  return sh;
}
function setupSystem() {
  ensureHeader_(APP.SHEETS.MASTER, MASTER_HEADERS);
  ensureHeader_(APP.SHEETS.ADMINS, ADMIN_HEADERS);
  ensureHeader_(APP.SHEETS.ATTACHMENTS, ATT_HEADERS);
  ensureHeader_(APP.SHEETS.NICKNAME, NICK_HEADERS);
  ensureHeader_(APP.SHEETS.AUDIT, AUDIT_HEADERS);
  const settings = ensureHeader_(APP.SHEETS.SETTINGS, SETTING_HEADERS);
  if (settings.getLastRow() < 2) settings.appendRow(['RootFolderId', APP.ROOT_FOLDER_ID]);
  seedNicknames_();
  resetJasonAdmin();
  return {ok:true, message:'Setup completed'};
}
function setupLite_() {
  ensureHeader_(APP.SHEETS.MASTER, MASTER_HEADERS);
  ensureHeader_(APP.SHEETS.ADMINS, ADMIN_HEADERS);
  ensureHeader_(APP.SHEETS.ATTACHMENTS, ATT_HEADERS);
  ensureHeader_(APP.SHEETS.NICKNAME, NICK_HEADERS);
  ensureHeader_(APP.SHEETS.AUDIT, AUDIT_HEADERS);
  ensureHeader_(APP.SHEETS.SETTINGS, SETTING_HEADERS);
  seedNicknames_();
}
function seedNicknames_() {
  const sh = sheet_(APP.SHEETS.NICKNAME);
  if (sh.getLastRow() <= 1) {
    sh.appendRow(['小云','Yun','Active','System',now_()]);
    sh.appendRow(['小杰','Jay','Active','System',now_()]);
    sh.appendRow(['小慧','Hui','Active','System',now_()]);
  }
}
function audit_(user, action, target, details) {
  try { ensureHeader_(APP.SHEETS.AUDIT, AUDIT_HEADERS).appendRow([now_(), user || '', action || '', target || '', details || '']); } catch(e) {}
}
function resetJasonAdmin() {
  const sh = ensureHeader_(APP.SHEETS.ADMINS, ADMIN_HEADERS);
  const email = APP.DEFAULT_ADMIN_EMAIL.toLowerCase().trim();
  const s = salt_();
  const passHash = hash_(APP.DEFAULT_ADMIN_PASSWORD, s);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase().trim() === email) {
      sh.getRange(i + 1, 1, 1, 9).setValues([[email, 'Jason Poen', passHash, s, 'Super Admin', 'Active', rows[i][6] || now_(), now_(), 'Approved']]);
      audit_('System', 'RESET_SUPER_ADMIN', email, 'Password reset and approved');
      return {ok:true};
    }
  }
  sh.appendRow([email, 'Jason Poen', passHash, s, 'Super Admin', 'Active', now_(), now_(), 'Approved']);
  audit_('System', 'CREATE_SUPER_ADMIN', email, 'Created');
  return {ok:true};
}
function adminLogin(email, password) {
  setupLite_();
  email = String(email || '').toLowerCase().trim();
  password = String(password || '');
  if (!email || !password) return {ok:false, message:'请输入电邮和密码。'};
  const sh = sheet_(APP.SHEETS.ADMINS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).toLowerCase().trim() !== email) continue;
    if (String(r[5] || '') !== 'Active') return {ok:false, message:'此账号已停用。'};
    if (String(r[8] || '') !== 'Approved') return {ok:false, message:'您的管理员帐号尚未获得批准，请联系 Super Admin。'};
    const stored = String(r[2] || '');
    const salt = String(r[3] || '');
    let ok = false;
    if (salt && stored === hash_(password, salt)) ok = true;
    if (!ok && stored === password) {
      ok = true;
      const ns = salt_();
      sh.getRange(i + 1, 3, 1, 2).setValues([[hash_(password, ns), ns]]);
      sh.getRange(i + 1, 8).setValue(now_());
    }
    if (!ok) return {ok:false, message:'电邮或密码不正确。'};
    const token = uuid_() + uuid_();
    const expiry = new Date(Date.now() + 12 * 60 * 60 * 1000);
    PropertiesService.getScriptProperties().setProperty('TOKEN_' + token, JSON.stringify({email:email, expiry:expiry.getTime()}));
    audit_(email, 'LOGIN', 'Admin', 'Success');
    return {ok:true, token:token, user:{email:email, name:r[1] || email, role:r[4] || 'Admin', isSuper:String(r[4]) === 'Super Admin'}};
  }
  return {ok:false, message:'电邮或密码不正确。'};
}
function getUserByToken_(token) {
  setupLite_();
  if (!token) throw new Error('请重新登录。');
  const raw = PropertiesService.getScriptProperties().getProperty('TOKEN_' + token);
  if (!raw) throw new Error('请重新登录。');
  const sess = JSON.parse(raw);
  if (Number(sess.expiry) < Date.now()) throw new Error('登录已过期，请重新登录。');
  const sh = sheet_(APP.SHEETS.ADMINS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase().trim() === String(sess.email).toLowerCase().trim()) {
      if (String(rows[i][5]) !== 'Active' || String(rows[i][8]) !== 'Approved') throw new Error('账号权限已变更，请重新登录。');
      return {email:String(rows[i][0]), name:String(rows[i][1] || rows[i][0]), role:String(rows[i][4] || 'Admin'), isSuper:String(rows[i][4]) === 'Super Admin', row:i + 1};
    }
  }
  throw new Error('请重新登录。');
}
function requireSuper_(token) {
  const u = getUserByToken_(token);
  if (!u.isSuper) throw new Error('只有 Super Admin 可以执行此操作。');
  return u;
}
function logout(token) {
  try { PropertiesService.getScriptProperties().deleteProperty('TOKEN_' + token); } catch(e) {}
  return {ok:true};
}
function registerAdminAccount(payload) {
  setupLite_();
  payload = payload || {};
  const name = String(payload.name || '').trim();
  const email = String(payload.email || '').toLowerCase().trim();
  const password = String(payload.password || '');
  if (!name || !email || !password) return {ok:false, message:'请填写完整资料。'};
  if (password.length < 8) return {ok:false, message:'密码至少需要 8 个字符。'};
  const sh = sheet_(APP.SHEETS.ADMINS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase().trim() === email) return {ok:false, message:'这个电邮已经申请或存在。'};
  }
  const s = salt_();
  sh.appendRow([email, name, hash_(password, s), s, 'Admin', 'Active', now_(), now_(), 'Pending']);
  audit_(email, 'APPLY_ADMIN', email, 'Pending approval');
  return {ok:true, message:'申请已提交，请等待 Super Admin 批准。'};
}
function getInitialData() {
  setupLite_();
  return {ok:true, nicknames:getActiveNicknames_()};
}
function getActiveNicknames_() {
  const values = sheet_(APP.SHEETS.NICKNAME).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) if (String(values[i][2]) === 'Active' && values[i][0]) out.push({cn:String(values[i][0]), en:String(values[i][1] || '')});
  return out;
}
function safeName_(text) { return String(text || '').replace(/[\\/:*?"<>|]/g, ' ').trim() || 'New Agent'; }
function rootFolder_() { return DriveApp.getFolderById(APP.ROOT_FOLDER_ID); }
function getOrCreateFolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function submitAgentForm(payload) {
  setupLite_();
  payload = payload || {};
  const data = payload.data || {};
  const files = payload.files || [];
  ['englishName','chineseName','ic','phone','agentEmail','birthday','firstWorkingDate','education','positionCN'].forEach(function(k){ if (!data[k]) throw new Error('请填写完整资料。'); });
  const id = 'NH-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0,6).toUpperCase();
  const pending = getOrCreateFolder_(rootFolder_(), 'Pending');
  const folder = pending.createFolder(id + ' - ' + safeName_(data.englishName || data.chineseName));
  const folderUrl = folder.getUrl();
  const att = sheet_(APP.SHEETS.ATTACHMENTS);
  files.forEach(function(f){
    if (!f || !f.base64) return;
    const parts = String(f.base64).split(',');
    const content = parts.length > 1 ? parts[1] : parts[0];
    const blob = Utilities.newBlob(Utilities.base64Decode(content), f.mimeType || 'application/octet-stream', f.name || 'file');
    const file = folder.createFile(blob);
    att.appendRow([id, f.type || '文件', f.name || file.getName(), file.getUrl(), file.getId(), now_()]);
  });
  sheet_(APP.SHEETS.MASTER).appendRow([now_(), id, 'Pending', data.englishName || '', data.chineseName || '', data.ic || '', data.phone || '', data.birthday || '', data.nationality || '', data.firstWorkingDate || '', data.usVisa || '', data.education || '', data.positionCN || '', data.positionEN || '', data.group || '', data.remark || '', folderUrl, 'Agent', now_(), data.agentEmail || '']);
  audit_('Agent', 'SUBMIT', id, 'New hire submitted');
  return {ok:true, id:id};
}
function getDashboardData(token) {
  const user = getUserByToken_(token);
  const records = getRecords_();
  const nicknames = getNicknames_();
  const counts = {all:records.length, pending:0, approved:0, rejected:0, revision:0};
  records.forEach(function(r){ if (r.status === 'Pending') counts.pending++; if (r.status === 'Approved') counts.approved++; if (r.status === 'Rejected') counts.rejected++; if (r.status === 'Need Revision') counts.revision++; });
  return {ok:true, user:user, records:records, nicknames:nicknames, admins:user.isSuper ? getAdmins_() : [], counts:counts};
}
function getRecords_() {
  const values = sheet_(APP.SHEETS.MASTER).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i]; if (!r[1]) continue;
    out.push({ row:i + 1, submittedAt:dateOnly_(r[0]), id:String(r[1] || ''), status:String(r[2] || ''), englishName:String(r[3] || ''), chineseName:String(r[4] || ''), ic:String(r[5] || ''), phone:String(r[6] || ''), birthday:dateOnly_(r[7]), nationality:String(r[8] || ''), firstWorkingDate:dateOnly_(r[9]), usVisa:String(r[10] || ''), education:String(r[11] || ''), positionCN:String(r[12] || ''), positionEN:String(r[13] || ''), group:String(r[14] || ''), remark:String(r[15] || ''), folderUrl:String(r[16] || ''), createdBy:String(r[17] || ''), updatedAt:dateOnly_(r[18]), agentEmail:String(r[19] || '') });
  }
  return out.reverse();
}
function updateRecordStatus(token, id, status) {
  const user = getUserByToken_(token);
  const sh = sheet_(APP.SHEETS.MASTER);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][1]) === String(id)) { sh.getRange(i + 1, 3).setValue(status); sh.getRange(i + 1, 19).setValue(now_()); audit_(user.email, 'UPDATE_STATUS', id, status); return {ok:true}; }
  throw new Error('找不到资料。');
}

function updateRecordData(token, id, data) {
  const user = getUserByToken_(token);
  if (String(user.role) === 'Viewer') throw new Error('Viewer 只能查看，不能修改资料。');
  data = data || {};
  const sh = sheet_(APP.SHEETS.MASTER);
  const rows = sh.getDataRange().getValues();
  const allowedStatus = ['Pending','Approved','Rejected','Need Revision'];
  const status = allowedStatus.indexOf(String(data.status || 'Pending')) >= 0 ? String(data.status || 'Pending') : 'Pending';
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(id)) {
      const rowNo = i + 1;
      const oldData = {
        status: rows[i][2], englishName: rows[i][3], chineseName: rows[i][4], ic: rows[i][5], phone: rows[i][6],
        birthday: rows[i][7], nationality: rows[i][8], firstWorkingDate: rows[i][9], usVisa: rows[i][10], education: rows[i][11],
        positionCN: rows[i][12], positionEN: rows[i][13], group: rows[i][14], remark: rows[i][15], agentEmail: rows[i][19]
      };
      const values = [
        status,
        data.englishName || '',
        data.chineseName || '',
        data.ic || '',
        data.phone || '',
        data.birthday || '',
        data.nationality || '',
        data.firstWorkingDate || '',
        data.usVisa || '',
        data.education || '',
        data.positionCN || '',
        data.positionEN || '',
        data.group || '',
        data.remark || ''
      ];
      sh.getRange(rowNo, 3, 1, 14).setValues([values]);
      sh.getRange(rowNo, 19).setValue(now_());
      sh.getRange(rowNo, 20).setValue(data.agentEmail || '');
      audit_(user.email, 'EDIT_RECORD', id, JSON.stringify({before: oldData, after: data}));
      return {ok:true, message:'资料已更新'};
    }
  }
  throw new Error('找不到资料。');
}

function deleteRecord(token, id) {
  const user = getUserByToken_(token);
  const sh = sheet_(APP.SHEETS.MASTER);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][1]) === String(id)) { sh.deleteRow(i + 1); audit_(user.email, 'DELETE_RECORD', id, 'Deleted from Master'); return {ok:true}; }
  throw new Error('找不到资料。');
}
function getNicknames_() {
  const values = sheet_(APP.SHEETS.NICKNAME).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) if (values[i][0]) out.push({row:i + 1, cn:String(values[i][0] || ''), en:String(values[i][1] || ''), status:String(values[i][2] || ''), createdBy:String(values[i][3] || ''), createdAt:dateOnly_(values[i][4])});
  return out;
}
function addNickname(token, cn, en) { const user = getUserByToken_(token); cn = String(cn || '').trim(); en = String(en || '').trim(); if (!cn) throw new Error('请输入中文花名。'); sheet_(APP.SHEETS.NICKNAME).appendRow([cn, en, 'Active', user.email, now_()]); audit_(user.email, 'ADD_NICKNAME', cn, en); return {ok:true}; }
function setNicknameStatus(token, row, status) { const user = getUserByToken_(token); row = Number(row); if (row < 2) throw new Error('无效资料。'); sheet_(APP.SHEETS.NICKNAME).getRange(row, 3).setValue(status); audit_(user.email, 'NICKNAME_STATUS', row, status); return {ok:true}; }
function deleteNickname(token, row) { const user = getUserByToken_(token); row = Number(row); if (row < 2) throw new Error('无效资料。'); sheet_(APP.SHEETS.NICKNAME).deleteRow(row); audit_(user.email, 'DELETE_NICKNAME', row, 'Deleted'); return {ok:true}; }
function getAdmins_() {
  const rows = sheet_(APP.SHEETS.ADMINS).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) if (rows[i][0]) out.push({row:i + 1, email:String(rows[i][0] || ''), name:String(rows[i][1] || ''), role:String(rows[i][4] || ''), status:String(rows[i][5] || ''), createdAt:dateOnly_(rows[i][6]), lastPasswordChanged:dateOnly_(rows[i][7]), approval:String(rows[i][8] || '')});
  return out;
}
function updateAdminAccount(token, row, role, status, approval) {
  const user = requireSuper_(token);
  row = Number(row); if (row < 2) throw new Error('无效资料。');
  const sh = sheet_(APP.SHEETS.ADMINS);
  sh.getRange(row, 5, 1, 2).setValues([[role || 'Admin', status || 'Active']]);
  sh.getRange(row, 9).setValue(approval || 'Pending');
  audit_(user.email, 'UPDATE_ADMIN', row, [role,status,approval].join(' / '));
  return {ok:true};
}
function resetAdminPassword(token, row, newPassword) {
  const user = requireSuper_(token);
  row = Number(row); if (row < 2) throw new Error('无效资料。');
  newPassword = String(newPassword || '').trim();
  if (newPassword.length < 8) throw new Error('新密码至少需要 8 个字符。');
  const s = salt_();
  sheet_(APP.SHEETS.ADMINS).getRange(row, 3, 1, 2).setValues([[hash_(newPassword, s), s]]);
  sheet_(APP.SHEETS.ADMINS).getRange(row, 8).setValue(now_());
  audit_(user.email, 'RESET_ADMIN_PASSWORD', row, 'Password reset');
  return {ok:true};
}

function changeOwnPassword(token, currentPassword, newPassword) {
  const user = getUserByToken_(token);
  currentPassword = String(currentPassword || '');
  newPassword = String(newPassword || '').trim();
  if (!currentPassword || !newPassword) throw new Error('请填写当前密码和新密码。');
  if (newPassword.length < 8) throw new Error('新密码至少需要 8 个字符。');
  const sh = sheet_(APP.SHEETS.ADMINS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const email = String(rows[i][0] || '').toLowerCase().trim();
    if (email !== String(user.email).toLowerCase().trim()) continue;
    const stored = String(rows[i][2] || '');
    const salt = String(rows[i][3] || '');
    let ok = false;
    if (salt && stored === hash_(currentPassword, salt)) ok = true;
    if (!ok && stored === currentPassword) ok = true;
    if (!ok) throw new Error('当前密码不正确。');
    const ns = salt_();
    sh.getRange(i + 1, 3, 1, 2).setValues([[hash_(newPassword, ns), ns]]);
    sh.getRange(i + 1, 8).setValue(now_());
    audit_(user.email, 'CHANGE_OWN_PASSWORD', user.email, 'Password changed by user');
    return {ok:true, message:'密码已成功更换，请使用新密码登入。'};
  }
  throw new Error('找不到管理员账号。');
}

function exportSelectedCsv(token, fields, statusFilter, dateFrom, dateTo) {
  getUserByToken_(token);
  const map = {submittedAt:'提交时间', id:'编号', status:'状态', englishName:'英文全名', chineseName:'中文全名', ic:'身份证号码', phone:'手机号码', birthday:'出生日期', nationality:'国籍', firstWorkingDate:'首次工作日期', usVisa:'美国绿卡', education:'最高学历', positionCN:'工作花名【中文】', positionEN:'工作花名【英文】', group:'想进的组别', remark:'备注', folderUrl:'资料夹链接', agentEmail:'Agent Email'};
  fields = fields && fields.length ? fields : Object.keys(map);
  let records = getRecords_();
  if (statusFilter && statusFilter !== '全部') records = records.filter(function(r){ return r.status === statusFilter; });
  if (dateFrom) records = records.filter(function(r){ return String(r.submittedAt || '').slice(0,10) >= String(dateFrom); });
  if (dateTo) records = records.filter(function(r){ return String(r.submittedAt || '').slice(0,10) <= String(dateTo); });
  const csv = [fields.map(function(f){ return csv_(map[f] || f); }).join(',')].concat(records.map(function(r){ return fields.map(function(f){ return csv_(r[f]); }).join(','); })).join('\n');
  return {ok:true, filename:'NewHire_Export_' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd_HHmmss') + '.csv', csv:csv};
}
function csv_(v) { v = String(v == null ? '' : v); return '"' + v.replace(/"/g, '""') + '"'; }
