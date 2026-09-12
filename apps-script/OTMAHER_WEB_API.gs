/**
 * OTMAHER WEB PANELİ – SALT OKUNUR API
 *
 * Bu dosya web paneline yalnızca OTMAHER rapor verilerini JSON olarak verir.
 * Adisyo/Trendyol API çağrısı yapmaz, sipariş durumu değiştirmez ve Sheet'e yazmaz.
 *
 * Kurulum:
 * 1) Apps Script > Proje Ayarları > Script Properties içine OTMAHER_WEB_TOKEN ekleyin.
 * 2) Bu dosyayı mevcut OTMAHER Apps Script projesine ekleyin.
 * 3) Web app olarak dağıtın. doGet yalnızca GET kabul eder.
 */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (String(p.view || '') !== 'dashboard') return otmaherJson_({ok:false,error:'Geçersiz görünüm'});
    var expected = String(PropertiesService.getScriptProperties().getProperty('OTMAHER_WEB_TOKEN') || '');
    var supplied = String(p.token || '');
    if (!expected || !supplied || supplied !== expected) return otmaherJson_({ok:false,error:'Yetkisiz erişim'});
    return otmaherJson_(otmaherDashboardPayload_());
  } catch (err) {
    return otmaherJson_({ok:false,error:String(err && err.message || err)});
  }
}

function otmaherDashboardPayload_() {
  var ss = SpreadsheetApp.openById('1bcPJFxpg1nPApDAl_Gngp70rGQTOoDnJRtsLLbcYo3s');
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Istanbul';
  var karGecmis = otmaherValues_(ss, 'ADISYO_KAR_GECMIS', 17);
  var karPanel = otmaherValues_(ss, 'ADISYO_KAR_PANELI', 10);
  var sistem = otmaherValues_(ss, 'ADISYO_SISTEM_DURUMU', 4);
  var karHesap = otmaherValues_(ss, 'ADISYO_KAR_HESAP', 28);
  var ham = otmaherValues_(ss, 'ADISYO_HAM', 35);

  var history = [];
  for (var i=1;i<karGecmis.length;i++) {
    var r=karGecmis[i]; if (!r[0]) continue;
    history.push({
      date:otmaherDateIso_(r[0],tz), orders:otmaherNum_(r[2]), mainProducts:otmaherNum_(r[3]),
      revenue:otmaherNum_(r[4]), productCost:otmaherNum_(r[7]), platformCut:otmaherNum_(r[8]), vat:otmaherNum_(r[9]),
      preFixedProfit:otmaherNum_(r[10]), fixedCost:otmaherNum_(r[11]), temporaryNet:otmaherNum_(r[12]),
      definiteNet:r[13]===''||r[13]==null?null:otmaherNum_(r[13]), accumulatedDefinite:otmaherNum_(r[14]), issue:String(r[15]||''), status:String(r[16]||'')
    });
  }
  history.sort(function(a,b){return String(a.date).localeCompare(String(b.date));});
  var overview = history.length ? history[history.length-1] : {};

  var panelMap={};
  karPanel.forEach(function(r){if(r[0])panelMap[String(r[0]).trim()]=r[1];});
  var accumulated={
    startDate:otmaherDateIso_(panelMap['Başlangıç tarihi'],tz),
    calculatedDays:otmaherNum_(panelMap['Hesaplanan gün']),
    definiteNet:otmaherNum_(panelMap['Birikimli kesin net kâr']),
    lastOperationDay:otmaherDateIso_(panelMap['Son işlem günü'],tz),
    lastDayTemporaryNet:otmaherNum_(panelMap['Son gün geçici net kâr']),
    lastDayDefiniteNet:panelMap['Son gün kesin net kâr']===''?null:otmaherNum_(panelMap['Son gün kesin net kâr']),
    lastDayStatus:String(panelMap['Son gün durumu']||'')
  };

  var sysMap={}; sistem.forEach(function(r){if(r[0])sysMap[String(r[0]).trim()]=r[1];});
  var system={
    status:String(sysMap['Durum']||''), lastAttempt:otmaherDateTimeIso_(sysMap['Son deneme'],tz), lastSuccess:otmaherDateTimeIso_(sysMap['Son başarılı senkron'],tz),
    consecutiveErrors:otmaherNum_(sysMap['Ardışık hata']), closedOrders:otmaherNum_(sysMap['Kapanmış sipariş']), idMatched:otmaherNum_(sysMap['Kimlikle eşleşen satır']),
    nameMatched:otmaherNum_(sysMap['Adla yedek eşleşen satır']), unmatched:String(sysMap['Eşleşmeyen']||'Yok')
  };

  var latestDate=overview.date||'';
  var breakdownMap={}; var issuesMap={};
  for (var j=1;j<karHesap.length;j++) {
    var k=karHesap[j]; if(!k[0])continue;
    var date=otmaherDateIso_(k[0],tz); var issue=String(k[27]||'').trim();
    if(issue){var ik=[k[4],k[5],k[6],issue].join('|');issuesMap[ik]={kitchen:String(k[4]||''),platform:String(k[5]||''),product:String(k[6]||''),issue:issue};}
    if(date!==latestDate)continue;
    var key=[k[4],k[5]].join('|');
    if(!breakdownMap[key])breakdownMap[key]={kitchen:String(k[4]||'Bilinmiyor'),platform:String(k[5]||'Bilinmiyor'),orders:0,revenue:0,cost:0,cut:0,vat:0,profit:0,issues:0};
    var b=breakdownMap[key]; b.orders+=otmaherNum_(k[10]); b.revenue+=otmaherNum_(k[11]); b.cost+=otmaherNum_(k[17])+otmaherNum_(k[18]); b.cut+=otmaherNum_(k[22]); b.vat+=otmaherNum_(k[23]);
    if(otmaherNum_(k[10])===1)b.profit+=otmaherNum_(k[11])-otmaherNum_(k[17])-otmaherNum_(k[18])-otmaherNum_(k[22])-otmaherNum_(k[23]); if(issue)b.issues++;
  }
  var breakdown=Object.keys(breakdownMap).map(function(x){return breakdownMap[x];});
  var issues=Object.keys(issuesMap).map(function(x){return issuesMap[x];});

  var ordersById={};
  for(var h=1;h<ham.length;h++){
    var x=ham[h], id=String(x[1]||'').trim(); if(!id)continue;
    var dt=otmaherDateTimeIso_(x[3],tz); if(!dt)continue;
    if(!ordersById[id])ordersById[id]={orderId:id,orderNo:String(x[2]||''),dateTime:dt,time:Utilities.formatDate(new Date(dt),tz,'HH:mm'),status:String(x[6]||''),payment:otmaherNum_(x[14])||otmaherNum_(x[9]),discount:otmaherNum_(x[10]),platform:String(x[18]||x[16]||''),kitchen:otmaherKitchen_(x[20]),products:[]};
    var o=ordersById[id]; var product=String(x[25]||'').trim(); if(product&&o.products.indexOf(product)<0)o.products.push(product);
  }
  var recentOrders=Object.keys(ordersById).map(function(id){return ordersById[id];}).sort(function(a,b){return String(b.dateTime).localeCompare(String(a.dateTime));}).slice(0,80);

  return {ok:true,generatedAt:new Date().toISOString(),overview:overview,accumulated:accumulated,history:history,breakdownDate:latestDate,breakdown:breakdown,issues:issues,system:system,recentOrders:recentOrders};
}

function otmaherValues_(ss,name,width){var sh=ss.getSheetByName(name);if(!sh)return [];var rows=sh.getLastRow();if(rows<1)return [];return sh.getRange(1,1,rows,width).getValues();}
function otmaherNum_(v){if(v===''||v==null)return 0;if(typeof v==='number')return isFinite(v)?v:0;var s=String(v).replace(/₺/g,'').replace(/\s/g,'').replace(/\./g,'').replace(',','.').replace('%','');var n=Number(s);return isFinite(n)?n:0;}
function otmaherDateIso_(v,tz){if(!v)return '';if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,'yyyy-MM-dd');var s=String(v).trim(),m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);if(m)return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);m=s.match(/^(\d{4}-\d{2}-\d{2})/);return m?m[1]:'';}
function otmaherDateTimeIso_(v,tz){if(!v)return '';if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,"yyyy-MM-dd'T'HH:mm:ssXXX");var d=new Date(v);return isNaN(d.getTime())?String(v):d.toISOString();}
function otmaherKitchen_(v){var s=String(v||'').toLocaleLowerCase('tr-TR');if(s.indexOf('acele et')>=0)return 'Acele Et Izgara';if(s.indexOf('cızzgara')>=0||s.indexOf('cizzgara')>=0)return 'Cızzgara';if(s.indexOf('pilav durağı')>=0||s.indexOf('pilav duragi')>=0||s.indexOf('ekmek arası')>=0||s.indexOf('ekmek arasi')>=0)return 'Pilav Durağı';return String(v||'Bilinmiyor');}
function otmaherJson_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
