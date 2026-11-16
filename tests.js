// Automated checks for the Account app (index.html). The app never loads this file.
// Everything in here is made up: no real names, balances or amounts.
//
// Run:   node tests.js            (in the folder that holds index.html, with Node.js 18 or newer)
//        node tests.js other.html (to check a different file)
// Pass:  the last line reads "ALL PASSED" and the exit code is 0.
//
// How it works: each suite gets a fresh copy of the app's script (taken from the HTML file) running against
// a fake page, fake storage, a made-up setup (FX below, in the shape of the three App… tables in Excel) and a
// fake Excel (fetch is replaced inside each suite). Nothing touches the network or a workbook.
// Inside a suite, ok(condition, 'what it proves') records one check.
//
// When you change the app: run this first, make the change, add or adjust checks for it, run again.
// Layout and tap behaviour (phone width, tabs, Pay favorites in a real browser) are not covered here.
'use strict';
const fs = require('fs'), vm = require('vm'), path = require('path');

const file = process.argv[2] || ['index.html', 'cashflow-tracker.html'].find(f => fs.existsSync(path.join(__dirname, f)));
if (!file) { console.error('No index.html next to tests.js. Pass the file: node tests.js path/to/app.html'); process.exit(2); }
const html = fs.readFileSync(path.isAbsolute(file) ? file : path.join(__dirname, file), 'utf8');
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
if (!scripts.length) { console.error(`No inline <script> found in ${file}`); process.exit(2); }
const APP = scripts[scripts.length - 1][1];

// The fake page every suite starts from: storage, elements that remember what is set on them,
// an online browser, and the made-up setup
const PAGE = String.raw`
var global = globalThis;
const store={}; var localStorage={getItem:k=>store[k]??null,setItem:(k,v)=>store[k]=v,removeItem:k=>delete store[k]};
const els={}; const mk=()=>({classList:{add(){},remove(){},toggle(){},contains(){return false}},style:{},textContent:'',innerHTML:'',disabled:false,value:'',focus(){},appendChild(){},dataset:{},querySelector(){return mk()}});
var window={addEventListener(){},location:{pathname:'/'}}; var document={getElementById:id=>els[id]||(els[id]=mk()),addEventListener(){},querySelectorAll:()=>[],querySelector:()=>null,createElement:()=>({})};
var navigator={onLine:true};
let P=0,F=0; const ok=(c,m)=>{c?P++:(F++,console.log('✗ '+m));};
// Made-up setup, in the shape the app reads from Excel (header row first). No real names or amounts.
const FX = {
  AppSettings: [['Key', 'From', 'Value'],
    ['gas.label', '', 'Gas bill'], ['gas.venue', '', 'GasCo'], ['gas.activities', '', 'Gas'],
    ['internet.label', '', 'Internet'], ['internet.venue', '', 'NetCo'], ['internet.activities', '', 'Internet plan'], ['internet.amount', '', 30],
    ['electric.label', '', 'Electric'], ['electric.venue', '', 'PowerCo'], ['electric.activities', '', 'Electric bill'],
    ['hsaEmployer.label', '', 'Employer HSA'], ['hsaEmployer.venue', '', 'Acme'], ['hsaEmployer.activities', '', 'Employer contribution'], ['hsaEmployer.amount', '', 100],
    ['hsaMine.label', '', 'My HSA'], ['hsaMine.venue', '', 'Acme'], ['hsaMine.activities', '', 'My contribution'], ['hsaMine.amount', '', 200], ['hsaMine.amount', '2027-01', 250],
    ['interest.label', '', 'HSA interest'], ['interest.venue', '', 'Acme'], ['interest.activities', '', 'Interest'],
    ['pay.label', '', 'Pay'], ['pay.venue', '', 'Acme'], ['pay.activities', '', 'Pay'], ['pay.amount', '', 4000],
    ['rent.label', '', 'Rent bill'], ['rent.venue', '', 'Oak Homes'], ['rent.rent', '', 1000], ['rent.parking', '', 100],
    ['phone.label', '', 'Phone'], ['phone.venue', '', 'Pat'], ['phone.activities', '', 'Phone bill'], ['phone.amount', '', 25],
    ['savings.family', '', 500], ['savings.family', '2026-11', 600], ['savings.amount', '2026-11', 2000],
    ['fidelity.amount', '', 300], ['fidelity.amount', '2026-11', 400],
    ['retire.label', '', 'Retirement'], ['retire.venue', '', 'RetireCo'], ['retire.part1', '', 'Employer'], ['retire.part2', '', 'Personal'],
    ['retire.activities1', '', 'Employer Contribution'], ['retire.activities2', '', 'Personal Contribution'],
    ['retire.amount1', '', 800], ['retire.amount1', '2026-11', 1000], ['retire.amount2', '', 150], ['retire.amount2', '2026-11', 200],
    ['analysis.rows', '', 'F&D, Groceries FD, Out FD, Caffeine FD, Transport, Gas TR, Want, Dr, Travel, Need, Hair, Eyes, Home, Utilities, Pay, Savings']],
  FavTable: [['Label', 'Section', 'Table', 'Venue', 'Activities', 'Method', 'Category', 'Sub', 'Cost', 'Remarks'],
    ['Grocer', 'everyday', 'Credit', 'Corner Grocer', 'Groceries', 'Savor1', 'F&D', 'Groceries', '', ''],
    ['Park', 'everyday', 'Credit', 'City Parking', 'Parking by ', 'Savor1', 'Transport', 'Parking', '', ''],
    ['Cloud', 'subscription', 'Credit', 'CloudCo', 'Storage', 'WF', 'Want', 'Service', 2.5, '11th'],
    ['Music', 'subscription', 'Credit', 'Tunes', 'Monthly sub', 'Savor1', 'Want', 'Music', 9.99, '25th'],
    ['Discover CB', 'cash back', 'Credit', 'Discover', 'Cash back', 'Discover', '', '', '', ''],
    ['Chase CB', 'cash back', 'Credit', 'Chase', 'Cash back', 'Chase', '', '', '', ''],
    ['DeltaAmex CB', 'cash back', 'Credit', 'Amex', 'Statement credit', 'DeltaAmexGold', '', '', '', ''],
    ['Discover Pay', 'payment', 'Credit', 'Discover', 'Discover Payment', 'Discover', '', '', '', ''],
    ['Chase Pay', 'payment', 'Credit', 'Chase', 'Chase Payment', 'Chase', '', '', '', ''],
    ['DeltaAmex Pay', 'payment', 'Credit', 'Amex', 'Amex Payment', 'DeltaAmexGold', '', '', '', ''],
    ['Coffee', '', 'Cash', 'Cafe', 'Coffee', 'Venmo', 'F&D', 'Out', 4, '']],
  CatTable: [['F&D', 'Transport', 'Want', 'Need', 'Travel', 'Pay', 'Health', 'Home', 'Utilities'],
    ['Groceries', 'Parking', 'Music', 'Phone', '', '', '', 'Rent', 'Wifi'],
    ['Out', 'Gas', 'Service', '', '', '', '', '', 'Gas'],
    ['Caffeine', '', '', '', '', '', '', '', 'Electric'],
    ['', '', '', '', '', '', '', '', 'Others']]
};
// Called once the app's script has loaded: the phone already holds a setup read from Excel
function useFixture() {
  config = { ...config, settings: parseSettings(FX.AppSettings), favorites: parseFavorites(FX.FavTable), categories: parseCategories(FX.CatTable), loaded: 'earlier', missing: [] };
  catSub = orderedCategories(); favorites = config.favorites;
}
`;

const SUITES = {
  // ─────────────────────────────────────────────────────────────────────────────
  'Core: money, storage, transfers, rent split, sync engine, balances, tabs': async () => {
let msg=''; showToast=m=>msg=m; global.setTimeout=fn=>fn(); confirm=()=>true; renderQueueReview=()=>{};
const g=id=>document.getElementById(id);
// Money + dates
ok(parseMoneyInput('1234')===12.34 && parseMoneyInput('12.5')===12.5 && parseMoneyInput('1,234.50')===1234.5 && isNaN(parseMoneyInput('')),'money input: no dot = cents');
ok(fmtMoney(-1234.5)==='1,234.50' && round2(0.1+0.2)===0.3,'fmtMoney / round2');
ok(localDateStr(new Date(2026,0,5))==='2026-01-05','local date');
ok(cellToSerial(46307)===46307 && cellToSerial('2026-10-13')===excelSerial(2026,10,13) && cellToSerial('junk')===null,'date cells');
// Storage safety
store.cf_bad='{oops'; ok(readJSON('cf_bad',[]).length===0 && readJSON('cf_missing',7)===7,'damaged/missing saved data falls back');
// Method preview
currentSheet='Cash'; g('f-method').value='HSA Reimb'; onMethodChange(); ok(g('account-preview').innerHTML.includes('HSA Reimb'),'method preview shows account');
g('f-method').value=''; onMethodChange(); ok(g('account-preview').innerHTML.includes('select method first'),'no method → hint');
// Transfers
let t=transferRows('SavingsCO','Checking',100,'',''); ok(t[0].activities==='Transfer (-savings)' && t[1].category==='Savings' && !t[0].category,'out of savings');
t=transferRows('Checking','SavingsCO',100,'',''); ok(t[0].category==='Savings' && t[0].cost===100 && t[1].cost===-100,'into savings');
t=transferRows('Cash','Checking',20,'Mom',''); ok(t[0].venue==='Mom' && t[1].venue==='Wells Fargo','cash venue');
// Rent bill split / apartment totals (rent 1,000 + parking 100 in the made-up setup)
const w=rentRows(1170,'2026-10',{venue:'Oak Homes'}); ok(w[0].cost===1000 && w[1].cost===100 && w[2].cost===70 && w[1].remarks==='Total Oak Homes payment = 1,170.00' && w.every(r=>r.venue==='Oak Homes'),'rent bill split: rent, parking, sewer = the rest');
ok(apartmentTotals({rent:1170,gas:40,electric:60,internet:30},'2026-10')==='Gas + electric + sewer = 170.00 · Total apartment paid = 1,300.00','apartment totals');
// Sync engine
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; accessToken='tok';
const COLS={CashTable:['Date','Venue','Method','Cost','Checking','Remarks','Inputted','Uploaded'],CreditTable:['Date','Venue','Method','Cost','WF','Remarks','Inputted','Uploaded']};
let log=[],mode={},colCalls=0;
global.fetch=async(url,opt={})=>{ const u=decodeURIComponent(url); const tn=(u.match(/tables\/([^/]+)\//)||[])[1]; const r=(s,b)=>({ok:s<300,status:s,json:async()=>b});
  if(/\/columns(\?\$select=name)?$/.test(u)){ colCalls++; await new Promise(r=>setImmediate(r)); return r(200,{value:COLS[tn].map(name=>({name}))}); }
  if(u.endsWith('/createSession')) return r(201,{id:'S'}); if(u.endsWith('/closeSession')) return r(204,{});
  if(u.endsWith('/rows/add')){ const v=JSON.parse(opt.body).values; log.push({tn,v,sid:(opt.headers||{})['workbook-session-id']}); const m=mode[tn];
    if(m==='400') return r(400,{}); if(m==='504') return r(504,{}); if(m==='drop') throw new Error('x'); return r(201,{}); }
  throw new Error('unexpected '+u); };
const row=(sheet,c)=>({sheet,payload:{date:'2026-11-02',venue:'V'+c,method:'Checking',cost:c,remarks:'R',inputted:'now'}});
const reset=()=>{log=[];mode={};queue=[];saveQueue();columnCache={};colCalls=0;msg='';};
reset(); await Promise.all([getTableColumns('CashTable'),getTableColumns('CashTable'),getTableColumns('CashTable')]); ok(colCalls===1,'concurrent column lookups share one request');
reset(); queue=[row('Cash',1),row('Credit',2),row('Cash',3)]; saveQueue(); await syncQueue();
ok(log.length===2 && log.every(l=>l.sid==='S') && queue.length===0 && msg==='✓ All synced','2 tables → 2 batched writes in a session');
const cv=log.find(l=>l.tn==='CashTable').v; ok(cv.length===2 && cv[0][4]===null && cv[0][5]==='R' && cv[0][7],'formula cols null, data mapped, Uploaded stamped');
reset(); mode.CreditTable='504'; queue=[row('Cash',1),row('Credit',2)]; saveQueue(); await syncQueue();
ok(queue.length===1 && queue[0].uncertain && queue[0].sheet==='Credit','5xx → ⚠ uncertain kept, confirmed rows removed');
log=[]; await syncQueue(); ok(log.length===0,'uncertain never auto-resent');
reset(); mode.CashTable='400'; queue=[row('Cash',1)]; saveQueue(); await syncQueue(); ok(queue.length===1 && !queue[0].uncertain,'4xx → plain failure, retryable');
reset(); mode.CashTable='drop'; queue=[row('Cash',1)]; saveQueue(); await syncQueue(); ok(queue[0].uncertain,'dropped connection → uncertain');
reset(); queue=[{...row('Cash',1),sending:true}]; ok(true,'');
// Direct send
queueMode=false; reset(); let d=await commitRows([row('Cash',1),row('Credit',2)],'x'); ok(d && log.length===2 && queue.length===0,'direct send writes both tables');
reset(); mode.CashTable='400'; d=await commitRows([row('Cash',1)],'x'); ok(d===false && queue.length===0,'first table definite failure → keep form');
reset(); mode.CreditTable='504'; d=await commitRows([row('Cash',1),row('Credit',2)],'x'); ok(d && queue.length===1 && queue[0].uncertain,'second table unsure → queued ⚠');
queueMode=true; reset(); d=await commitRows([row('Cash',1)],'x'); ok(d && queue.length===1 && log.length===0,'queue mode → queued, nothing sent');
// Today balances incl. derived COne Savings
balanceCache={Cash:{Checking:1000,SavingsCO:2000,SavingsF:500,SavingsFL:10,'COne Savings':2510,'Cost (MYR)':9},Credit:{Discover:1,WF:300,Chase:3,Savor1:4,Venmo:7,Total:0}};
futureCache={Cash:{sums:{checking:-200,savingsco:-900},count:2,after:{},credits:{},cutoff:monthEndSerial()},Credit:{sums:{wf:14.16},count:1,after:{},credits:{},cutoff:monthEndSerial()}};
balMode='today'; ok(balValue('Cash','Checking',1000)===800 && balValue('Cash','COne Savings',2510)===1610 && balValue('Credit','WF',300)===285.84,'today values (Cash adds back, Credit subtracts, derived sums)');
setBalMode('today'); const mh=g('balance-main').innerHTML, eh=g('balance-expand').innerHTML;
ok(mh.includes('Total Credit') && mh.includes('$293.84') && !eh.includes('Cost (MYR)') && !eh.includes('(Credit)') && eh.includes('$1,610.00'),'render: Total Credit, MYR/Credit Venmo hidden, derived shown');
setBalMode('all'); ok(g('balance-main').innerHTML.includes('$308.00'),'All mode uses raw values');
setBalMode('today');
// Tabs keep a chosen favorite
switchSheet('Credit'); const pi=favorites.findIndex(x=>x.label==='Discover Pay'); activeFav=pi; g('f-venue').value='Discover';
switchSheet('Balances'); switchSheet('Credit'); switchSheet('Credit'); ok(activeFav===pi && g('f-venue').value==='Discover','back to same tab (incl. via Balances) keeps the favorite');
switchSheet('Cash'); ok(activeFav===null && g('f-venue').value==='' && currentSheet==='Cash','other table with a favorite → form cleared');
g('f-venue').value='Typed'; switchSheet('Credit'); ok(g('f-venue').value==='Typed','typed entry (no favorite) carries over');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Setup read from Excel (settings, favorites, categories)': async () => {
let toasts=[]; showToast=m=>toasts.push(m); global.setTimeout=fn=>fn(); const g=id=>document.getElementById(id);
store.cf_file_path='a.xlsx';
let redirects=0, silentFail=false;
const signedIn={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>{ if(silentFail) throw new Error('x'); return {accessToken:'tok'}; },acquireTokenRedirect:()=>{redirects++;},loginRedirect:()=>{redirects++;}};
msalInstance=signedIn;
// ── Reading the tables
const ser=excelSerial(2026,11,1);
let st=parseSettings([['Key','From','Value'],[' Pay.Amount ','',100],['pay.amount','2027-1',300],['PAY.amount',ser,200],['pay.venue','',' Acme '],['blank.value','',''],['','', 5]]);
ok(JSON.stringify(st['pay.amount'])===JSON.stringify([{from:'',value:100},{from:'2026-11',value:200},{from:'2027-01',value:300}]),'settings: keys trimmed + any case, From as text or as an Excel date, kept in month order');
ok(st['pay.venue'][0].value==='Acme' && !st['blank.value'] && Object.keys(st).length===2,'values trimmed; rows without a key or value ignored');
ok(JSON.stringify(parseSettings([['Value','Key','From'],[5,'x.y','2026-11']]))==='{"x.y":[{"from":"2026-11","value":5}]}','columns found by header, in any order');
config={...config,settings:st};
ok(cfgNum('pay.amount','2026-10')===100 && cfgNum('pay.amount','2026-11')===200 && cfgNum('pay.amount','2026-12')===200 && cfgNum('pay.amount','2027-03')===300 && cfgNum('pay.amount')===300,'a month gets the latest row not after it; no month → newest');
config={...config,settings:parseSettings([['Key','From','Value'],['a.b','2026-11',9]])}; ok(cfgNum('a.b','2026-10')===undefined && cfgText('a.b','2026-10')==='' && cfgText('nope')==='' && cfgNum('A.B','2026-11')===9,'before its first From, or unknown → nothing');
const fv=parseFavorites([FX.FavTable[0],['A','Subscription','credit','V','Act ','WF','C','S','12.5','r'],['B','Cash Back','Credit','V','x','WF','','','',''],['C','CB','Credit','V','x','WF','','','',''],
  ['D','Payments','Credit','Bank','Bank Payment','Chase','','','',''],['E','','Sheet9','V','x','WF','','','',''],['','everyday','Credit','V','x','WF','','','',''],[' F ',' everyday ','CASH',' V ','x','Venmo','','',-3,5]]);
ok(fv.map(f=>f.label).join()==='A,B,C,D,F','favorites: rows without a label or with an unknown table are skipped');
ok(fv[0].group==='sub' && fv[0].sheet==='Credit' && fv[0].cost===12.5 && fv[0].activities==='Act ' && fv[1].group==='cb' && fv[1].negativeDefault && fv[2].group==='cb' && fv[4].group===undefined && fv[4].sheet==='Cash' && fv[4].venue==='V' && fv[4].cost===-3 && fv[4].remarks==='5','sections in any wording/case, numbers as text, activities keeps its end space');
ok(fv[3].group==='pay' && fv[3].autoFillCol==='Chase' && JSON.stringify(fv[3].linked)===JSON.stringify({sheet:'Cash',venue:'Bank',activities:'Bank Payment',method:'Checking',category:'Credit',sub:''}),'payment row → amount lookup + Checking side');
ok(JSON.stringify(parseCategories([['A','','B',' C '],['x','junk','',' y '],['','','','z']]))==='{"A":["x"],"B":[],"C":["y","z"]}','categories: one column each, blanks dropped');
ok(JSON.stringify(parseCategories([['Category','Pay','Home'],['Sub-categories below','','Rent'],['','','Move']]))==='{"Pay":[],"Home":["Rent","Move"]}','the Cat sheet\'s own "Category" label column is not a category');
config={...config,categories:parseCategories([['Pay','Home','F&D','Debt','Dr'],['','Rent',' Out ','','']]),settings:parseSettings([['Key','From','Value'],['categories.order','',' f&d , Missing, Dr,Pay ']])};
ok(Object.keys(orderedCategories()).join()==='F&D,Dr,Pay,Home,Debt' && orderedCategories()['F&D'][0]==='Out','categories.order: named ones first (any case, unknown names ignored), the rest as on the sheet');
config={...config,settings:{}}; ok(Object.keys(orderedCategories()).join()==='Pay,Home,F&D,Debt,Dr','no order setting → sheet order');
useFixture();
// ── Loading from Excel
let calls=[], mode={}, T=JSON.parse(JSON.stringify(FX));
global.fetch=async(url)=>{ const u=decodeURIComponent(url); calls.push(u); const r=(s,b)=>({ok:s<300,status:s,json:async()=>b}); const tn=(u.match(/tables\/([^/]+)\/range\?\$select=values$/)||[])[1];
  if(!tn||!T[tn]) throw new Error('unexpected '+u); if(mode[tn]==='drop') throw new Error('net'); if(mode[tn]) return r(mode[tn],{}); return r(200,{values:T[tn]}); };
config={settings:{},favorites:[],categories:{},loaded:null,missing:[]}; catSub=config.categories; favorites=config.favorites; store.cf_favorites='old'; store.cf_cat_sub='old'; store.cf_fav_deleted='old';
viewMonth='2026-11'; expandedMonthly=null; renderMonthly(); ok(g('monthly-list').innerHTML.includes('AppSettings table in Excel'),'nothing loaded yet → Monthly explains where its setup comes from');
openConfig(); ok(g('cfg-excel-status').textContent.startsWith('Not loaded yet'),'setup shows "not loaded yet"');
switchSheet('Credit'); ok(g('fav-section').style.display==='none','no favorites yet → Quick add hidden');
ok(await loadConfig()===true && calls.length===3 && ['AppSettings','FavTable','CatTable'].every(t=>calls.some(u=>u.includes(`tables/${t}/range?$select=values`))),'one read per table, three in all');
ok(favorites.length===11 && Object.keys(catSub).length===9 && cfgNum('pay.amount','2026-11')===4000 && JSON.parse(store.cf_config).favorites.length===11,'setup applied and kept on the phone');
T.AppSettings=[...FX.AppSettings,['categories.order','','Want, F&D']]; await loadConfig(true); ok(Object.keys(catSub).slice(0,3).join()==='Want,F&D,Transport' && Object.keys(catSub).length===9,'adding categories.order in Excel reorders the form\'s list on reload'); T.AppSettings=FX.AppSettings; await loadConfig(true); configLoadedAt=Date.now();
ok(g('fav-chips').innerHTML.includes('>Grocer<') && g('fav-section').style.display==='' && !g('monthly-list').innerHTML.includes('AppSettings table'),'chips and Monthly redraw themselves');
ok(/^Loaded .+ · 11 favorites · 9 categories · 45 settings$/.test(g('cfg-excel-status').textContent),'setup shows what was loaded');
ok(!('cf_favorites' in store) && !('cf_cat_sub' in store) && !('cf_fav_deleted' in store),'old on-phone copies removed once all three tables load');
calls=[]; ok(await loadConfig()===false && calls.length===0,'on its own: at most every 10 minutes');
ok(await loadConfig(true)===true && calls.length===3 && toasts.at(-1)==='✓ Loaded from Excel','Reload button always reads');
// nothing changed → nothing redrawn
g('fav-chips').innerHTML='MARK'; await loadConfig(true); ok(g('fav-chips').innerHTML==='MARK','same setup → screen left alone');
renderFavChips();
// ── Failures keep what the phone has
mode.FavTable=404; await loadConfig(true); ok(favorites.length===11 && config.missing.join()==='FavTable' && toasts.at(-1)==='Not found in Excel: FavTable' && g('cfg-excel-status').textContent.includes('Not found in Excel: FavTable'),'a missing table is named; the saved favorites stay');
mode={AppSettings:500,FavTable:'drop',CatTable:503}; const before=store.cf_config; ok(await loadConfig(true)===false && store.cf_config===before && toasts.at(-1).startsWith('Couldn’t reach Excel'),'Excel unreachable → nothing changes'); mode={};
silentFail=true; calls=[]; ok(await loadConfig(true)===false && redirects===0 && calls.length===0 && toasts.at(-1).startsWith('Sign in again'),'expired sign-in → no redirect'); silentFail=false;
navigator.onLine=false; ok(await loadConfig(true)===false && toasts.at(-1).startsWith('Offline'),'offline → says so'); ok(await loadConfig()===false,'…quietly when automatic'); navigator.onLine=true;
msalInstance={getAllAccounts:()=>[]}; ok(await loadConfig(true)===false && toasts.at(-1)==='Sign in first','signed out → says so'); msalInstance=signedIn;
await loadConfig(true); ok(config.missing.length===0,'recovers on the next read');
// ── A reload while you're mid-entry
balanceCache={Credit:{Discover:500}}; futureCache={Credit:{sums:{},after:{},credits:{},cutoff:monthEndSerial(),count:0}}; balanceTs=Date.now();
switchSheet('Credit'); applyFav(favorites.findIndex(f=>f.label==='Discover Pay')); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r));
g('f-remarks').value='typed by hand'; const was=activeFav;
T.FavTable=[FX.FavTable[0],['New first','everyday','Credit','V','A','WF','','','',''],...FX.FavTable.slice(1)]; T.CatTable=[[...FX.CatTable[0],'Extra'],...FX.CatTable.slice(1).map(r=>[...r,''])];
await loadConfig(true);
ok(activeFav===was+1 && favorites[activeFav].label==='Discover Pay' && g('f-venue').value==='Discover' && g('f-cost').value==='500.00' && g('f-method').value==='Discover' && g('f-remarks').value==='typed by hand' && costNegative && g('pay-amt').style.display==='','list changed in Excel mid-entry → entry kept, favorite found again by name');
queueMode=true; queue=[]; g('f-date').value='2026-11-16'; await submitTransaction(); ok(queue.length===2 && queue[1].sheet==='Cash' && queue[1].payload.method==='Checking','…and it still writes both rows');
ok(Object.keys(catSub).includes('Extra') && g('fav-chips').innerHTML.includes('>New first<'),'new category and favorite show up');
applyFav(favorites.findIndex(f=>f.label==='Discover Pay')); await new Promise(r=>setImmediate(r)); T.FavTable=T.FavTable.filter(r=>r[0]!=='Discover Pay'); await loadConfig(true);
ok(activeFav===null && g('f-venue').value==='' && toasts.some(t=>t.startsWith('That favorite changed')),'the favorite you were using was removed in Excel → form cleared, never half an entry');
// typed entry (no favorite) with a custom category survives a reload
g('f-venue').value='Shop'; g('f-cost').value='5.00'; g('f-category').value='__other__'; onCategoryChange(); g('f-category-custom').value='Odd'; g('f-sub-custom').value='Bits';
T.CatTable=FX.CatTable; await loadConfig(true); ok(g('f-venue').value==='Shop' && g('f-category').value==='__other__' && g('f-category-custom').value==='Odd' && g('f-sub-custom').value==='Bits','hand-typed entry with a custom category survives');
resetForm();
// ── Draft across a sign-in redirect finds the favorite by name
T=JSON.parse(JSON.stringify(FX)); await loadConfig(true); applyFav(favorites.findIndex(f=>f.label==='Music')); saveFormDraft(); ok(JSON.parse(store.cf_draft).favLabel==='Music','draft remembers the favorite by name');
resetForm(); favorites=[...favorites].reverse(); restoreFormDraft(); ok(favorites[activeFav].label==='Music' && g('f-venue').value==='Tunes','restored onto the right favorite even if the order changed'); favorites=config.favorites; resetForm();
store.cf_draft=JSON.stringify({sheet:'Analysis',venue:'x'}); restoreFormDraft(); ok(g('f-venue').value==='','a draft for a view tab is ignored');
// ── Monthly needs its settings
T.AppSettings=FX.AppSettings.filter(r=>r[0]!=='gas.venue' && r[0]!=='rent.parking'); await loadConfig(true); viewMonth='2026-11'; expandedMonthly='gas'; renderMonthly(); let h=g('monthly-list').innerHTML;
ok(h.includes('add to AppSettings: gas.venue') && h.includes('add to AppSettings: rent.parking') && !h.includes('id="m-amt"'),'missing setting → the item says which, and can\'t be opened');
queue=[]; await commitMonthly('gas'); ok(queue.length===0 && toasts.at(-1).startsWith('Add the missing settings'),'…or added');
ok(g('monthly-totals').innerHTML.includes('Total apartment paid'),'totals still render');
T=JSON.parse(JSON.stringify(FX)); await loadConfig(true);
// ── Older checkmarks filed under another name move to "rent"
monthly={'2026-10':{gas:{amount:1,date:'d'},oldname:{amount:1170,date:'d'}},'2026-09':{rent:{amount:5,date:'d'},x:{amount:1,date:'d'}},'2026-08':{a:{amount:1,date:'d'},b:{amount:1,date:'d'}}}; migrateMonthly();
ok(monthly['2026-10'].rent.amount===1170 && !monthly['2026-10'].oldname && monthly['2026-09'].x && Object.keys(monthly['2026-08']).join()==='a,b' && JSON.parse(store.cf_monthly)['2026-10'].rent,'one unknown checkmark → rent; anything ambiguous left alone');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Warm-up + fast sync': async () => {
let toasts=[]; showToast=m=>toasts.push(m); confirm=()=>true; renderQueueReview=()=>{}; let loginCalls=0; login=()=>{loginCalls++;};
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
let silentFail=false; const acct={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>{ if(silentFail) throw new Error('x'); return {accessToken:'tok'}; }};
msalInstance=acct;
let COLS={CashTable:['Date','Venue','Cost','Checking'],CreditTable:['Date','Venue','Cost','WF']};
let log=[], selectMode='ok', colStatus=200, closeHang=false, gate=null;
global.fetch=async(url,opt={})=>{ const u=decodeURIComponent(url); const t=(u.match(/tables\/([^/]+)\//)||[])[1]; const r=(s,b)=>({ok:s<300,status:s,json:async()=>b});
  if(u.includes('/columns')){ const sel=u.endsWith('?$select=name'); log.push((sel?'colsel:':'col:')+t); if(gate) await gate;
    if(colStatus!==200) return r(colStatus,{error:{message:'nope'}});
    if(sel && selectMode==='400') return r(400,{}); return r(200,{value:COLS[t].map(name=>({name, values:[[1]]}))}); }
  if(u.endsWith('/createSession')){ log.push('open'); if(gate) await gate; return r(201,{id:'S'}); }
  if(u.endsWith('/closeSession')){ log.push('close'); if(closeHang) return new Promise(()=>{}); return r(204,{}); }
  if(u.endsWith('/rows/add')){ log.push('add:'+t+':'+JSON.parse(opt.body).values[0].length); return r(201,{}); }
  throw new Error('unexpected '+u); };
const reset=()=>{log=[];toasts=[];columnCache={};lastWarmUp=0;loginCalls=0;silentFail=false;selectMode='ok';colStatus=200;closeHang=false;gate=null;queue=[];saveQueue();};
const row=s=>({sheet:s,payload:{date:'2026-10-13',venue:'V',cost:1,inputted:'now'}});

// Warm-up
reset(); await warmUp();
ok(columnCache.CashTable?.join()==='Date,Venue,Cost,Checking' && columnCache.CreditTable?.length===4,'warm-up fills both tables\' columns');
ok(log.every(l=>l.startsWith('colsel:')) && log.length===2,'warm-up uses $select=name, read-only, both in parallel');
log=[]; await warmUp(); ok(log.length===0,'throttled: no second warm-up within 60s');
reset(); silentFail=true; await warmUp(); ok(loginCalls===0 && toasts.length===0 && log.length===0,'expired sign-in: no redirect, no toast, no request');
reset(); colStatus=500; await warmUp(); ok(toasts.length===0 && !columnCache.CashTable,'failure is silent');
reset(); msalInstance={getAllAccounts:()=>[]}; await warmUp(); ok(log.length===0,'signed out: skipped'); msalInstance=acct;
reset(); isSyncing=true; await warmUp(); ok(log.length===0,'mid-sync: skipped'); isSyncing=false;
reset(); navigator.onLine=false; await warmUp(); ok(log.length===0,'offline: skipped'); navigator.onLine=true;
reset(); await warmUp(); COLS.CashTable=['Date','Venue','Cost','Checking','New']; lastWarmUp=0; await warmUp();
ok(columnCache.CashTable.length===5,'warm-up picks up column changes made in Excel'); COLS.CashTable=['Date','Venue','Cost','Checking'];
reset(); let res; gate=new Promise(r=>res=r); const p=warmUp(); isSyncing=true; res(); await p; ok(!columnCache.CashTable,'columns not swapped if a sync started meanwhile'); isSyncing=false;
reset(); openQueueReview(); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r)); ok(log.length===2,'opening pending triggers warm-up');

// $select fallback
reset(); selectMode='400'; const c=await getTableColumns('CashTable');
ok(c?.length===4 && log.join()==='colsel:CashTable,col:CashTable','$select rejected → falls back to plain request');
reset(); colStatus=404; const c2=await getTableColumns('CashTable'); ok(c2===null && toasts.some(t=>t.includes('Table error')),'real table error still reported');

// Sync: parallel setup, writes after, no wait on close
reset(); queue=[row('Cash'),row('Credit')]; saveQueue(); gate=new Promise(r=>res=r);
const sp=syncQueue(); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r));
ok(log.includes('open') && log.includes('colsel:CashTable') && log.includes('colsel:CreditTable') && !log.some(l=>l.startsWith('add')),'session + both column lookups start together, before any write');
res(); await sp;
ok(log.filter(l=>l.startsWith('add')).length===2 && log.filter(l=>l.startsWith('colsel')).length===2,'columns fetched once each, 2 writes');
ok(queue.length===0 && toasts.at(-1)==='✓ All synced','synced');
reset(); closeHang=true; queue=[row('Cash'),row('Credit')]; saveQueue();
const done=await Promise.race([syncQueue().then(()=>'done'), new Promise(r=>setTimeout(()=>r('hung'),300))]);
ok(done==='done' && toasts.at(-1)==='✓ All synced' && !isSyncing,'✓ shown without waiting for close');
reset(); queue=[row('Cash')]; saveQueue(); await syncQueue(); ok(!log.includes('open') && log.filter(l=>l.startsWith('add')).length===1,'single table: no session');
reset(); await warmUp(); log=[]; queue=[row('Cash')]; saveQueue(); await syncQueue(); ok(log.join()==='add:CashTable:4','after warm-up, sync is just the write');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'HSA Reimb links': async () => {
let msg='',mtype=''; showToast=(m,t)=>{msg=m;mtype=t;}; global.setTimeout=fn=>fn(); confirm=()=>true;
const g=id=>document.getElementById(id);
msalInstance={getAllAccounts:()=>[{}]}; queueMode=true; activeFav=null;
const setForm=(sheet,cost,method,cat,venue='CVS',neg=false)=>{ currentSheet=sheet; g('f-date').value='2026-10-14'; g('f-cost').value=cost; costNegative=neg;
  g('f-venue').value=venue; g('f-activities').value='Rx'; g('f-method').value=method; g('f-category').value=cat; g('f-sub').value=''; g('f-sub2').value=''; g('f-sub3').value=''; g('f-remarks').value='r'; };
const reset=()=>{queue=[];saveQueue();msg='';};
// Rule
ok(needsHsaLink({category:'Health',method:'Chase'}) && needsHsaLink({category:' health ',method:'Checking'}) && needsHsaLink({category:'Health',method:''}),'Health + non-Optum → link (any case, blank method too)');
ok(!needsHsaLink({category:'Health',method:'Optum'}) && !needsHsaLink({category:'Health',method:'optum'}) && !needsHsaLink({category:'Health',method:'HSA Reimb'}) && !needsHsaLink({category:'Dr',method:'Chase'}),'Optum / HSA Reimb itself / other categories → no link');
ok(SHEETS.Cash.methods.indexOf('HSA Reimb')===SHEETS.Cash.methods.indexOf('Optum')+1,'HSA Reimb in Cash methods right after Optum');
// Submit
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction();
let L=queue[1]; ok(queue.length===2 && queue[0].sheet==='Credit' && L.sheet==='Cash' && L.hsa,'Credit Health expense → + Cash HSA row');
ok(L.payload.method==='HSA Reimb' && L.payload.cost===-45 && L.payload.category==='' && L.payload.venue==='CVS' && L.payload.activities==='Rx' && L.payload.date==='2026-10-14' && L.payload.inputted===queue[0].payload.inputted,'link: HSA Reimb, −45, no category, venue/activities/date copied');
ok(queue[0].group.id===L.group.id && L.group.n===2,'grouped 1 of 2');
reset(); setForm('Cash','30.00','Checking','Health','Clinic',true); await submitTransaction(); ok(queue[1].payload.cost===30,'refund (−30) → link +30');
reset(); setForm('Cash','30.00','Optum','Health'); await submitTransaction(); ok(queue.length===1,'Optum Health → no link');
reset(); setForm('Credit','12.00','Chase','F&D'); await submitTransaction(); ok(queue.length===1,'non-Health → no link');
// Pending edits
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction();
toggleQueueEdit(0); g('q-sheet').value='Credit'; g('q-date').value='2026-10-15'; g('q-amount').value='50.00'; g('q-venue').value='Walgreens'; g('q-activities').value='Rx2'; g('q-method').value='Chase'; g('q-category').value='Health'; ['q-sub','q-sub2','q-sub3','q-remarks'].forEach(i=>g(i).value='');
saveQueueEdit(); L=queue[1];
ok(L.payload.cost===-50 && L.payload.date==='2026-10-15' && L.payload.venue==='Walgreens' && msg.includes('HSA Reimb row too'),'edit cost/date/venue → link follows');
toggleQueueEdit(0); g('q-amount').value='50.00'; g('q-method').value='Optum'; g('q-category').value='Health'; saveQueueEdit();
ok(queue.length===1 && !queue[0].group && msg.includes('removed'),'switch to Optum → link removed, group cleared');
toggleQueueEdit(0); g('q-method').value='Chase'; g('q-category').value='Health'; saveQueueEdit();
ok(queue.length===2 && queue[1].hsa && queue[1].payload.cost===-50 && queue[0].group.n===2 && msg.includes('added'),'back to Health non-Optum → link re-added after it');
toggleQueueEdit(0); g('q-category').value='F&D'; saveQueueEdit(); ok(queue.length===1,'category away from Health → removed');
// Edit non-Health into Health
reset(); setForm('Credit','8.00','Chase','F&D'); await submitTransaction(); toggleQueueEdit(0); g('q-amount').value='8.00'; g('q-method').value='Chase'; g('q-category').value='Health'; saveQueueEdit();
ok(queue.length===2 && queue[1].payload.cost===-8 && queue[0].group.id===queue[1].group.id,'edited into Health → link added');
// Delete
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction(); setForm('Cash','5.00','Checking','F&D'); await submitTransaction();
deleteQueueItem(0); ok(queue.length===1 && queue[0].payload.venue==='CVS' && queue[0].payload.method==='Checking','deleting expense deletes its HSA row, others kept');
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction(); deleteQueueItem(1);
ok(queue.length===1 && !queue[0].group && !queue[0].hsaLinked,'deleting HSA row alone keeps expense, unlinks');
toggleQueueEdit(0); g('q-amount').value='60.00'; g('q-method').value='Chase'; g('q-category').value='Health'; saveQueueEdit();
ok(queue.length===2 && queue[1].payload.cost===-60,'after removing it by hand, a later edit re-adds a fresh link');
// Link already synced (only the expense still pending)
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction(); queue.splice(1,1); saveQueue();
toggleQueueEdit(0); g('q-amount').value='45.00'; g('q-method').value='Chase'; g('q-category').value='Health'; g('q-remarks').value='x'; saveQueueEdit();
ok(queue.length===1 && mtype==='success','synced link, cost unchanged → no warning, no duplicate');
toggleQueueEdit(0); g('q-amount').value='55.00'; saveQueueEdit(); ok(queue.length===1 && mtype==='error' && msg.includes('update its cost there'),'synced link + cost change → warn, no duplicate');
toggleQueueEdit(0); g('q-category').value='F&D'; saveQueueEdit(); ok(mtype==='error' && msg.includes('remove it there'),'synced link + no longer Health → warn');
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction(); queue.splice(1,1); msg=''; deleteQueueItem(0); ok(queue.length===0 && msg.includes('already in Excel'),'delete expense whose link synced → warn');
// Editing the HSA row by hand
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction(); toggleQueueEdit(1); g('q-sheet').value='Cash'; g('q-amount').value='40.00'; editNegative=true; g('q-method').value='HSA Reimb'; g('q-category').value=''; saveQueueEdit();
ok(queue[1].payload.cost===-40 && queue[0].hsaLinked.cost===-40,'hand-edited HSA row is kept as the reference');
// Persistence
reset(); setForm('Credit','45.00','Chase','Health'); await submitTransaction(); const re=JSON.parse(store.cf_queue); ok(re[1].hsa===true && re[0].hsaLinked.cost===-45,'link flags survive reload');
// Renders
renderQueueReview(); toggleQueueEdit(0); ok(g('queue-list').innerHTML.includes('row updates to match'),'edit form explains auto-update'); toggleQueueEdit(0);
// Transfer pay-back
let t=transferRows('HSA Reimb','Checking',120,'','');
ok(t.length===3 && t.every(r=>r.venue==='Optum'),'pay-back → 3 rows, venue Optum');
ok(t[0].method==='Optum'&&t[0].cost===120 && t[1].method==='Checking'&&t[1].cost===-120 && t[2].method==='HSA Reimb'&&t[2].cost===120 && !t[2].category,'Optum −, Checking +, HSA Reimb −');
t=transferRows('HSA Reimb','Fidelity',50,'',''); ok(t[0].category==='Savings' && t[0].activities==='Transfer (+savings)' && t[2].category==='','savings rule still applies to the Optum side');
t=transferRows('Optum','Checking',50,'',''); ok(t.length===2 && t[1].venue==='Wells Fargo','plain Optum transfer unchanged (2 rows)');
ok(transferProblem('HSA Reimb','Optum') && !transferProblem('HSA Reimb','Cash') && transferProblem('Checking','Checking'),'HSA → Optum blocked');
ok(!needsCashVenue('HSA Reimb','Cash') && needsCashVenue('Venmo','Cash'),'pay-back to Cash needs no cash venue');
// commitTransfer + opens
g('t-from').options=[]; g('t-from').innerHTML=''; openTransfer(); ok(g('t-from').innerHTML.includes('HSA Reimb') && !g('t-to').innerHTML.includes('HSA Reimb'),'HSA Reimb only in From list');
reset(); closePanel=()=>{}; g('t-from').value='HSA Reimb'; g('t-to').value='Cash'; g('t-amount').value='75.00'; g('t-date').value='2026-10-14'; g('t-cash-venue').value=''; g('t-remarks').value='';
await commitTransfer(); ok(queue.length===3 && queue.every(q=>q.sheet==='Cash' && q.group.n===3) && queue[1].payload.method==='Cash' && queue[1].payload.venue==='Optum','pay-back to Cash commits 3 grouped rows');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Month-End bill + Pay amounts': async () => {
showToast=()=>{}; global.setTimeout=fn=>fn();
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; accessToken='tok';
const g=id=>document.getElementById(id);
const t=new Date(); const today=excelSerial(t.getFullYear(),t.getMonth()+1,t.getDate());
const pm=new Date(t.getFullYear(),t.getMonth(),0); const C=excelSerial(pm.getFullYear(),pm.getMonth()+1,pm.getDate());
const ME=excelSerial(t.getFullYear(),t.getMonth()+2,0); const TODAY_IN_MONTH=Math.min(today+10,ME);
ok(monthEndSerial()===C && monthEndDate().getDate()===pm.getDate(),'cutoff = last day of last month');
ok(new Date(2026,11,0).getDate()===30 && new Date(2026,10,0).getDate()===31 && new Date(2027,2,0).getDate()===28 && new Date(2027,0,0).getFullYear()===2026,'month ends: Dec→Nov 30, Nov→Oct 31, Mar→Feb 28, Jan→Dec 31 prev year');
// Discover: before cutoff stuff, on-cutoff purchase, after: +200 purchase, −500 payment, −10 cash back, −100 future payment; last 1000
// Chase: after −125 payment; last −25.  WF: nothing after; last 300. Savor1: 0.
const CREDIT={cols:['Date','Venue','Activities','Method','Cost','Discover','WF','Chase','Savor1','Total'],
  date:[C-5,C,C+1,C+1,C+1,TODAY_IN_MONTH,C+1,'junk',C-1,ME+5], method:['Discover','Discover','Discover','discover','Discover','Discover','Chase','Discover','WF','Discover'],
  cost:[40,60,200,-500,-10,-100,-125,999,5,-300], last:['','','','','',700,300,-25,0,975]};
const CASH={cols:['Date','Method','Cost','Checking'],date:[C+1],method:['Checking'],cost:[10],last:['','','',500]};
const T={CashTable:CASH,CreditTable:CREDIT}; let failCols=false, calls=0;
global.fetch=async(url)=>{ calls++; const u=decodeURIComponent(url); const tn=(u.match(/tables\/([^/]+)\//)||[])[1]; const tb=T[tn]; const r=(s,b)=>({ok:s<300,status:s,json:async()=>b});
  if(/\/columns(\?\$select=name)?$/.test(u)) return r(200,{value:tb.cols.map(name=>({name}))});
  if(u.endsWith('/range/lastRow')) return r(200,{values:[tb.last]});
  const col=(u.match(/columns\/([^/]+)\/dataBodyRange/)||[])[1];
  if(col){ if(failCols) return r(500,{}); const key={Date:'date',Method:'method',Cost:'cost'}[col]; return r(200,{values:tb[key].map(v=>[v])}); }
  throw new Error('unexpected '+u); };
await loadBalances();
const f=futureCache.Credit;
ok(f.after.discover===-710 && f.credits.discover===-610 && f.cutoff===C,'after-cutoff: all rows rewind the balance; only this month\'s credits pay the bill (next-month −300 excluded)');
const fut = TODAY_IN_MONTH>today ? -400 : -300;
ok(f.sums.discover===fut,'"as of today" still only counts future rows');
let b=billFor('Discover'); ok(b.owed===1410 && b.left===800,'Discover: owed 1,410 (balance on cutoff), left 800 after payment+cash back+scheduled this-month payment');
ok(billFor('WF').owed===300 && billFor('WF').left===300,'no activity since → left = owed');
ok(billFor('Chase').owed===100 && billFor('Chase').left===-25,'Chase overpaid by 25');
ok(billFor('Savor1').owed===0 && billFor('Savor1').left===0,'Savor1 zero');
ok(!g('balance-main').innerHTML.includes('left $'),'Today mode: no month-end info');
setBalMode('month'); let h=g('balance-main').innerHTML;
ok(g('bal-more-toggle').style.display==='none' && g('balance-expand').style.display==='none' && !h.includes('Checking'),'Month-End: credit cards only, More accounts hidden');
ok(g('bal-note').textContent.startsWith('Bill as of '+pm.toLocaleDateString('en-US',{month:'short',day:'numeric'})),'Month-End note shows cutoff');
ok(h.includes('$1,410.00') && h.includes('left $800.00') && h.includes('overpaid $25.00') && h.includes('paid ✓'),'bill cards: owed, left, overpaid, paid');
ok(/bal-card full"><div class="bal-label">Total<\/div><div class="bal-value positive">\$1,810\.00<\/div><div class="bal-left due">left \$1,100\.00/.test(h),'Total row: owed 1,810, left 1,100 (overpaid card not offset)');
setBalMode('today'); ok(g('bal-more-toggle').style.display==='' && g('balance-main').innerHTML.includes('Checking'),'back to Today: everything shown again');
ok(g('balance-main').innerHTML.includes('$'+fmtMoney(700-fut)),'Today value for Discover (excludes future-dated rows)');
setBalMode('all'); ok(g('balance-main').innerHTML.includes('$700.00') && !g('balance-main').innerHTML.includes('left $'),'All = raw, no bill'); setBalMode('today');
// Pay autofill source
let calls0=calls; ok(await getPayAmount('Discover','month')===800 && calls===calls0,'Pay Month-End = left to pay, fresh cache (no requests)');
ok(await getPayAmount('Discover','today')===700-fut,'Pay Today = balance as of today');
ok(await getPayAmount('Chase','month')===0 && await getPayAmount('Chase','today')===0,'overpaid → 0');
balanceTs=0; calls0=calls; ok(await getPayAmount('WF','month')===300 && calls>calls0,'stale cache → refetches');
futureCache.Credit.cutoff=C-30; calls0=calls; await getPayAmount('WF','month'); ok(calls>calls0,'new month → refetches');
// Pay buttons in the form
const fi=favorites.findIndex(x=>x.autoFillCol==='Discover'); currentSheet='Credit'; switchSheet&&0;
applyFav(fi); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r));
ok(g('pay-amt').style.display==='' && payKind==='month' && g('f-cost').value==='800.00' && costNegative,'Pay favorite: buttons shown, Month-End default, amount filled, sign −');
setPayAmount('today'); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r));
ok(g('f-cost').value===fmtMoney(700-fut) && payKind==='today','tap Today → fills today amount');
resetForm(); ok(g('pay-amt').style.display==='none' && payKind==='month','reset hides buttons');
const nf=favorites.findIndex(x=>!x.autoFillCol && x.sheet==='Credit'); applyFav(fi); applyFav(nf); ok(g('pay-amt').style.display==='none','non-Pay favorite hides buttons');
resetForm();
failCols=true; balanceTs=0; ok(await getPayAmount('Discover','month')===700 && await getPayAmount('Discover','today')===700,'dates unreadable → both fall back to latest balance');
await loadBalances(); setBalMode('month'); ok(g('balance-main').innerHTML.includes('isn’t available'),'dates unreadable → Month-End says so'); setBalMode('today'); failCols=false;
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Retirement item + Retire table': async () => {
let msg=''; showToast=m=>msg=m; global.setTimeout=fn=>fn(); confirm=()=>true; renderQueueReview=()=>{};
const g=id=>document.getElementById(id);
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; accessToken='tok';
ok(tableFor('Retire')==='RetireTable' && tableFor('Cash')==='CashTable','Retire defaults to RetireTable, others from setup');
// Setup field
g('cfg-client-id').value=''; g('cfg-file-path').value=''; g('cfg-cash-table').value='CashTable'; g('cfg-credit-table').value='CreditTable'; g('cfg-retire-table').value='MyRetire';
saveConfig(); ok(store.cf_table_retire==='MyRetire' && tableFor('Retire')==='MyRetire','setup saves Retire table name');
openConfig(); ok(g('cfg-retire-table').value==='MyRetire' && g('cfg-cash-table').value==='CashTable','setup shows saved names');
g('cfg-retire-table').value=''; saveConfig(); ok(store.cf_table_retire==='MyRetire','blank field keeps the saved name (same as Cash/Credit)');
delete store.cf_table_retire;
// Monthly opens last month through the 8th
const RD=Date; const at=(y,m,d)=>{ global.Date=class extends RD{constructor(...a){ if(a.length) super(...a); else super(y,m,d);} static now(){return new RD(y,m,d).getTime();}}; const k=defaultMonthKey(); global.Date=RD; return k; };
ok(at(2026,10,8)==='2026-10' && at(2026,10,9)==='2026-11' && at(2027,0,3)==='2026-12','Monthly: last month through the 8th (incl. January → December)');
const item=MONTHLY_ITEMS.find(i=>i.key==='retire');
ok(item && item.sheet==='Retire' && item.due==='eom' && monthlyDefault(item,'2026-10')===950 && monthlyDefault(item,'2026-11')===1200,'retirement item: Retire table, month end, default total by month');
// Render form
viewMonth='2026-10'; expandedMonthly='retire'; renderMonthly(); const h=g('monthly-list').innerHTML;
ok(h.includes('id="m-amt-0"') && h.includes('<label>Employer</label>') && h.includes('<label>Personal</label>') && h.includes('value="800.00"') && h.includes('value="150.00"') && h.includes('m-received') && h.includes('value="2026-10-31"'),'form: two amounts, received, date = Oct 31');
// Commit (queue mode)
const COLS=['Date','Venue','Activities','Method','Category','Sub','Sub2','Sub3','Cost','403b','Remarks','Inputted','Uploaded']; let adds=[];
global.fetch=async(url,opt={})=>{ const u=decodeURIComponent(url); const r=(s,b)=>({ok:s<300,status:s,json:async()=>b});
  if(/\/columns(\?\$select=name)?$/.test(u)) return r(200,{value:COLS.map(name=>({name}))});
  if(u.endsWith('/createSession')) return r(201,{id:'S'}); if(u.endsWith('/closeSession')) return r(204,{});
  if(u.endsWith('/rows/add')){ adds.push({t:(u.match(/tables\/([^/]+)\//)||[])[1],v:JSON.parse(opt.body).values}); return r(201,{}); }
  throw new Error(u); };
queueMode=true; queue=[]; monthly={};
g('m-amt-0').value='800.00'; g('m-amt-1').value='150.00'; g('m-date').value='2026-10-31'; g('m-received').value='2026-11-05';
await commitMonthly('retire');
ok(queue.length===2 && queue.every(q=>q.sheet==='Retire') && queue[0].group.n===2,'2 grouped Retire rows queued');
const [a,b]=queue.map(q=>q.payload);
ok(a.venue==='RetireCo' && a.activities==='Employer Contribution' && a.method==='403b' && a.cost===-800 && a.category==='' && a.date==='2026-10-31','employer row (wording from settings)');
ok(b.activities==='Personal Contribution' && b.cost===-150,'personal row');
ok(a.remarks==='Received Nov 5, 2026' && b.remarks===a.remarks,'Received date in Remarks with year');
ok(monthly['2026-10'].retire.amount===950 && history[0].sheet==='Retire' && history[0].cost===-950 && history[0].venue==='Retirement','checkmark + history');
// No received → blank remarks; edited amounts
queue=[]; monthly={}; expandedMonthly='retire'; renderMonthly(); g('m-amt-0').value='900'; g('m-amt-1').value='150.00'; g('m-date').value='2026-10-31'; g('m-received').value='';
await commitMonthly('retire'); ok(queue[0].payload.cost===-9 && queue[0].payload.remarks==='' ,'no-dot input = cents (900 → 9.00), no Received → blank remarks');
queue=[]; monthly={}; expandedMonthly='retire'; renderMonthly(); g('m-amt-0').value='abc'; g('m-amt-1').value='1.00'; g('m-date').value='2026-10-31';
await commitMonthly('retire'); ok(queue.length===0 && msg==='Enter a valid amount','invalid amount blocked');
// Sync goes to RetireTable
queue=[]; monthly={}; expandedMonthly='retire'; renderMonthly(); g('m-amt-0').value='800.00'; g('m-amt-1').value='150.00'; g('m-date').value='2026-10-31'; g('m-received').value='';
await commitMonthly('retire'); columnCache={}; await syncQueue();
ok(adds.length===1 && adds[0].t==='RetireTable' && adds[0].v.length===2 && adds[0].v[0][9]===null && adds[0].v[0][8]===-800 && queue.length===0,'syncs both rows to RetireTable in one write, formula column untouched');
// Other monthly items unaffected
expandedMonthly='internet'; renderMonthly(); ok(g('monthly-list').innerHTML.includes('id="m-amt"') && !g('monthly-list').innerHTML.includes('m-received'),'single-amount items unchanged');
// Edit form offers Retire; favorites can't target it
queue=[{sheet:'Retire',payload:{date:'2026-10-31',venue:'RetireCo',method:'403b',cost:-1}}]; ok(queueEditForm(queue[0]).includes('<option value="Retire" selected>') && queueEditForm(queue[0]).includes('value="403b"'),'pending edit shows Retire + 403b');
ok(!FORM_SHEETS.includes('Retire'),'no Retire entry tab / favorites');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Monthly amounts + wording by month': async () => {
let msg=''; showToast=m=>msg=m; global.setTimeout=fn=>fn(); confirm=()=>true; renderQueueReview=()=>{};
const g=id=>document.getElementById(id); queueMode=true; msalInstance={getAllAccounts:()=>[{}]};
const I=k=>MONTHLY_ITEMS.find(i=>i.key===k);
// Amounts follow the page's month (made-up setup: HSA 200 → 250 from 2027-01, Fidelity 300 → 400 and savings 2,000 from 2026-11)
ok(monthlyDefault(I('hsaMine'),'2026-12')===200 && monthlyDefault(I('hsaMine'),'2027-01')===250 && monthlyDefault(I('hsaMine'),'2027-06')===250,'amount changes from its From month onward');
ok(monthlyDefault(I('fidelity'),'2026-10')===300 && monthlyDefault(I('fidelity'),'2026-11')===400,'older pages keep the older amount');
ok(monthlyDefault(I('savings'),'2026-10')===undefined && monthlyDefault(I('savings'),'2026-11')===2000,'no amount before its first From → blank field');
ok(monthlyDefault(I('gas'),'2026-11')===undefined && monthlyDefault(I('hsaEmployer'),'2030-01')===100,'no amount set → blank; no From → always applies');
const run=async(key,month,amt)=>{ queue=[]; monthly={}; viewMonth=month; expandedMonthly=key; renderMonthly(); g('m-amt').value=amt; g('m-date').value=monthDueDate(month,I(key).due); await commitMonthly(key); return queue.map(q=>[q.payload.venue,q.payload.activities,q.payload.method,q.payload.category,q.payload.cost]); };
viewMonth='2026-11'; expandedMonthly='savings'; renderMonthly(); ok(g('monthly-list').innerHTML.includes('value="2,000.00"'),'page pre-fills the amount');
let r=await run('savings','2026-11','2000.00'); ok(JSON.stringify(r.map(x=>[x[2],x[4]]))===JSON.stringify([['Checking',2000],['SavingsF',-600],['SavingsCO',-1400]]),'savings: Family fixed (600), Discretionary the rest');
r=await run('savings','2026-12','2300.00'); ok(r[1][4]===-600 && r[2][4]===-1700,'edited total: the rest goes to Discretionary');
r=await run('savings','2026-10','1000.00'); ok(r[1][4]===-500 && r[2][4]===-500,'older page keeps the older Family amount');
r=await run('savings','2026-11','550.00'); ok(r.length===0 && msg.includes('600.00'),'total below Family is blocked');
r=await run('savings','2026-10','550.00'); ok(r.length===3,'…against that month\'s Family amount');
viewMonth='2027-01'; expandedMonthly='hsaMine'; renderMonthly(); ok(g('monthly-list').innerHTML.includes('value="250.00"'),'later page pre-fills the newer amount');
viewMonth='2026-11'; expandedMonthly='fidelity'; renderMonthly(); ok(g('monthly-list').innerHTML.includes('value="400.00"') && g('monthly-list').innerHTML.includes('$400.00 · due'),'collapsed line shows the amount too');
// Wording comes from the setup
r=await run('pay','2026-11','4000.00'); ok(JSON.stringify(r)===JSON.stringify([['Acme','Pay','Checking','Pay',-4000]]),'pay row: venue/activities from settings, negative');
r=await run('hsaMine','2026-11','200.00'); ok(JSON.stringify(r)===JSON.stringify([['Acme','My contribution','Optum','HSA',-200]]) ,'HSA row');
r=await run('rent','2026-11','1170.00'); ok(r.length===3 && r.map(x=>x[4]).join()==='1000,100,70' && r.every(x=>x[0]==='Oak Homes'),'rent bill: three rows');
r=await run('rent','2026-11','1050.00'); ok(r.length===0 && msg.includes('1,100.00'),'rent bill below rent + parking is blocked');
viewMonth='2026-11'; expandedMonthly=null; renderMonthly(); const hh=g('monthly-list').innerHTML;
ok(['Gas bill','Internet','Electric','Employer HSA','My HSA','HSA interest','Pay','Rent bill','Phone','Savings (total)','Fidelity','Retirement'].every(l=>hh.includes(`<span class="m-title">${l}</span>`)),'every item labelled from the setup');
// Apartment totals land on the last of the four bills
monthly={ '2026-11': { gas:{amount:40,date:'x'}, electric:{amount:60,date:'x'}, internet:{amount:30,date:'x'} } }; queue=[]; viewMonth='2026-11'; expandedMonthly='rent'; renderMonthly();
g('m-amt').value='1170.00'; g('m-date').value='2026-12-01'; await commitMonthly('rent');
ok(queue[2].payload.remarks==='Gas + electric + sewer = 170.00 · Total apartment paid = 1,300.00' && queue[1].payload.remarks==='Total Oak Homes payment = 1,170.00','last apartment bill carries the totals');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Favorites + Quick add sections': async () => {
showToast=()=>{}; const g=id=>document.getElementById(id);
ok(SHEETS.Cash.methods[0]==='Venmo' && SHEETS.Cash.methods.length===11 && new Set(SHEETS.Cash.methods).size===11,'Venmo first in Cash, nothing lost');
ok(favorites.map(f=>f.label).join()==='Grocer,Park,Cloud,Music,Discover CB,Chase CB,DeltaAmex CB,Discover Pay,Chase Pay,DeltaAmex Pay,Coffee','favorites come from the setup, in row order');
// Chips: one line per section, for the current table only
currentSheet='Credit'; switchSheet('Credit'); renderFavChips(); let grp=g('fav-chips').innerHTML.split('<div class="fav-group">').slice(1);
ok(grp.length===4,'four sections');
ok(grp[0].includes('>Grocer<') && grp[0].includes('>Park<') && !grp[0].includes('Cloud') && !grp[0].includes('Coffee'),'everyday line (Cash favorites not shown on Credit)');
ok(grp[1].includes('>Cloud<') && grp[1].includes('>Music<') && !grp[1].includes('CB'),'subscriptions line');
ok(grp[2].includes('Discover CB') && !grp[2].includes('Pay') && grp[3].includes('Chase Pay') && !grp[3].includes('CB'),'cash back / payments lines');
switchSheet('Cash'); ok(g('fav-chips').innerHTML.includes('>Coffee<') && !g('fav-chips').innerHTML.includes('Grocer'),'Cash tab shows Cash favorites');
// A favorite with a blank Section, or one the app doesn't know, sits with everyday and is never hidden
favorites=parseFavorites([FX.FavTable[0],['Mine','something else','Credit','V','A','WF','','','','']]); switchSheet('Credit'); renderFavChips();
ok(g('fav-chips').innerHTML.split('<div class="fav-group">')[1].includes('>Mine<') && favorites[0].group===undefined,'unknown section → everyday');
favorites=config.favorites; renderFavChips();
// Apply + submit: text you finish yourself, cost typed
queueMode=true; msalInstance={getAllAccounts:()=>[{}]}; queue=[];
applyFav(favorites.findIndex(f=>f.label==='Park'));
ok(g('f-activities').value==='Parking by ' && g('f-cost').value==='' && g('f-venue').value==='City Parking' && g('f-method').value==='Savor1' && g('f-category').value==='Transport','chip fills the form, cost left blank');
g('f-activities').value='Parking by the stadium'; g('f-cost').value='12.00'; g('f-sub').value='Parking'; g('f-date').value='2026-11-03';
await submitTransaction(); const q=queue[0].payload; ok(queue.length===1 && q.activities==='Parking by the stadium' && q.cost===12 && q.method==='Savor1','entry written');
applyFav(favorites.findIndex(f=>f.label==='Music')); ok(g('f-cost').value==='9.99' && g('f-remarks').value==='25th' && !costNegative,'subscription: fixed cost and remark filled');
resetForm();
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'DeltaAmexGold card': async () => {
showToast=()=>{}; const g=id=>document.getElementById(id);
ok(SHEETS.Credit.methods.at(-1)==='DeltaAmexGold' && SHEETS.Credit.methods.length===5,'DeltaAmexGold appended to Credit methods');
ok(MAIN_BALANCES.filter(b=>b.sheet==='Credit').map(b=>b.col).join()==='Discover,WF,Chase,Savor1,DeltaAmexGold','main balances follow the card list');
currentSheet='Credit'; g('f-method').value='DeltaAmexGold'; onMethodChange(); ok(g('account-preview').innerHTML.includes('DeltaAmexGold'),'method preview');
const C=monthEndSerial();
balanceCache={Cash:{Checking:1,Cash:2,Venmo:3},Credit:{Discover:100,WF:0,Chase:0,Savor1:0,DeltaAmexGold:250.5,Venmo:9,Total:0}};
futureCache={Cash:{sums:{},count:0,after:{},credits:{},cutoff:C},Credit:{sums:{deltaamexgold:50},count:1,after:{deltaamexgold:50.5},credits:{deltaamexgold:-100},cutoff:C}};
setBalMode('today'); let h=g('balance-main').innerHTML;
ok(/bal-card half"><div class="bal-label">DeltaAmexGold<\/div><div class="bal-value positive">\$200\.50<\/div><\/div><div class="bal-card half"><div class="bal-label">Total Credit<\/div><div class="bal-value positive">\$300\.50/.test(h),'Today: 250.50 − 50 future; DeltaAmexGold bottom-left, Total Credit bottom-right (3 rows)');
ok(!g('balance-expand').innerHTML.includes('DeltaAmexGold'),'not repeated under More accounts');
setBalMode('month'); h=g('balance-main').innerHTML;
ok(/DeltaAmexGold<\/div><div class="bal-value positive">\$200\.00<\/div><div class="bal-left due">left \$100\.00/.test(h),'Month-End: owed 200 (250.50 − 50.50 since), left 100');
ok(h.includes('<div class="bal-card half"><div class="bal-label">Total</div><div class="bal-value positive">$300.00'),'Month-End total includes it, beside DeltaAmexGold');
delete balanceCache.Credit.DeltaAmexGold; h=renderBill(); ok(!h.includes('DeltaAmexGold') && h.includes('Discover') && /Savor1<\/div><div class="bal-value[^"]*">\$0\.00<\/div><div class="bal-left done">paid ✓<\/div><\/div><div class="bal-card full"><div class="bal-label">Total/.test(h),'card missing in Excel → skipped, others shown, Total back to full row');
setBalMode('today');
ok(totalSpan(5)==='half' && totalSpan(4)==='full','Total width follows card count');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Payment + cash back favorites': async () => {
showToast=()=>{}; global.setTimeout=fn=>fn(); const g=id=>document.getElementById(id);
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; queueMode=true;
const f=favorites.find(x=>x.label==='DeltaAmex Pay'); ok(f.group==='pay' && f.autoFillCol==='DeltaAmexGold' && SHEETS.Credit.methods.includes(f.method) && f.linked.sheet==='Cash' && f.linked.method==='Checking' && f.linked.category==='Credit' && f.linked.venue===f.venue && f.linked.activities===f.activities,'a payment row gets its amount lookup and Checking side automatically');
switchSheet('Credit'); renderFavChips(); const last=g('fav-chips').innerHTML.split('<div class="fav-group">').at(-1); ok(last.includes('>Chase Pay<') && last.indexOf('Chase Pay')<last.indexOf('DeltaAmex Pay'),'chips keep the row order');
const C=monthEndSerial(); balanceCache={Credit:{DeltaAmexGold:250.5}}; futureCache={Credit:{sums:{deltaamexgold:20},after:{deltaamexgold:50.5},credits:{deltaamexgold:-100},cutoff:C,count:1}}; balanceTs=Date.now();
applyFav(favorites.findIndex(x=>x.label==='DeltaAmex Pay')); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r));
ok(g('f-cost').value==='100.00' && g('f-method').value==='DeltaAmexGold' && costNegative,'Month-End amount: 200 owed − 100 credits = 100');
setPayAmount('today'); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r)); ok(g('f-cost').value==='230.50','Today amount: 250.50 − 20 future');
g('f-date').value='2026-11-16'; g('f-category').value=''; queue=[]; await submitTransaction();
const q=queue.map(x=>[x.sheet,x.payload.venue,x.payload.activities,x.payload.method,x.payload.category,x.payload.cost]);
ok(JSON.stringify(q)===JSON.stringify([['Credit','Amex','Amex Payment','DeltaAmexGold','',-230.5],['Cash','Amex','Amex Payment','Checking','Credit',230.5]]),'writes the Credit row + the Checking row');
resetForm(); switchSheet('Credit'); renderFavChips(); const cbLine=g('fav-chips').innerHTML.split('<div class="fav-group">').at(-2);
ok(cbLine.indexOf('Chase CB')<cbLine.indexOf('DeltaAmex CB') && cbLine.indexOf('DeltaAmex CB')>0 && !cbLine.includes('Pay'),'cash back line');
applyFav(favorites.findIndex(x=>x.label==='DeltaAmex CB')); ok(costNegative && g('f-cost').value==='' && g('f-method').value==='DeltaAmexGold' && g('f-venue').value==='Amex' && g('f-activities').value==='Statement credit' && g('pay-amt').style.display==='none','cash back chip: negative sign, blank amount, no Pay buttons');
g('f-cost').value='25.00'; g('f-date').value='2026-11-16'; queue=[]; await submitTransaction(); ok(queue.length===1 && queue[0].sheet==='Credit' && queue[0].payload.cost===-25 && queue[0].payload.method==='DeltaAmexGold','writes one Credit row, −25');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Retire + Travel balances': async () => {
let toasts=[]; showToast=m=>toasts.push(m); global.setTimeout=fn=>fn();
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; accessToken='tok';
const g=id=>document.getElementById(id);
const t=new Date(); const today=excelSerial(t.getFullYear(),t.getMonth()+1,t.getDate());
const T={
 CashTable:{cols:['Date','Method','Cost','Checking','Savings'],date:[today],method:['Checking'],cost:[1],last:['','','',500,300]},
 CreditTable:{cols:['Date','Method','Cost','Discover','WF','Chase','Savor1','DeltaAmexGold'],date:[today],method:['WF'],cost:[1],last:['','','',1,2,3,4,5]},
 RetireTable:{cols:['Date','Venue','Method','Cost','403b','Roth 403b','Other'],date:[today+3,today-10],method:['403b','Roth 403b'],cost:[-800,-100],last:['','','','',20800,5100,77]},
};
let travel={status:200,text:[['Delta Miles','United Miles','Delta Credit','United Credit'],['45,210','12,800','$200.00','$0.00']],values:[['Delta Miles','United Miles','Delta Credit','United Credit'],[45210,12800,200,0]]};
let retireFail=false;
global.fetch=async(url)=>{ const u=decodeURIComponent(url); const r=(s,b)=>({ok:s<300,status:s,json:async()=>b});
  if(u.includes('tables/TravelTable/range?$select=text,values')) return r(travel.status,{text:travel.text,values:travel.values});
  const tn=(u.match(/tables\/([^/]+)\//)||[])[1]; const tb=T[tn];
  if(tn==='RetireTable' && retireFail) return r(404,{error:{message:'x'}});
  if(/\/columns(\?\$select=name)?$/.test(u)) return r(200,{value:tb.cols.map(name=>({name}))});
  if(u.endsWith('/range/lastRow')) return r(200,{values:[tb.last]});
  const col=(u.match(/columns\/([^/]+)\/dataBodyRange/)||[])[1];
  if(col){ const key={Date:'date',Method:'method',Cost:'cost'}[col]; return r(200,{values:tb[key].map(v=>[v])}); }
  throw new Error('unexpected '+u); };
ok(tableFor('Travel')==='TravelTable' && tableFor('Retire')==='RetireTable','Travel/Retire default table names');
await loadBalances(); setBalMode('today'); let h=g('balance-expand').innerHTML;
const iR=h.indexOf('>Retire<'), iT=h.indexOf('>Travel<'), iS=h.indexOf('>Savings<');
ok(iS>=0 && iS<iR && iR<iT,'order: other accounts, Retire, Travel last');
ok(h.includes('>403b</div><div class="bal-value positive">$20,000.00') && h.includes('>Roth 403b</div><div class="bal-value positive">$5,100.00') && !h.includes('>Other<'),'Today: 403b leaves out the future-dated 800; only 403b + Roth 403b');
ok(h.includes('>Delta Miles</div><div class="bal-value">45,210') && h.includes('$200.00') && !h.includes('United Credit'),'Travel shows Excel text, 0 hidden');
ok(g('bal-note').textContent.includes('leaving out 1 future-dated row'),'note counts Retire future row too');
setBalMode('all'); h=g('balance-expand').innerHTML; ok(h.includes('$20,800.00') && h.includes('45,210'),'All: raw 403b, Travel same');
setBalMode('month'); ok(g('balance-expand').style.display==='none','Month-End: cards only'); setBalMode('today');
travel.text=[['Delta Miles'],['<b>x</b>']]; travel.values=[['Delta Miles'],['<b>x</b>']]; await loadBalances(); ok(g('balance-expand').innerHTML.includes('&lt;b&gt;x&lt;/b&gt;'),'Travel text escaped');
travel.status=404; await loadBalances(); h=g('balance-expand').innerHTML; ok(!h.includes('>Travel<') && h.includes('>Retire<') && g('balance-main').innerHTML.includes('Total Credit'),'no TravelTable → just no Travel section');
travel.status=200; travel.text=[['Delta Miles']]; travel.values=[['Delta Miles']]; await loadBalances(); ok(!g('balance-expand').innerHTML.includes('>Travel<'),'headers only → nothing');
retireFail=true; columnCache={}; await loadBalances(); h=g('balance-expand').innerHTML; ok(!h.includes('>Retire<') && g('balance-main').innerHTML.includes('Total Credit'),'RetireTable missing → rest still loads');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'More accounts': async () => {
showToast=()=>{}; const g=id=>document.getElementById(id);
balanceCache={Cash:{Date:'',Method:'',Checking:1,Cash:2,Venmo:3,'COne Savings':10,'HSA Reimb':'',Optum:500,'Cost (MYR)':'',MYR:'x'},Credit:{Discover:1,WF:2,Chase:3,Savor1:4,DeltaAmexGold:5}};
futureCache={Cash:{sums:{},count:0},Credit:{sums:{},count:0}}; travelCache=[];
setBalMode('today'); let h=g('balance-expand').innerHTML;
ok(h.includes('>HSA Reimb</div><div class="bal-value zero">—') && h.indexOf('HSA Reimb')<h.indexOf('>Optum<') && h.indexOf('COne Savings')<h.indexOf('HSA Reimb'),'blank HSA Reimb still shows (—), between COne Savings and Optum');
ok(!h.includes('Cost (MYR)') && !h.includes('>MYR<'),'non-account blanks/text still hidden');
balanceCache.Cash['HSA Reimb']=45; setBalMode('today'); ok(g('balance-expand').innerHTML.includes('>HSA Reimb</div><div class="bal-value positive">$45.00'),'with a value: $45.00');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Sync safety + offline': async () => {
let toasts=[]; showToast=m=>toasts.push(m); global.setTimeout=fn=>fn(); let confirms=0; confirm=()=>{confirms++;return true;}; renderQueueReview=()=>{};
const g=id=>document.getElementById(id);
store.cf_file_path='a.xlsx'; store.cf_table_cash='CashTable'; store.cf_table_credit='CreditTable';
let redirects=0, silentFail=false;
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>{ if(silentFail) throw new Error('x'); return {accessToken:'tok'}; },acquireTokenRedirect:()=>{redirects++;},loginRedirect:()=>{redirects++;}}; accessToken='tok';
const COLS={CashTable:['Date','Method','Cost','Checking'],CreditTable:['Date','Method','Cost','Discover','WF','Chase','Savor1','DeltaAmexGold'],RetireTable:['Date','Method','Cost','403b']};
let calls=[], rowsFail=false;
global.fetch=async(url,opt={})=>{ const u=decodeURIComponent(url); calls.push(u); const r=(s,b)=>({ok:s<300,status:s,json:async()=>b}); const tn=(u.match(/tables\/([^/]+)\//)||[])[1];
  if(u.includes('TravelTable')) return r(404,{});
  if(/\/columns(\?\$select=name)?$/.test(u)) return r(200,{value:COLS[tn].map(name=>({name}))});
  if(u.endsWith('/range/lastRow')) return rowsFail ? r(500,{}) : r(200,{values:[COLS[tn].map((c,i)=>i<3?'':100)]});
  if(u.includes('dataBodyRange')) return r(200,{values:[['']]});
  if(u.endsWith('/rows/add')) return r(201,{});
  if(u.endsWith('/createSession')) return r(201,{id:'S'}); if(u.endsWith('/closeSession')) return r(204,{});
  throw new Error('unexpected '+u); };
const row=(sheet,c)=>({sheet,payload:{date:'2026-11-16',venue:'V',method:'Checking',cost:c,inputted:'now'}});
// 1. ⚠ confirm guard
queue=[{...row('Cash',1),uncertain:true},{...row('Cash',2),uncertain:true}]; isSyncing=true; confirmQueueItem(0);
ok(queue[0].uncertain && confirms===0 && toasts.at(-1)==='Wait for sync to finish','⚠ "send" blocked mid-sync (list may be out of step)'); isSyncing=false;
confirmQueueItem(5); ok(confirms===0,'stale index → nothing happens');
confirmQueueItem(0); ok(!queue[0].uncertain && queue[1].uncertain && confirms===1,'normal: un-flags exactly that row');
queue=[row('Cash',1)]; confirmQueueItem(0); ok(confirms===1,'not a ⚠ row → no prompt');
// 2. offline: no sign-in redirect
navigator.onLine=false; silentFail=true; toasts=[]; const tk=await getToken();
ok(tk===null && redirects===0 && toasts.at(-1).startsWith('Offline'),'offline + expired sign-in → no redirect, says offline');
balanceCache=null; calls=[]; await loadBalances(); ok(calls.length===0 && g('balance-main').innerHTML.includes('Offline'),'Balances offline → message, no requests');
balanceCache={Cash:{Checking:5},Credit:{}}; g('balance-main').innerHTML='KEEP'; toasts=[]; await loadBalances(); ok(g('balance-main').innerHTML==='KEEP' && toasts.at(-1).includes('last balances'),'offline with earlier balances → keeps them');
navigator.onLine=true; silentFail=false;
// 2b. going offline mid-sync → remaining rows are plain failures (retryable), not ⚠
queue=[row('Cash',1)]; saveQueue(); columnCache={}; const orig=getTableColumns; 
const p=syncQueue(); navigator.onLine=false; await p; navigator.onLine=true;
ok(queue.length===1 && !queue[0].uncertain && !queue[0].sending,'offline before sending → kept, retryable, not ⚠');
// 3. refresh only when Balances is showing
queue=[row('Cash',1)]; saveQueue(); g('balance-section').style.display='none'; balanceCache={Cash:{},Credit:{}}; balanceTs=Date.now(); calls=[];
await syncQueue(); await new Promise(r=>setImmediate(r));
ok(!calls.some(u=>u.includes('lastRow')) && balanceTs===0 && queue.length===0,'after sync, Balances hidden → no reload, marked stale');
queue=[row('Cash',1)]; saveQueue(); g('balance-section').style.display=''; calls=[]; await syncQueue(); await new Promise(r=>setTimeout(r,0)); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r));
ok(calls.some(u=>u.includes('lastRow')),'after sync, Balances showing → reloads');
g('balance-section').style.display='none'; queueMode=false; calls=[]; balanceTs=5; await commitRows([row('Cash',1)],'x'); ok(!calls.some(u=>u.includes('lastRow')) && balanceTs===0,'direct add, Balances hidden → no reload'); queueMode=true;
// stale → Pay autofill refetches
calls=[]; await getPayAmount('Discover','month'); ok(calls.some(u=>u.includes('CreditTable/range/lastRow')),'stale balances → Pay favorite refetches');
// 4. nothing loads → Failed to load
rowsFail=true; columnCache={}; balanceCache=null; await loadBalances(); ok(g('balance-main').innerHTML.includes('Failed to load'),'no data back → "Failed to load", not dashes'); rowsFail=false;
await loadBalances(); ok(g('balance-main').innerHTML.includes('Total Credit'),'loads normally again');
  },

  // ─────────────────────────────────────────────────────────────────────────────
  'Analysis tab': async () => {
const RD=Date; global.Date=class extends RD{constructor(...a){ if(a.length) super(...a); else super(2026,10,16);} static now(){return new RD(2026,10,16).getTime();}};
let toasts=[]; showToast=m=>toasts.push(m); global.setTimeout=fn=>fn(); const g=id=>document.getElementById(id);
store.cf_file_path='a.xlsx';
msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; accessToken='tok';
const H=['Year','Month','Date','Pay','Credit','Health','HSA','Savings','Rent HM','Move HM','Unassigned HM','Home','Wifi UT','Gas UT','Utilities','Groceries FD','Out FD','Caffeine FD','Unassigned FD','F&D','Dr','Need','Hair','Eyes','Gas TR','Parking TR','Transport','Want','Travel','Fam','Total Spending','Total Savings'];
const val=(over={})=>H.map((h,i)=>over[h]!==undefined?over[h]:i*10);
let NOW=val({Year:2026,Month:11,Date:46327,'F&D':412.5,'Groceries FD':300,'Gas TR':-45.1,Pay:4321.09});
let AVG=[['Average',12,46000,...H.slice(3).map((h,i)=>100+i)],['2-YR',24,45600,...H.slice(3).map((h,i)=>200+i)]];
let mode={}, calls=[], avgCol=1;
global.fetch=async(url)=>{ const u=decodeURIComponent(url); calls.push(u); const r=(s,b)=>({ok:s<300,status:s,json:async()=>b});
  if(u.includes('tables/AnalysisTable/headerRowRange')) return mode.table===404?r(404,{}):r(200,{values:[H],columnIndex:1});
  if(u.includes('tables/AnalysisTable/range/lastRow')) return mode.table===404?r(404,{}):r(200,{values:[NOW],columnIndex:1});
  if(u.includes('tables/AnalysisAvg/range')) return (mode.avg===404||mode.avg==='name')?r(404,{}):r(200,{values:[H,...AVG],columnIndex:avgCol});
  if(u.includes('names/AnalysisAvg/range')) return mode.avg==='name'?r(200,{values:AVG,columnIndex:avgCol}):r(404,{});
  throw new Error('unexpected '+u); };
ok(tableFor('Analysis')==='AnalysisTable','default table name');
switchSheet('Analysis'); await new Promise(r=>setImmediate(r)); await new Promise(r=>setImmediate(r)); await loadAnalysis();
ok(g('analysis-section').style.display==='' && g('form-wrapper').style.display==='none' && g('balance-section').style.display==='none','Analysis tab shows its section only');
ok(calls.every(u=>u.includes('?$select=values,columnIndex')) && !calls.some(u=>/rows\/add|PATCH/.test(u)),'read-only requests');
calls=[]; await loadAnalysis(); ok(calls.length===3 && !calls.some(u=>u.includes('names/')),'AnalysisAvg as a table: exactly 3 requests, no fallback');
let h=g('analysis-main').innerHTML; const rows=h.split('<tr').slice(2).map(r=>r.replace(/<[^>]+>/g,'|'));
const names=h.split('<tr').slice(2).map(r=>(r.match(/<td>([^<]*)<\/td>/)||[])[1]);
ok(names.join()==='F&amp;D,Groceries,Out,Caffeine,Transport,Gas,Want,Dr,Travel,Need,Hair,Eyes,Home,Utilities,Pay,Savings','base rows in your order, sub codes dropped, no Unassigned');
ok(h.includes('<th>Nov 2026</th><th>12-mo avg</th><th>24-mo avg</th>'),'headings: month from the last row, windows from Excel');
ok(/class="an-sub"><td>Groceries<\/td><td class="an-now">\$300\.00<\/td><td>\$112\.00<\/td><td>\$212\.00<\/td>/.test(h),'sub row: this month + both averages, lined up by column');
ok(h.includes('<td>F&amp;D</td><td class="an-now">$412.50</td><td>$116.00</td><td>$216.00</td>') && h.includes('<td>Gas</td><td class="an-now">−$45.10</td>') && h.includes('$4,321.09'),'values and negatives formatted');
AVG[0][1]=6; await loadAnalysis(); ok(g('analysis-main').innerHTML.includes('<th>6-mo avg</th>'),'change the 12 in Excel → heading follows'); AVG[0][1]=12;
// Show all
toggleAnalysisAll(); h=g('analysis-main').innerHTML; const all=h.split('<tr').slice(2).map(r=>(r.match(/<td>([^<]*)<\/td>/)||[])[1]);
ok(all.join()==='Pay,Credit,Health,HSA,Savings,Home,Rent,Move,Unassigned,Utilities,Wifi,Gas,F&amp;D,Groceries,Out,Caffeine,Unassigned,Dr,Need,Hair,Eyes,Transport,Gas,Parking,Want,Travel,Fam,Total Spending,Total Savings','Show all: every column, total first then its subs, no Year/Month/Date');
ok(g('an-all-toggle').innerHTML.includes('Show less'),'toggle label'); toggleAnalysisAll(); ok(g('analysis-main').innerHTML.split('<tr').length-2===16,'back to base');
// named rows starting in a different column still line up
mode.avg='name'; avgCol=4; const saved=AVG; AVG=AVG.map(r=>r.slice(3)); await loadAnalysis(); h=g('analysis-main').innerHTML;
ok(h.includes('<td>Groceries</td><td class="an-now">$300.00</td><td>$112.00</td>') && h.includes('<th>Average</th>'),'named range starting at another column → still aligned'); AVG=saved; avgCol=1; mode.avg=0;
// AnalysisAvg as its own table (header row + the two rows), or a name that includes a header row
// the averages table with its columns in a different order and its own first two headers still lines up, by name
const realH=H, realAVG=AVG; const perm=[0,1,2,...H.slice(3).map((x,i)=>i+3).reverse()]; const savedFetch=global.fetch;
global.fetch=async(url)=>{ const u=decodeURIComponent(url); if(u.includes('tables/AnalysisAvg/range')) return {ok:true,status:200,json:async()=>({values:[perm.map((p,i)=>i===0?'Name':i===1?'Months':realH[p]),...realAVG.map(r=>perm.map(p=>r[p]))],columnIndex:7})}; return savedFetch(url); };
await loadAnalysis(); h=g('analysis-main').innerHTML;
ok(h.includes('<th>Nov 2026</th><th>12-mo avg</th><th>24-mo avg</th>') && h.includes('<td>Groceries</td><td class="an-now">$300.00</td><td>$112.00</td><td>$212.00</td>') && h.includes('<td>F&amp;D</td><td class="an-now">$412.50</td><td>$116.00</td><td>$216.00</td>'),'averages table matched by header name (reordered columns, its own Name/Months headers, anywhere on the sheet)');
global.fetch=savedFetch;
mode.avg='name'; await loadAnalysis(); h=g('analysis-main').innerHTML;
ok(h.includes('<th>Nov 2026</th><th>12-mo avg</th><th>24-mo avg</th>') && h.includes('<td>Groceries</td><td class="an-now">$300.00</td><td>$112.00</td><td>$212.00</td>'),'no table but a named range → fallback gives the same result'); mode.avg=0;
// no named range
mode.avg=404; await loadAnalysis(); h=g('analysis-main').innerHTML; ok(h.includes('<td>Pay</td><td class="an-now">$4,321.09</td></tr>') && g('an-note').textContent.includes('AnalysisAvg'),'no named rows → this month only + hint'); mode.avg=0;
mode.avg=0; await loadAnalysis();
// which rows show comes from the analysis.rows setting
const keep=config.settings; const shownNames=()=>g('analysis-main').innerHTML.split('<tr').slice(2).map(r=>(r.match(/<td>([^<]*)<\/td>/)||[])[1]).join();
config.settings=parseSettings([['Key','From','Value'],['Analysis.Rows','',' want ,No Such Column, gas tr,PAY']]); renderAnalysis();
ok(shownNames()==='Want,Gas,Pay' && g('analysis-main').innerHTML.includes('class="an-sub"><td>Gas<'),'analysis.rows: your order, any letter case, unknown names skipped, coded column shown as a sub');
config.settings=parseSettings([['Key','From','Value']]); renderAnalysis();
ok(shownNames()==='Pay,Credit,Health,HSA,Savings,Home,Utilities,F&amp;D,Dr,Need,Hair,Eyes,Transport,Want,Travel,Fam,Total Spending,Total Savings','no analysis.rows → every category total, no subs');
config.settings=keep; renderAnalysis(); ok(shownNames().startsWith('F&amp;D,Groceries,Out,Caffeine,Transport,Gas,Want'),'setting back → base list back');
ok(!g('analysis-main').innerHTML.includes('an-new'),'last row is this month → no "Start" button');
// failures
analysisCache=null; mode.table=404; await loadAnalysis(); ok(g('analysis-main').innerHTML.includes('Couldn’t read AnalysisTable') && g('an-all-toggle').style.display==='none','table missing → message'); mode.table=0;
navigator.onLine=false; analysisCache=null; calls=[]; await loadAnalysis(); ok(calls.length===0 && g('analysis-main').innerHTML.includes('Offline'),'offline → message, no requests'); navigator.onLine=true;
msalInstance={getAllAccounts:()=>[]}; await loadAnalysis(); ok(g('analysis-main').innerHTML.includes('Sign in'),'signed out → message'); msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})};
// tabs: form state survives a visit to Analysis
switchSheet('Credit'); const pi=favorites.findIndex(x=>x.label==='Grocer'); activeFav=pi; g('f-venue').value='Corner Grocer'; switchSheet('Analysis'); switchSheet('Credit');
ok(activeFav===pi && g('f-venue').value==='Corner Grocer' && g('analysis-section').style.display==='none','entry form untouched by visiting Analysis');
switchSheet('Balances'); ok(g('analysis-section').style.display==='none' && g('balance-section').style.display==='','Balances hides Analysis');
  },

  'Analysis: start new month': async () => {
let toasts=[]; showToast=m=>toasts.push(m); global.setTimeout=fn=>fn(); const g=id=>document.getElementById(id);
store.cf_file_path='a.xlsx'; msalInstance={getAllAccounts:()=>[{}],acquireTokenSilent:async()=>({accessToken:'tok'})}; accessToken='tok';
const RD=Date; global.Date=class extends RD{constructor(...a){ if(a.length) super(...a); else super(2026,10,16);} static now(){return new RD(2026,10,16).getTime();}};
ok(colName(0)==='A' && colName(25)==='Z' && colName(26)==='AA' && colName(73)==='BV' && colName(701)==='ZZ' && colName(702)==='AAA','column letters');
ok(nextMonth('2026-12')==='2027-01' && nextMonth('2026-01')==='2026-02' && monthName('2026-11')==='Nov 2026','month helpers');
// A made-up Analysis sheet. Columns A…I; the table is B…F (header on sheet row 2); H and I are check columns outside it.
const HEAD=['Year','Month','Date','Pay','Total'];
const month=(y,m)=>['',y,m,'=DATE(RC[-2],RC[-1],1)','=SUMIFS(Cost,Yr,RC[-3],Mo,RC[-2])','=RC[-1]*2','','=SUM(RC[-3],RC[-2])',7];
const blank=()=>['','','','','','','','',''];
let grid, sheetName, fail={}, calls=[], seen={}, answer=true; global.confirm=m=>{ seen.confirm=m; return answer; };
// months = finished months, oldest first; then the green line; then the current row labelled `last`
const build=(finished,last,green=true)=>{ grid=[['note','','','','','','','',''],['',...HEAD,'','chk','chk2'],...finished.map(k=>month(...k)),...(green?[blank()]:[]),month(...last),blank(),['','Name','Months','','','','','','']]; };
const tableEnd=()=>grid.findIndex((r,i)=>i>1 && r[1]==='Name')-2;          // index of the table's last row
const col=s=>s.split('').reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;
const valueOf=c=>(typeof c==='string'&&c.startsWith('='))?1:c;
global.fetch=async(url,opt={})=>{ const u=decodeURIComponent(url), method=opt.method||'GET'; calls.push({u,method,session:(opt.headers||{})['workbook-session-id']});
  const r=(s,b)=>({ok:s<300,status:s,json:async()=>b}); const body=opt.body?JSON.parse(opt.body):null;
  if(u.endsWith('/createSession')) return r(201,{id:'S1'});
  if(u.endsWith('/closeSession')) return r(204,null);
  if(u.includes('tables/AnalysisTable/headerRowRange')) return r(200,{values:[HEAD],columnIndex:1});
  if(u.includes('tables/AnalysisTable/range/lastRow')) return r(200,{values:[grid[tableEnd()].slice(1,6).map(valueOf)],columnIndex:1});
  if(u.includes('AnalysisAvg')) return r(404,{});
  if(u.includes('tables/AnalysisTable/range?')) { if(fail.table) return r(500,{error:{message:'busy'}});
    return r(200,{address:`${sheetName.includes(' ')?`'${sheetName}'`:sheetName}!B2:F${tableEnd()+1}`,rowIndex:1,columnIndex:1,values:grid.slice(1,tableEnd()+1).map(row=>row.slice(1,6).map(valueOf))}); }
  if(!u.includes(`worksheets/${sheetName}/`)) throw new Error('wrong sheet '+u);
  if(u.includes('/usedRange')) return r(200,{columnIndex:0,columnCount:9});
  const m=u.match(/range\(address='([A-Z]*)(\d+):([A-Z]*)(\d+)'\)(\/insert)?/); if(!m) throw new Error('unexpected '+u);
  const row=+m[2]-1;
  if(m[5]) { seen.busy=busySending(); if(fail.insert) return r(500,{error:{message:'no'}}); if(m[1]||m[3]||body.shift!=='Down') throw new Error('not a whole row'); grid.splice(row,0,blank()); return r(200,{}); }
  if(method==='GET') return r(200,{formulasR1C1:[grid[row].slice(col(m[1]),col(m[3])+1)]});
  if(method==='PATCH') { if(fail.fill && body.formulasR1C1) return r(500,{error:{message:'no'}});
    const cells=(body.formulasR1C1||body.values)[0]; if(cells.length!==col(m[3])-col(m[1])+1) throw new Error('size mismatch');
    cells.forEach((c,k)=>{ if(c!==null) grid[row][col(m[1])+k]=c; }); return r(200,{}); }
  throw new Error('unexpected '+method+' '+u); };
const writes=()=>calls.filter(c=>c.method!=='GET' && !/Session$/.test(c.u));
const start=async(finished,last,green)=>{ build(finished,last,green); analysisCache=null; await loadAnalysis(); calls=[]; toasts=[]; seen={}; };

// 1. One month behind (it's Nov 2026, the last row still says Oct)
sheetName='Analysis'; await start([[2026,8],[2026,9]],[2026,10]);
let h=g('analysis-main').innerHTML;
ok(analysisCache.lastKey==='2026-10' && h.includes('class="an-new"') && h.includes('still Oct 2026') && h.includes('startNewMonth()') && h.includes('Start Nov 2026'),'behind → banner with a "Start Nov 2026" button');
answer=false; await startNewMonth(); ok(calls.length===0 && seen.confirm.includes('a row above the green line for Oct 2026') && seen.confirm.includes('last row to Nov 2026'),'asks first, spelling out the change; Cancel → no requests');
answer=true; const before=JSON.stringify(grid[3]); await startNewMonth();
// sheet rows now: 3 Aug, 4 Sep, 5 NEW Oct, 6 green, 7 current (Nov)
ok(grid.length===9 && JSON.stringify(grid[2].slice(1,3))==='[2026,8]' && JSON.stringify(grid[3])===before,'older rows untouched');
ok(JSON.stringify(grid[4])===JSON.stringify(['',2026,10,'=DATE(RC[-2],RC[-1],1)','=SUMIFS(Cost,Yr,RC[-3],Mo,RC[-2])','=RC[-1]*2','','=SUM(RC[-3],RC[-2])','']),'new row: the ended month + every formula of the row above (check column too); typed values not duplicated');
ok(grid[5].every(c=>c==='') ,'green line still blank, now below the new row');
ok(JSON.stringify(grid[6])===JSON.stringify(month(2026,11)),'last row relabelled to the new month, formulas and other cells intact');
ok(grid[8][1]==='Name','everything below moved down one row');
const w=writes(); ok(w.length===3 && w[0].u.includes("range(address='5:5')/insert") && w[1].u.includes("range(address='A5:I5')") && w[2].u.includes("range(address='B7:F7')"),'exactly three changes: insert whole row 5, fill A5:I5, relabel B7:F7');
const iRead=calls.findIndex(c=>c.u.includes("range(address='A4:I4')")), iIns=calls.findIndex(c=>c.u.includes('/insert'));
ok(iRead!==-1 && iRead<iIns,'formulas are read from the row above before anything changes');
ok(calls.filter(c=>!/createSession$/.test(c.u) && c.u.includes('worksheets/')).every(c=>c.session==='S1') && calls.some(c=>c.u.endsWith('/closeSession')),'one workbook session, closed afterwards');
ok(seen.busy===true && busySending()===false,'sync and saving wait while the sheet changes, free again after');
ok(toasts.some(t=>t.includes('Nov 2026 started')) && analysisCache.lastKey==='2026-11' && !g('analysis-main').innerHTML.includes('an-new'),'success → view reloaded, button gone');
calls=[]; await startNewMonth(); ok(calls.length===0,'already on this month → does nothing');

// 2. Three months behind → catches up month by month
await start([[2026,6],[2026,7]],[2026,8]); await startNewMonth();
ok(seen.confirm.includes('3 rows') && seen.confirm.includes('Aug 2026, Sep 2026, Oct 2026'),'confirm lists every month');
ok(grid.slice(2,9).map(r=>r[1]===''?'—':r[1]+'-'+r[2]).join()==='2026-6,2026-7,2026-8,2026-9,2026-10,—,2026-11' && writes().length===9,'three rows added in order, green line kept, last row = Nov');
ok(grid.slice(4,7).every(r=>r[5]==='=RC[-1]*2' && r[8]===''),'each added row carries the formulas');

// 3. Sheet name with a space; year boundary
sheetName='My Analysis'; global.Date=class extends RD{constructor(...a){ if(a.length) super(...a); else super(2027,0,3);} static now(){return new RD(2027,0,3).getTime();}};
await start([[2026,10],[2026,11]],[2026,12]); await startNewMonth();
ok(JSON.stringify(grid[4].slice(1,3))==='[2026,12]' && JSON.stringify(grid[6].slice(1,3))==='[2027,1]','December → January of the next year; quoted sheet name handled');
sheetName='Analysis'; global.Date=class extends RD{constructor(...a){ if(a.length) super(...a); else super(2026,10,16);} static now(){return new RD(2026,10,16).getTime();}};

// 4. Layout not as expected → stops before changing anything
const untouched=async(label,finished,last,green)=>{ await start(finished,last,green); const snap=JSON.stringify(grid); await startNewMonth();
  ok(writes().length===0 && JSON.stringify(grid)===snap && toasts.some(t=>t.startsWith('Nothing was changed') && t.includes('laid out')) && busySending()===false,label); };
await untouched('no green line → nothing changed',[[2026,8],[2026,9]],[2026,10],false);
await untouched('a month missing before the green line → nothing changed',[[2026,7],[2026,8]],[2026,10]);
await start([[2026,8],[2026,9]],[2026,10]); const H0=HEAD[0]; HEAD[0]='Yr'; grid[1][1]='Yr'; let snap=JSON.stringify(grid); analysisCache={...analysisCache,lastKey:'2026-10'}; await startNewMonth();
ok(writes().length===0 && JSON.stringify(grid)===snap,'no Year column → nothing changed'); HEAD[0]=H0;

// 5. Guards
await start([[2026,8],[2026,9]],[2026,10]);
isSyncing=true; await startNewMonth(); ok(calls.length===0 && toasts.pop().includes('Wait for sync'),'mid-sync → blocked'); isSyncing=false;
navigator.onLine=false; await startNewMonth(); ok(calls.length===0 && toasts.pop().includes('Offline'),'offline → blocked'); navigator.onLine=true;
analysisCache.lastKey='2025-10'; seen={}; await startNewMonth(); ok(calls.length===0 && !seen.confirm && toasts.pop().includes('More than a year'),'more than 12 months behind → left to Excel');
analysisCache.lastKey=null; await startNewMonth(); ok(calls.length===0,'month unknown → does nothing');

// 6. Failures part-way
await start([[2026,8],[2026,9]],[2026,10]); fail.table=true; snap=JSON.stringify(grid); await startNewMonth();
ok(JSON.stringify(grid)===snap && toasts.some(t=>t.startsWith('Nothing was changed: busy')),'Excel refuses the first read → nothing changed, reason shown'); fail={};
await start([[2026,8],[2026,9]],[2026,10]); fail.insert=true; snap=JSON.stringify(grid); await startNewMonth();
ok(JSON.stringify(grid)===snap && writes().length===1 && toasts.some(t=>t.startsWith('Nothing was changed')),'insert refused → nothing changed, no fill attempted'); fail={};
await start([[2026,8],[2026,9]],[2026,10]); fail.fill=true; await startNewMonth();
ok(grid[4].every(c=>c==='') && JSON.stringify(grid[6].slice(1,3))==='[2026,10]' && toasts.some(t=>t.includes('blank row was added')) && busySending()===false && calls.some(c=>c.u.endsWith('/closeSession')),'fill refused after the insert → says a blank row was added and how to fix it; last row not relabelled'); fail={};

// 7. While the sheet is changing, entries queue and Sync waits
changingSheet=true; toasts=[]; await syncQueue(); ok(toasts.some(t=>t.includes('new month')),'Sync waits'); 
ok(busySending()===true,'counts as busy'); changingSheet=false;
  },

};

(async () => {
  let passed = 0, failed = 0;
  for (const [name, suite] of Object.entries(SUITES)) {
    const ctx = vm.createContext({ console, setTimeout, clearTimeout, setImmediate });
    let p = 0, f = 0, crash = null;
    try {
      // page → the app's script → "the phone already has a setup" → the suite
      await vm.runInContext(`${PAGE}\n${APP}\n;useFixture();(${suite.toString()})()`, ctx, { filename: name });
    } catch (e) { crash = e; }
    try { [p, f] = vm.runInContext('[P, F]', ctx); } catch (e) { /* page never loaded */ }
    if (crash) { f++; console.log(`✗ ${name}: stopped early — ${crash && crash.stack ? crash.stack.split('\n').slice(0, 3).join(' | ') : crash}`); }
    console.log(`${f ? '✗' : '✓'} ${name}: ${p} passed${f ? `, ${f} failed` : ''}`);
    passed += p; failed += f;
  }
  console.log(`\n${failed ? `${failed} FAILED, ` : ''}${passed} passed — ${failed ? 'NOT OK' : 'ALL PASSED'} (${file})`);
  process.exit(failed ? 1 : 0);
})();
