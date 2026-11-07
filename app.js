'use strict';
const $=id=>document.getElementById(id);
const sources={split:{file:'file',select:'splitSheet',status:'status'},orders:{file:'orderFile',select:'orderSheet',status:'inventoryStatus'},inventory:{file:'inventoryFile',select:'inventorySheet',status:'inventoryStatus'}};
let output=null,inventoryOutput=null;
function status(id,message,error=false){$(id).textContent=message;$(id).classList.toggle('error',error);}
function invalidate(key){
 if(key==='split'){output=null;$('result').hidden=true;}
 else{inventoryOutput=null;$('downloadInventory').disabled=true;$('calculateInventory').disabled=!(sources.orders.book&&sources.inventory.book);}
}
function validateFile(file){if(!/\.xlsx?$/i.test(file.name))throw Error('請選擇 .xlsx 或 .xls Excel 檔案。');}
async function loadSource(key,file){
 const source=sources[key],version=source.revision=(source.revision??0)+1;
 source.book=null;$(source.select).replaceChildren();$(source.select).disabled=true;invalidate(key);
 if(!file){status(source.status,key==='split'?'選擇檔案後，會自動計算分類結果。':'請選擇兩份 Excel，再按「計算可售庫存」。');return;}
 status(source.status,'正在讀取 '+file.name+'…');
 try{
  validateFile(file);
  const bytes=await file.arrayBuffer();if(version!==source.revision)return;
  const book=OrderTools.readWorkbook(bytes);source.book=book;source.name=file.name;
  for(const name of book.SheetNames)$(source.select).add(new Option(name,name));
  $(source.select).disabled=false;
  if(key==='split')calculateSplit();
  else{invalidate(key);status(source.status,'已讀取 '+file.name+'。請確認兩份檔案與工作表，再按「計算可售庫存」。');}
 }catch(error){if(version!==source.revision)return;source.book=null;invalidate(key);status(source.status,'無法處理：'+error.message,true);}
}
function calculateSplit(){
 invalidate('split');const source=sources.split;if(!source.book)return;
 try{
  const data=OrderTools.splitWorkbook(source.book,$('splitSheet').value);output=data;
  for(const id of ['preOrders','stockOrders'])$(id).textContent=data[id].toLocaleString()+' 張訂單';
  for(const id of ['preRows','stockRows'])$(id).textContent=data[id].toLocaleString()+' 筆商品資料';
  status('status',source.name+' · 工作表：'+$('splitSheet').value+' · 共 '+data.total.toLocaleString()+' 筆商品資料'+(data.preRows===0?'\n「選項」欄位沒有找到預購字樣，所有資料均歸入現貨單。':''));
  $('result').hidden=false;
 }catch(error){status('status','無法處理：'+error.message,true);}
}
function download(book,name){
 try{
  const bytes=OrderTools.writeWorkbook(book),url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})),a=document.createElement('a');
  a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
 }catch(error){status(name==='可售庫存.xlsx'?'inventoryStatus':'status','無法下載：'+error.message,true);}
}
for(const [key,source] of Object.entries(sources)){
 $(source.file).addEventListener('change',e=>loadSource(key,e.target.files[0]));
 $(source.select).addEventListener('change',()=>{
  if(key==='split')calculateSplit();
  else{invalidate(key);status('inventoryStatus','工作表已變更，請重新計算可售庫存。');}
 });
}
for(const name of ['dragenter','dragover'])$('drop').addEventListener(name,e=>{e.preventDefault();$('drop').classList.add('over');});
$('drop').addEventListener('dragleave',()=>$('drop').classList.remove('over'));
$('drop').addEventListener('drop',e=>{e.preventDefault();$('drop').classList.remove('over');if(e.dataTransfer.files.length)loadSource('split',e.dataTransfer.files[0]);});
$('pre').onclick=()=>{if(output)download(output.pre,'預購單.xlsx');};
$('stock').onclick=()=>{if(output)download(output.stock,'現貨單.xlsx');};
$('both').onclick=()=>{if(output){download(output.pre,'預購單.xlsx');download(output.stock,'現貨單.xlsx');}};
$('calculateInventory').onclick=()=>{
 inventoryOutput=null;$('downloadInventory').disabled=true;
 try{
  if(!sources.orders.book||!sources.inventory.book)return;
  const data=OrderTools.inventoryWorkbook(sources.orders.book,$('orderSheet').value,sources.inventory.book,$('inventorySheet').value);inventoryOutput=data.book;
  status('inventoryStatus','完成：訂單明細 '+data.orderRows.toLocaleString()+' 筆、庫存明細 '+data.inventoryRows.toLocaleString()+' 筆。\n共 '+data.count.toLocaleString()+' 個貨號；庫存明細缺少 '+data.missing.toLocaleString()+' 個（已加備註），負庫存 '+data.negative.toLocaleString()+' 個。');
  $('downloadInventory').disabled=false;
 }catch(error){status('inventoryStatus','無法處理：'+error.message,true);}
};
$('downloadInventory').onclick=()=>{if(inventoryOutput)download(inventoryOutput,'可售庫存.xlsx');};
if(!globalThis.XLSX){
 for(const source of Object.values(sources))$(source.file).disabled=true;
 status('status','Excel 元件載入失敗，請確認 vendor 資料夾與網頁放在一起。',true);
 status('inventoryStatus','Excel 元件載入失敗，請重新載入網頁。',true);
}
