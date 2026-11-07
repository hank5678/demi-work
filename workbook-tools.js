(function(root){
'use strict';
const X=typeof module==='object'&&module.exports?require('./vendor/xlsx.full.min.js'):root.XLSX;
function readWorkbook(bytes){
 const signature=new Uint8Array(bytes);
 const zip=signature[0]===0x50&&signature[1]===0x4b;
 const cfb=[0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1].every((v,i)=>signature[i]===v);
 const biff=signature[0]===0x09&&[0x00,0x02,0x04,0x08].includes(signature[1]);
 if(!zip&&!cfb&&!biff)throw Error('請選擇有效的 .xlsx 或 .xls Excel 檔案。');
 let book;
 try{book=X.read(bytes,{type:'array',cellNF:true,bookFiles:true});}
 catch{throw Error('無法讀取 Excel，請確認檔案完整且未加密。');}
 const valid=zip?book.keys?.includes('xl/workbook.xml'):cfb?book.cfb?.FullPaths.some(path=>/\/(Workbook|Book)$/.test(path)):biff;
 if(!valid||!book.SheetNames.length)throw Error('請選擇有效的 .xlsx 或 .xls Excel 檔案。');
 delete book.files;delete book.keys;delete book.cfb;
 return book;
}
function text(cell){
 if(!cell)return '';
 if(cell.t==='e')throw Error('儲存格含有 Excel 錯誤，請修正後重試。');
 if(cell.f&&cell.v==null)throw Error('公式缺少已儲存的計算結果，請用 Excel 重新計算並儲存後重試。');
 if(cell.t==='z')return '';
 return cell.t==='n'?X.utils.format_cell(cell):String(cell.v??'');
}
function table(book,sheetName,required,label){
 const sheet=book.Sheets[sheetName];
 if(!sheet?.['!ref'])throw Error(label+'是空的。');
 const range=X.utils.decode_range(sheet['!ref']);
 const cell=(r,c)=>sheet[X.utils.encode_cell({r,c})];
 const names=Array.from({length:range.e.c+1},(_,c)=>text(cell(0,c)).trim());
 for(const name of required)if(names.filter(x=>x===name).length!==1)throw Error(label+'第一列需要且只能有一個「'+name+'」欄位。');
 const rows=[];
 for(let r=1;r<=range.e.r;r++){
  if(names.some((_,c)=>{const v=cell(r,c);return v&&(v.f||v.t==='e'||(v.v!=null&&String(v.v).trim()!==''));}))rows.push(r);
 }
 return {sheet,names,rows,cell};
}
function subset(book,source,rows,name){
 const sheet={};
 [0,...rows].forEach((r,out)=>{
  source.names.forEach((_,c)=>{
   const original=source.cell(r,c);if(!original)return;
   text(original);
   // Export stored values: row moves must not change formula references.
   const copy={...original};
   for(const key of ['f','F','D','l','c'])delete copy[key];
   sheet[X.utils.encode_cell({r:out,c})]=copy;
  });
 });
 sheet['!ref']=X.utils.encode_range({s:{r:0,c:0},e:{r:rows.length,c:source.names.length-1}});
 if(source.sheet['!cols'])sheet['!cols']=source.sheet['!cols'].map(c=>({...c}));
 const result=X.utils.book_new();
 result.Workbook={WBProps:{date1904:!!book.Workbook?.WBProps?.date1904}};
 X.utils.book_append_sheet(result,sheet,name);return result;
}
function splitWorkbook(book,sheetName){
 const data=table(book,sheetName,['訂單號碼','選項'],'訂單明細');
 const orderIndex=data.names.indexOf('訂單號碼'),optionIndex=data.names.indexOf('選項');
 const preIds=new Set(),allIds=new Set(),ids=new Map();
 for(const r of data.rows){
  const id=text(data.cell(r,orderIndex)).trim();
  if(!id)throw Error('訂單明細第 '+(r+1)+' 列缺少訂單號碼。');
  ids.set(r,id);allIds.add(id);
  if(/預購|预购/.test(text(data.cell(r,optionIndex))))preIds.add(id);
 }
 const pre=[],stock=[];
 for(const r of data.rows)(preIds.has(ids.get(r))?pre:stock).push(r);
 return {pre:subset(book,data,pre,'預購單'),stock:subset(book,data,stock,'現貨單'),preOrders:preIds.size,stockOrders:allIds.size-preIds.size,preRows:pre.length,stockRows:stock.length,total:data.rows.length};
}
function aggregate(book,sheetName,skuColumn,quantityColumn,label){
 const data=table(book,sheetName,[skuColumn,quantityColumn],label),totals=new Map();
 const skuIndex=data.names.indexOf(skuColumn),quantityIndex=data.names.indexOf(quantityColumn);
 for(const r of data.rows){
  const at=label+'第 '+(r+1)+' 列',sku=text(data.cell(r,skuIndex)).trim(),cell=data.cell(r,quantityIndex);
  if(!sku)throw Error(at+'缺少商品貨號。');
  text(cell);
  let amount;
  if(cell?.t==='n'){
   if(!Number.isSafeInteger(cell.v))throw Error(at+'的「'+quantityColumn+'」必須是有效的安全整數。');
   amount=BigInt(cell.v);
  }else{
   const raw=String(cell?.v??'').trim();
   if(!/^[+-]?\d+$/.test(raw))throw Error(at+'的「'+quantityColumn+'」必須是整數，不能空白。');
   amount=BigInt(raw);
  }
  totals.set(sku,(totals.get(sku)??0n)+amount);
 }
 return {totals,rows:data.rows.length};
}
function inventoryWorkbook(orderBook,orderSheet,inventoryBook,inventorySheet){
 const orders=aggregate(orderBook,orderSheet,'商品貨號','數量','訂單明細');
 const inventory=aggregate(inventoryBook,inventorySheet,'卖家SKU','可用库存','庫存明細');
 const skus=new Set([...inventory.totals.keys(),...orders.totals.keys()]);
 const rows=[['商品貨號','可售庫存','備註']];let missing=0,negative=0;
 for(const sku of skus){
  const absent=!inventory.totals.has(sku),available=(inventory.totals.get(sku)??0n)-(orders.totals.get(sku)??0n);
  if(absent)missing++;if(available<0n)negative++;
  // Excel stores 15 significant digits. Keep larger integers as exact text.
  const value=available>=-999999999999999n&&available<=999999999999999n?Number(available):available.toString();
  rows.push([sku,value,absent?'庫存明細中無此貨號，庫存以 0 計算':'']);
 }
 const sheet=X.utils.aoa_to_sheet(rows);sheet['!cols']=[{wch:26},{wch:20},{wch:56}];
 const book=X.utils.book_new();X.utils.book_append_sheet(book,sheet,'可售庫存');
 return {book,count:skus.size,missing,negative,orderRows:orders.rows,inventoryRows:inventory.rows};
}
function writeWorkbook(book){return X.write(book,{type:'array',bookType:'xlsx',compression:true});}
const api={readWorkbook,splitWorkbook,inventoryWorkbook,writeWorkbook};
if(typeof module==='object'&&module.exports)module.exports=api;else root.OrderTools=api;
})(globalThis);
