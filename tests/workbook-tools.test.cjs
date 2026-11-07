const {test}=require('node:test');
const assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js');
const T=require('../workbook-tools.js');
function book(rows,name='資料'){
 const result=X.utils.book_new();X.utils.book_append_sheet(result,X.utils.aoa_to_sheet(rows),name);return result;
}
function roundtrip(b){return T.readWorkbook(T.writeWorkbook(b));}
function rows(b){return X.utils.sheet_to_json(b.Sheets[b.SheetNames[0]],{header:1,defval:''});}
test('xlsx roundtrip splits entire orders and preserves cell values, formats and order',()=>{
 const b=book([['訂單號碼','選項','商品貨號','數量','日期','備註'],['0001','現貨','0012',2,46000,'中文,引號"\n換行'],['0002','現貨','0023',3,46001,'=文字'],['0001','预购','0004',1,46002,''],['0003','預購','0005',4,46003,'']]);
 b.Sheets.資料.E2.z='yyyy-mm-dd';b.Sheets.資料.D2={t:'n',f:'1+1',v:2,z:'0'};
 b.Sheets.資料.A4={t:'n',v:1,z:'0000'};
 b.Workbook={WBProps:{date1904:true}};
 const input=roundtrip(b),data=T.splitWorkbook(input,'資料');
 assert.deepEqual([data.preOrders,data.stockOrders,data.preRows,data.stockRows],[2,1,3,1]);
 const pre=roundtrip(data.pre),stock=roundtrip(data.stock);
 assert.deepEqual(rows(pre).map(r=>r[2]),['商品貨號','0012','0004','0005']);
 assert.equal(rows(stock)[1][0],'0002');assert.equal(rows(stock)[1][5],'=文字');
 assert.equal(pre.Sheets.預購單.A2.t,'s');assert.equal(pre.Sheets.預購單.A3.z,'0000');
 assert.equal(pre.Sheets.預購單.E2.z,'yyyy-mm-dd');assert.equal(pre.Workbook.WBProps.date1904,true);
 assert.equal(pre.Sheets.預購單.D2.v,2);assert.equal(pre.Sheets.預購單.D2.f,undefined);
 assert.equal(rows(pre)[1][5],'中文,引號"\n換行');assert.equal(input.Sheets.資料.D2.f,'1+1');
});
test('empty classification still exports headers; blank worksheet rows are ignored',()=>{
 const data=T.splitWorkbook(book([['訂單號碼','選項'],['001','現貨'],[],['002','現貨']]),'資料');
 assert.equal(data.total,2);assert.deepEqual(rows(roundtrip(data.pre)),[['訂單號碼','選項']]);
 assert.equal(T.splitWorkbook(book([['訂單號碼','選項']]),'資料').total,0);
});
test('inventory sums duplicates, preserves SKU zeros, missing, negative and exact large integers',()=>{
 const orders=book([['商品貨號','數量'],['0012',2],['0012',3],['MISSING',4],['big','10000000000000000']]);
 const inventory=book([['卖家SKU','可用库存'],[12,10],['0012',1],['ONLY',8],['big','20000000000000000']]);
 inventory.Sheets.資料.A2.z='0000';inventory.Sheets.資料.B2.z='#,##0.00';
 const data=T.inventoryWorkbook(roundtrip(orders),'資料',roundtrip(inventory),'資料');
 assert.deepEqual([data.count,data.missing,data.negative,data.orderRows,data.inventoryRows],[4,1,1,4,4]);
 assert.deepEqual(rows(roundtrip(data.book)),[['商品貨號','可售庫存','備註'],['0012',6,''],['ONLY',8,''],['big','10000000000000000',''],['MISSING',-4,'庫存明細中無此貨號，庫存以 0 計算']]);
 assert.equal(data.book.Sheets.可售庫存.A2.t,'s');assert.equal(data.book.Sheets.可售庫存.B2.t,'n');
});
test('selecting another worksheet uses only that worksheet',()=>{
 const b=book([['說明'],['請選訂單頁']],'說明');
 X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([['訂單號碼','選項'],['001','預購']]),'訂單');
 assert.throws(()=>T.splitWorkbook(b,'說明'),/第一列/);assert.equal(T.splitWorkbook(roundtrip(b),'訂單').preOrders,1);
});
test('legacy XLS orders split and export to XLSX with Chinese, zeros and multiple sheets',()=>{
 const b=book([['說明']],'說明');
 X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([['訂單號碼','選項','商品貨號','數量'],['0001','現貨','0012',2],['0002','現貨','0023',1],['0001','預購','0012',3]]),'訂單');
 const bytes=X.write(b,{type:'array',bookType:'biff8'});
 assert.equal(new Uint8Array(bytes)[0],0xd0);
 const input=T.readWorkbook(bytes),result=T.splitWorkbook(input,'訂單');
 assert.deepEqual(input.SheetNames,['說明','訂單']);assert.equal(input.cfb,undefined);
 assert.deepEqual([result.preOrders,result.stockOrders,result.preRows],[1,1,2]);
 assert.deepEqual(rows(roundtrip(result.pre)),[['訂單號碼','選項','商品貨號','數量'],['0001','現貨','0012',2],['0001','預購','0012',3]]);
 assert.equal(new Uint8Array(T.writeWorkbook(result.pre))[0],0x50);
});
test('XLS and XLSX inventory inputs work in either combination',()=>{
 const orders=book([['商品貨號','數量'],['0012',2],['0012',3],['缺貨',1]]);
 const inventory=book([['卖家SKU','可用库存'],[12,10]]);inventory.Sheets.資料.A2.z='0000';
 for(const orderFormat of ['biff8','xlsx'])for(const stockFormat of ['biff8','xlsx']){
  const read=(b,format)=>T.readWorkbook(X.write(b,{type:'array',bookType:format}));
  const result=T.inventoryWorkbook(read(orders,orderFormat),'資料',read(inventory,stockFormat),'資料');
  assert.deepEqual(rows(roundtrip(result.book)),[['商品貨號','可售庫存','備註'],['0012',5,''],['缺貨',-1,'庫存明細中無此貨號，庫存以 0 計算']]);
 }
});
test('older BIFF XLS binary variants can be read',()=>{
 const b=book([['ID','value'],['001',2]]);
 for(const bookType of ['biff2','biff3','biff4','biff5']){
  const result=T.readWorkbook(X.write(b,{type:'array',bookType}));
  assert.deepEqual(rows(result),[['ID','value'],['001',2]]);
 }
});
test('invalid files, headers, order IDs, quantities and uncached formulas are rejected',()=>{
 assert.throws(()=>T.readWorkbook(new TextEncoder().encode('訂單號碼,選項\n001,預購')),/xlsx/);
 assert.throws(()=>T.readWorkbook(new Uint8Array([80,75,0,0])),/無法讀取/);
 assert.throws(()=>T.splitWorkbook(book([['訂單號碼','選項','選項']]),'資料'),/只能有一個/);
 assert.throws(()=>T.splitWorkbook(book([['訂單號碼','選項'],['','預購']]),'資料'),/第 2 列缺少/);
 const inventory=book([['卖家SKU','可用库存'],['001',10]]);
 for(const value of ['',1.5,'1.2',true,Number.MAX_SAFE_INTEGER+1]){
  assert.throws(()=>T.inventoryWorkbook(book([['商品貨號','數量'],['001',value]]),'資料',inventory,'資料'),/整數/);
 }
 const formula=book([['訂單號碼','選項','計算'],['001','預購']]);
 formula.Sheets.資料.C2={t:'n',f:'1+1'};
 assert.throws(()=>T.splitWorkbook(formula,'資料'),/計算結果/);
 formula.Sheets.資料.C2={t:'e',v:7};assert.throws(()=>T.splitWorkbook(formula,'資料'),/Excel 錯誤/);
});
