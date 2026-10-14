/**
 * KDKMP Belanja Desa — Apps Script API + tambahan barang subsidi.
 * Spreadsheet tabs: Produk, Pesanan, BarangSubsidi, PenerimaSubsidi, RiwayatSubsidi.
 * Set Script Property KDKMP_ADMIN_TOKEN (min. 12 chars) before deploying.
 *
 * PENTING: verifikasi NIK via JSONP pada versi frontend ini hanya prototipe.
 * Jangan gunakan data NIK/KK nyata pada website publik sampai endpoint dipindahkan
 * ke backend/proxy same-origin yang aman dan memiliki kontrol akses/rate-limit.
 */
const CONFIG = {
  SPREADSHEET_ID: '1UwhelqYXgJe1bg_a8VjjhZHDsEYTzPdnWSE0aeRDemk',
  PRODUCT_SHEET: 'Produk',
  ORDER_SHEET: 'Pesanan',
  SUBSIDY_PRODUCT_SHEET: 'BarangSubsidi',
  FAMILY_SHEET: 'PenerimaSubsidi',
  SUBSIDY_HISTORY_SHEET: 'RiwayatSubsidi'
};
const PRODUCT_HEADERS = ['SKU','Nama Barang','Kategori','Harga','Image_URL','Stok','Aktif'];
const ORDER_HEADERS = ['ID Pesanan','Timestamp','Nama','WhatsApp','Alamat','Catatan','Item JSON','Total','Status','Tanggal Pengiriman','Jam Pengiriman'];
const SUBSIDY_PRODUCT_HEADERS = ['SKU','Nama Barang','Kategori','Harga','Image_URL','Stok','Cooldown_Hari','Max_Order','Aktif'];
const FAMILY_HEADERS = ['No_KK','Nama_Kepala_Keluarga','NIK','Nama_Anggota','Aktif','Dibuat'];
const SUBSIDY_HISTORY_HEADERS = ['ID Transaksi','Timestamp','SKU','Nama Barang','No_KK','NIK Digunakan','Jumlah','Harga Satuan','Total','Status'];

function doGet(e) {
  try {
    const params=(e&&e.parameter)||{};
    const action=String(params.action||'products');
    let result;
    if(action==='products') result={ok:true,products:readProducts_()};
    else if(action==='subsidyProducts') result={ok:true,products:readSubsidyProducts_().map(publicSubsidyProduct_)};
    else if(action==='subsidyFamilies') { requireAdmin_(params.token); result={ok:true,families:readSubsidyFamilies_()}; }
    else if(action==='verifySubsidy') result=verifySubsidy_(params.sku,params.identity,params.qty);
    else if(action==='orderStatus') {
      const id=clean_(params.id); if(!id) throw new Error('ID pesanan wajib diisi.');
      const sh=getSheet_(CONFIG.ORDER_SHEET,ORDER_HEADERS);
      result={ok:true,id,exists:sh.getDataRange().getDisplayValues().slice(1).some(r=>String(r[0])===id)};
    } else if(action==='ping') result={ok:true,message:'KDKMP API aktif'};
    else throw new Error('Aksi GET tidak dikenal.');
    return output_(result,params.callback);
  } catch(err) { return output_({ok:false,error:String(err.message||err)},e&&e.parameter&&e.parameter.callback); }
}
function doPost(e) {
  try {
    const payload=JSON.parse(e&&e.postData&&e.postData.contents||'{}');
    const action=clean_(payload.action);
    if(['saveProduct','setActive','importProducts','saveSubsidyProduct','saveSubsidyFamily','setSubsidyProductActive','setSubsidyFamilyActive'].includes(action)) requireAdmin_(payload.token);
    let result;
    if(action==='saveProduct') result={ok:true,product:saveProduct_(payload.product||{})};
    else if(action==='setActive') result={ok:true,product:setActive_(payload.sku,payload.active)};
    else if(action==='importProducts') result={ok:true,imported:importProducts_(payload.products)};
    else if(action==='createOrder') result={ok:true,order:createOrder_(payload.order||{})};
    else if(action==='saveSubsidyProduct') result={ok:true,product:saveSubsidyProduct_(payload.product||{})};
    else if(action==='setSubsidyProductActive') result={ok:true,product:setSubsidyProductActive_(payload.sku,payload.active)};
    else if(action==='setSubsidyFamilyActive') result={ok:true,family:setSubsidyFamilyActive_(payload.kk,payload.active)};
    else if(action==='saveSubsidyFamily') result={ok:true,family:saveSubsidyFamily_(payload.kk,payload.familyName,payload.members)};
    else if(action==='createSubsidyOrder') result={ok:true,order:createSubsidyOrder_(payload.order||{})};
    else throw new Error('Aksi POST tidak dikenal.');
    return json_(result);
  } catch(err) { return json_({ok:false,error:String(err.message||err)}); }
}
function setupKDKMP() {
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ss.setSpreadsheetTimeZone('Asia/Jakarta');
  migrateProductSheet_(ss);
  ensureSheet_(ss,CONFIG.PRODUCT_SHEET,PRODUCT_HEADERS);
  ensureSheet_(ss,CONFIG.ORDER_SHEET,ORDER_HEADERS);
  ensureSheet_(ss,CONFIG.SUBSIDY_PRODUCT_SHEET,SUBSIDY_PRODUCT_HEADERS);
  ensureSheet_(ss,CONFIG.FAMILY_SHEET,FAMILY_HEADERS);
  ensureSheet_(ss,CONFIG.SUBSIDY_HISTORY_SHEET,SUBSIDY_HISTORY_HEADERS);
  ss.getSheetByName(CONFIG.ORDER_SHEET).getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  ss.getSheetByName(CONFIG.SUBSIDY_HISTORY_SHEET).getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  Logger.log('KDKMP siap: '+ss.getUrl());
}

function migrateProductSheet_(ss){
 const sh=ss.getSheetByName(CONFIG.PRODUCT_SHEET);if(!sh||sh.getLastRow()===0)return;
 const legacy=['SKU','Nama Barang','Kategori','Subkategori','Merek','Satuan','Harga','Image_URL','Status','Max_Order','Aktif'];
 const header=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
 if(PRODUCT_HEADERS.every((h,i)=>header[i]===h)&&header.length===PRODUCT_HEADERS.length)return;
 if(!legacy.every((h,i)=>header[i]===h))throw new Error('Header sheet Produk tidak cocok dengan format lama maupun baru.');
 const old=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,legacy.length).getValues():[];
 const rows=old.filter(r=>clean_(r[0])&&clean_(r[1])).map(r=>[r[0],r[1],r[2]||'Lainnya',Number(r[6])||0,r[7]||'',Math.max(0,Math.floor(Number(r[9])||0)),normalActive_(r[10])]);
 sh.clearContents();sh.getRange(1,1,1,PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
 if(rows.length)sh.getRange(2,1,rows.length,PRODUCT_HEADERS.length).setValues(rows);
 sh.setFrozenRows(1);sh.getRange(1,1,1,PRODUCT_HEADERS.length).setFontWeight('bold');
}

function ensureSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold');return sh;}
  const current=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getDisplayValues()[0].slice(0,headers.length);
  if(headers.some((h,i)=>String(current[i]||'').trim()!==h))throw new Error('Header sheet "'+name+'" tidak cocok. Dibutuhkan: '+headers.join(' | '));
  return sh;
}
function getSheet_(name,headers){
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID),sh=ss.getSheetByName(name);
  if(!sh)throw new Error('Sheet "'+name+'" tidak ditemukan. Jalankan setupKDKMP().');
  const current=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getDisplayValues()[0].slice(0,headers.length);
  if(headers.some((h,i)=>String(current[i]||'').trim()!==h))throw new Error('Header sheet "'+name+'" tidak sesuai. Jalankan pemeriksaan header.');
  return sh;
}
function readProducts_(){
  const sh=getSheet_(CONFIG.PRODUCT_SHEET,PRODUCT_HEADERS);if(sh.getLastRow()<2)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,PRODUCT_HEADERS.length).getValues().filter(r=>clean_(r[0])&&clean_(r[1])).map(r=>{
    const stock=Math.max(0,Math.floor(Number(r[5])||0));
    return {sku:String(r[0]),name:String(r[1]),category:String(r[2]||'Lainnya'),price:Number(r[3])||0,image:String(r[4]||''),stock,subcategory:'',brand:'',unit:'pcs',status:stock>0?'TERSEDIA':'HABIS',maxOrder:stock,active:normalActive_(r[6])};
  });
}
function saveProduct_(p){
  const product=validateProduct_(p),sh=getSheet_(CONFIG.PRODUCT_SHEET,PRODUCT_HEADERS),lock=LockService.getScriptLock();lock.waitLock(10000);
  try{const last=sh.getLastRow();let row=0;if(last>=2){const i=sh.getRange(2,1,last-1,1).getDisplayValues().findIndex(r=>String(r[0]).trim()===product.sku);if(i>=0)row=i+2;}if(!row)row=Math.max(2,last+1);sh.getRange(row,1,1,PRODUCT_HEADERS.length).setValues([[product.sku,product.name,product.category,product.price,product.image,product.stock,product.active]]);return product;}finally{lock.releaseLock();}
}
function setActive_(sku,active){
  sku=clean_(sku);if(!sku)throw new Error('SKU wajib diisi.');const sh=getSheet_(CONFIG.PRODUCT_SHEET,PRODUCT_HEADERS);if(sh.getLastRow()<2)throw new Error('Produk belum ada.');
  const i=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().findIndex(r=>String(r[0]).trim()===sku);if(i<0)throw new Error('SKU tidak ditemukan.');
  sh.getRange(i+2,7).setValue(normalActive_(active));return readProducts_().find(p=>p.sku===sku);
}
function importProducts_(list){if(!Array.isArray(list)||!list.length)throw new Error('Daftar produk kosong.');if(list.length>300)throw new Error('Maksimal 300 produk per impor.');list.forEach(saveProduct_);return list.length;}
function validateProduct_(p){
  const sku=clean_(p.sku),name=clean_(p.name),price=Number(p.price),stock=Number(p.stock),image=clean_(p.image||p.image_url||'');
  if(!sku||sku.length>60)throw new Error('SKU wajib diisi (maksimal 60 karakter).');
  if(!name||name.length>160)throw new Error('Nama produk wajib diisi (maksimal 160 karakter).');
  if(!Number.isFinite(price)||price<0)throw new Error('Harga tidak valid.');
  if(!Number.isInteger(stock)||stock<0)throw new Error('Stok harus bilangan bulat minimal 0.');
  if(image&&!/^https:\/\//i.test(image))throw new Error('URL gambar harus HTTPS.');
  return {sku,name,category:clean_(p.category)||'Lainnya',price,image,stock,active:normalActive_(p.active)};
}
function createOrder_(o){
  const id=clean_(o.id),name=clean_(o.name),phone=clean_(o.whatsapp),address=clean_(o.address),note=clean_(o.note);
  const today=Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd'),deliveryDate=clean_(o.deliveryDate)||today,deliveryTime=clean_(o.deliveryTime);
  const validTimes=['09.00 - 10.00 WIB','10.00 - 11.00 WIB','11.00 - 12.00 WIB','12.00 - 13.00 WIB','13.00 - 14.00 WIB','14.00 - 15.00 WIB','15.00 - 16.00 WIB','16.00 - 17.00 WIB','17.00 - 18.00 WIB'];
  if(!/^KDKMP-[A-Z0-9-]{5,60}$/.test(id))throw new Error('ID pesanan tidak valid.');
  if(!name||name.length>80)throw new Error('Nama penerima wajib diisi.');
  if(!/^62\d{8,13}$/.test(phone))throw new Error('Nomor WhatsApp tidak valid.');
  if(!address||address.length>500)throw new Error('Alamat wajib diisi.');
  if(note.length>300)throw new Error('Catatan terlalu panjang.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(deliveryDate)||deliveryDate<today)throw new Error('Tanggal pengiriman tidak valid.');
  if(!validTimes.includes(deliveryTime))throw new Error('Jam pengiriman tidak valid.');
  if(deliveryDate===today){const parts=Utilities.formatDate(new Date(),'Asia/Jakarta','HH:mm').split(':').map(Number),mins=parts[0]*60+parts[1],start=Number(deliveryTime.slice(0,2))*60+Number(deliveryTime.slice(3,5));if(start<=mins)throw new Error('Jam pengiriman sudah dimulai atau lewat.');}
  if(!Array.isArray(o.items)||!o.items.length||o.items.length>100)throw new Error('Item pesanan tidak valid.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const orders=getSheet_(CONFIG.ORDER_SHEET,ORDER_HEADERS);if(orders.getLastRow()>=2&&orders.getRange(2,1,orders.getLastRow()-1,1).getDisplayValues().some(r=>String(r[0])===id))return {id,duplicate:true};
    const regularSheet=getSheet_(CONFIG.PRODUCT_SHEET,PRODUCT_HEADERS),regularRows=regularSheet.getLastRow()>1?regularSheet.getRange(2,1,regularSheet.getLastRow()-1,PRODUCT_HEADERS.length).getValues():[],regular={};
    regularRows.forEach((r,i)=>{const sku=clean_(r[0]);if(sku)regular[sku]={row:i+2,sku,name:String(r[1]||''),price:Number(r[3])||0,stock:Math.max(0,Math.floor(Number(r[5])||0)),active:normalActive_(r[6])};});
    const subSheet=getSheet_(CONFIG.SUBSIDY_PRODUCT_SHEET,SUBSIDY_PRODUCT_HEADERS),subRows=subSheet.getLastRow()>1?subSheet.getRange(2,1,subSheet.getLastRow()-1,SUBSIDY_PRODUCT_HEADERS.length).getValues():[],subs={};
    subRows.forEach((r,i)=>{const sku=clean_(r[0]);if(sku)subs[sku]={row:i+2,sku,name:String(r[1]||''),price:Number(r[3])||0,stock:Math.max(0,Math.floor(Number(r[5])||0)),cooldownDays:Math.max(1,Math.floor(Number(r[6])||1)),maxOrder:Math.max(1,Math.floor(Number(r[7])||1)),active:normalActive_(r[8])};});
    const requested={},items=[];
    o.items.forEach(it=>{const sku=clean_(it.sku),qty=Number(it.qty),isSubsidy=it.isSubsidy===true;if(!sku||!Number.isInteger(qty)||qty<1)throw new Error('Jumlah/SKU barang tidak valid.');const key=(isSubsidy?'SUB:':'REG:')+sku;requested[key]=(requested[key]||0)+qty;if(isSubsidy){const p=subs[sku],identity=clean_(it.identity).replace(/\D/g,'');if(!p||p.active!=='AKTIF'||p.stock<qty)throw new Error('Barang subsidi tidak tersedia atau stok tidak cukup: '+sku);if(!/^\d{16}$/.test(identity))throw new Error('NIK subsidi belum diverifikasi.');if(qty>p.maxOrder)throw new Error('Jumlah barang subsidi melebihi kuota.');const family=findFamilyByIdentity_(identity);if(!family)throw new Error('NIK/KK tidak terdaftar sebagai penerima subsidi.');const last=latestFamilyPurchase_(sku,family.kk);if(last&&Date.now()-last.getTime()<p.cooldownDays*86400000)throw new Error('Kuota keluarga untuk '+p.name+' sudah digunakan atau masih dalam masa tunggu.');items.push({sku,name:p.name,qty,unit:'pcs',price:p.price,subtotal:p.price*qty,isSubsidy:true,identity,kk:family.kk,row:p.row});}else{const p=regular[sku];if(!p||p.active!=='AKTIF'||p.stock<qty)throw new Error('Produk tidak tersedia atau stok tidak cukup: '+sku);items.push({sku,name:p.name,qty,unit:'pcs',price:p.price,subtotal:p.price*qty,isSubsidy:false,row:p.row});}});
    const total=items.reduce((s,x)=>s+x.subtotal,0),sheetItems=items.map(({identity,kk,row,...item})=>item);
    orders.appendRow([id,new Date(),name,phone,address,note,JSON.stringify(sheetItems),total,'MENUNGGU KONFIRMASI ADMIN',deliveryDate,deliveryTime]);
    const history=getSheet_(CONFIG.SUBSIDY_HISTORY_SHEET,SUBSIDY_HISTORY_HEADERS);
    items.forEach(it=>{if(it.isSubsidy){subSheet.getRange(it.row,6).setValue(subs[it.sku].stock-it.qty);history.appendRow([id,new Date(),it.sku,it.name,it.kk,it.identity,it.qty,it.price,it.subtotal,'MENUNGGU KONFIRMASI ADMIN']);}else{regularSheet.getRange(it.row,6).setValue(regular[it.sku].stock-it.qty);}});
    SpreadsheetApp.flush();return {id,total,status:'MENUNGGU KONFIRMASI ADMIN',items:items.length,subsidyItems:items.filter(x=>x.isSubsidy).length};
  }finally{lock.releaseLock();}
}
/* SUBSIDI */
function readSubsidyProducts_(){
 const sh=getSheet_(CONFIG.SUBSIDY_PRODUCT_SHEET,SUBSIDY_PRODUCT_HEADERS);if(sh.getLastRow()<2)return [];
 return sh.getRange(2,1,sh.getLastRow()-1,SUBSIDY_PRODUCT_HEADERS.length).getValues().filter(r=>clean_(r[0])&&clean_(r[1])).map(r=>({sku:String(r[0]),name:String(r[1]),category:String(r[2]||'Subsidi'),price:Number(r[3])||0,image:String(r[4]||''),stock:Math.max(0,Math.floor(Number(r[5])||0)),cooldownDays:Math.max(1,Math.floor(Number(r[6])||1)),maxOrder:Math.max(1,Math.floor(Number(r[7])||1)),active:normalActive_(r[8])}));
}
function publicSubsidyProduct_(p){return {sku:p.sku,name:p.name,category:p.category,price:p.price,image:p.image,stock:p.stock,cooldownDays:p.cooldownDays,maxOrder:p.maxOrder,active:p.active};}
function saveSubsidyProduct_(p){
 const sku=clean_(p.sku),name=clean_(p.name),price=Number(p.price),stock=Number(p.stock),cool=Number(p.cooldownDays),max=Number(p.maxOrder),image=clean_(p.image);
 if(!sku||sku.length>60||!name||name.length>160)throw new Error('SKU/nama barang subsidi wajib diisi.');
 if(!Number.isFinite(price)||price<0||!Number.isInteger(stock)||stock<0||!Number.isInteger(cool)||cool<1||!Number.isInteger(max)||max<1)throw new Error('Harga/stok/cooldown/batas pembelian tidak valid.');
 if(image&&!/^https:\/\//i.test(image))throw new Error('URL gambar harus HTTPS.');
 const product={sku,name,category:clean_(p.category)||'Subsidi',price,image,stock,cooldownDays:cool,maxOrder:max,active:normalActive_(p.active)};
 const sh=getSheet_(CONFIG.SUBSIDY_PRODUCT_SHEET,SUBSIDY_PRODUCT_HEADERS),last=sh.getLastRow();let row=0;
 if(last>=2){const i=sh.getRange(2,1,last-1,1).getDisplayValues().findIndex(r=>String(r[0]).trim()===sku);if(i>=0)row=i+2;}if(!row)row=Math.max(2,last+1);
 sh.getRange(row,1,1,SUBSIDY_PRODUCT_HEADERS.length).setValues([[sku,name,product.category,price,image,stock,cool,max,product.active]]);return product;
}
function saveSubsidyFamily_(kk,familyName,members){
 kk=clean_(kk);familyName=clean_(familyName);if(!/^\d{16}$/.test(kk))throw new Error('Nomor KK harus 16 digit.');
 if(!Array.isArray(members)||!members.length||members.length>30)throw new Error('Anggota keluarga tidak valid (maksimal 30).');
 const seen={};members.forEach(m=>{m.nik=clean_(m.nik);m.name=clean_(m.name);if(!/^\d{16}$/.test(m.nik)||!m.name||m.name.length>100)throw new Error('Setiap anggota wajib memiliki NIK 16 digit dan nama.');if(seen[m.nik])throw new Error('NIK anggota duplikat.');seen[m.nik]=true;});
 const sh=getSheet_(CONFIG.FAMILY_SHEET,FAMILY_HEADERS),values=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,FAMILY_HEADERS.length).getValues():[];
 // Replace all existing rows for this KK; retain other families.
 const keep=values.filter(r=>String(r[0])!==kk),now=new Date();
 const add=members.map(m=>[kk,familyName,m.nik,m.name,'AKTIF',now]);
 if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,FAMILY_HEADERS.length).clearContent();
 const combined=keep.concat(add);if(combined.length)sh.getRange(2,1,combined.length,FAMILY_HEADERS.length).setValues(combined);
 return {kkMasked:kk.slice(0,4)+'********'+kk.slice(-4),memberCount:members.length};
}

function readSubsidyFamilies_(){
 const sh=getSheet_(CONFIG.FAMILY_SHEET,FAMILY_HEADERS);if(sh.getLastRow()<2)return [];
 const vals=sh.getRange(2,1,sh.getLastRow()-1,FAMILY_HEADERS.length).getValues(),map={};
 vals.forEach(r=>{const kk=String(r[0]||'');if(!kk)return;if(!map[kk])map[kk]={kk,familyName:String(r[1]||''),members:[],active:'NONAKTIF'};map[kk].members.push({nik:String(r[2]||''),name:String(r[3]||'')});if(normalActive_(r[4])==='AKTIF')map[kk].active='AKTIF';});
 return Object.values(map).map(f=>({...f,kkMasked:f.kk.slice(0,4)+'********'+f.kk.slice(-4),memberCount:f.members.length}));
}
function setSubsidyProductActive_(sku,active){
 sku=clean_(sku);if(!sku)throw new Error('SKU wajib diisi.');const sh=getSheet_(CONFIG.SUBSIDY_PRODUCT_SHEET,SUBSIDY_PRODUCT_HEADERS);if(sh.getLastRow()<2)throw new Error('Belum ada barang subsidi.');
 const rows=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues(),i=rows.findIndex(r=>String(r[0]).trim()===sku);if(i<0)throw new Error('SKU subsidi tidak ditemukan.');sh.getRange(i+2,9).setValue(normalActive_(active));return readSubsidyProducts_().find(p=>p.sku===sku);
}
function setSubsidyFamilyActive_(kk,active){
 kk=clean_(kk);if(!/^\d{16}$/.test(kk))throw new Error('Nomor KK tidak valid.');const sh=getSheet_(CONFIG.FAMILY_SHEET,FAMILY_HEADERS);if(sh.getLastRow()<2)throw new Error('Belum ada penerima subsidi.');
 const vals=sh.getRange(2,1,sh.getLastRow()-1,FAMILY_HEADERS.length).getValues();let count=0;vals.forEach((r,i)=>{if(String(r[0])===kk){sh.getRange(i+2,5).setValue(normalActive_(active));count++;}});if(!count)throw new Error('Nomor KK tidak ditemukan.');return {kkMasked:kk.slice(0,4)+'********'+kk.slice(-4),active:normalActive_(active)};
}
function verifySubsidy_(sku,identity,qty){
 sku=clean_(sku);identity=clean_(identity).replace(/\D/g,'');qty=Number(qty);
 if(!/^\d{16}$/.test(identity))throw new Error('NIK atau KK harus 16 digit.');
 if(!Number.isInteger(qty)||qty<1)throw new Error('Jumlah pembelian tidak valid.');
 const products=readSubsidyProducts_(),p=products.find(x=>x.sku===sku&&x.active==='AKTIF');
 if(!p||p.stock<1)throw new Error('Barang subsidi tidak tersedia.');
 if(qty>p.maxOrder)throw new Error('Pesanan Anda melebihi batas maksimal.');
 const family=findFamilyByIdentity_(identity);if(!family)throw new Error('NIK/KK tidak terdaftar sebagai penerima subsidi.');
 if(qty>p.stock)throw new Error('Stok barang subsidi tidak mencukupi.');
 const latest=latestFamilyPurchase_(sku,family.kk);
 if(latest&&Date.now()-latest.getTime()<p.cooldownDays*86400000){
   const next=new Date(latest.getTime()+p.cooldownDays*86400000);
   return {ok:true,eligible:false,message:'Kuota keluarga ini sudah digunakan. Coba kembali setelah '+Utilities.formatDate(next,'Asia/Jakarta','dd/MM/yyyy HH:mm')+' WIB.'};
 }
 return {ok:true,eligible:true,message:'Verifikasi berhasil',familyName:family.familyName,product:publicSubsidyProduct_(p)};
}
function findFamilyByIdentity_(identity){
 const sh=getSheet_(CONFIG.FAMILY_SHEET,FAMILY_HEADERS);if(sh.getLastRow()<2)return null;
 const rows=sh.getRange(2,1,sh.getLastRow()-1,FAMILY_HEADERS.length).getValues();
 const r=rows.find(x=>String(x[0])===identity&&normalActive_(x[4])==='AKTIF')||rows.find(x=>String(x[2])===identity&&normalActive_(x[4])==='AKTIF');
 return r?{kk:String(r[0]),familyName:String(r[1]||'')}:null;
}
function latestFamilyPurchase_(sku,kk){
 const sh=getSheet_(CONFIG.SUBSIDY_HISTORY_SHEET,SUBSIDY_HISTORY_HEADERS);if(sh.getLastRow()<2)return null;
 const rows=sh.getRange(2,1,sh.getLastRow()-1,SUBSIDY_HISTORY_HEADERS.length).getValues();
 const dates=rows.filter(r=>String(r[2])===sku&&String(r[4])===kk&&['DITERIMA','MENUNGGU KONFIRMASI ADMIN'].includes(String(r[9]))).map(r=>r[1] instanceof Date?r[1]:new Date(r[1])).filter(d=>!isNaN(d.getTime()));
 return dates.length?new Date(Math.max(...dates.map(d=>d.getTime()))):null;
}
function createSubsidyOrder_(o){
 const id=clean_(o.id),sku=clean_(o.sku),identity=clean_(o.identity).replace(/\D/g,''),qty=Number(o.qty),name=clean_(o.name),phone=clean_(o.whatsapp),address=clean_(o.address);
 if(!/^KDKMP-[A-Z0-9-]{5,60}$/.test(id))throw new Error('ID transaksi tidak valid.');
 if(!name||name.length>80||!/^62\d{8,13}$/.test(phone)||!address||address.length>500)throw new Error('Nama, WhatsApp, atau alamat tidak valid.');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const hist=getSheet_(CONFIG.SUBSIDY_HISTORY_SHEET,SUBSIDY_HISTORY_HEADERS);
  if(hist.getLastRow()>1&&hist.getRange(2,1,hist.getLastRow()-1,1).getDisplayValues().some(r=>String(r[0])===id))return {id,duplicate:true};
  const products=readSubsidyProducts_(),p=products.find(x=>x.sku===sku&&x.active==='AKTIF');
  if(!p||p.stock<1)throw new Error('Barang subsidi tidak tersedia.');
  if(!Number.isInteger(qty)||qty<1||qty>p.maxOrder)throw new Error('Pesanan Anda melebihi batas maksimal.');
  if(qty>p.stock)throw new Error('Stok barang subsidi tidak mencukupi.');
  const family=findFamilyByIdentity_(identity);if(!family)throw new Error('NIK/KK tidak terdaftar sebagai penerima subsidi.');
  const last=latestFamilyPurchase_(sku,family.kk);if(last&&Date.now()-last.getTime()<p.cooldownDays*86400000)throw new Error('Kuota keluarga ini sudah digunakan; masa tunggu belum selesai.');
  const productSheet=getSheet_(CONFIG.SUBSIDY_PRODUCT_SHEET,SUBSIDY_PRODUCT_HEADERS);
  const idx=productSheet.getRange(2,1,productSheet.getLastRow()-1,1).getDisplayValues().findIndex(r=>String(r[0])===sku);
  if(idx<0)throw new Error('SKU subsidi tidak ditemukan.');
  // Status MENUNGGU KONFIRMASI tidak dihitung sebagai pembelian kuota sampai admin menyetujui.
  // Stok tidak dikurangi di sini; petugas harus mengonfirmasi ketersediaan fisik.
  hist.appendRow([id,new Date(),sku,p.name,family.kk,identity,qty,p.price,p.price*qty,'MENUNGGU KONFIRMASI ADMIN']);
  SpreadsheetApp.flush();return {id,total:p.price*qty,status:'MENUNGGU KONFIRMASI ADMIN'};
 }finally{lock.releaseLock();}
}
function clean_(v){return String(v==null?'':v).trim();}
function normalActive_(v){return /nonaktif|non-active|inactive|false|0/i.test(clean_(v))?'NONAKTIF':'AKTIF';}
function requireAdmin_(token){
 const expected=PropertiesService.getScriptProperties().getProperty('KDKMP_ADMIN_TOKEN');
 if(!expected||expected.length<12)throw new Error('Token admin belum disetel pada Script Properties.');
 if(String(token||'')!==expected)throw new Error('Token admin salah.');
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function output_(obj,callback){
 const json=JSON.stringify(obj);
 if(callback&&/^[A-Za-z_$][0-9A-Za-z_$\.]{0,80}$/.test(callback))return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
 return json_(obj);
}
