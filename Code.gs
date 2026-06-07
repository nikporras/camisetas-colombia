// === CAMISETAS COLOMBIA ===
// Hombre: M L XL XXL  |  Mujer: S M L XL

const STOCK = '📦 Stock';
const VENTAS = '🧾 Ventas';
const GASTOS_SHEET = '💰 Gastos';

const HOMBRE = [['M',20],['L',20],['XL',10],['XXL',10]];
const MUJER  = [['S',15],['M',18],['L',15],['XL',10]];

const ROWS = {
  'Hombre|M':6,'Hombre|L':7,'Hombre|XL':8,'Hombre|XXL':9,
  'Mujer|S':14,'Mujer|M':15,'Mujer|L':16,'Mujer|XL':17
};

function buildSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach((s,i)=>{ if(i>0) ss.deleteSheet(s); });
  ss.getSheets()[0].clear();
  ss.getSheets()[0].setName(STOCK);

  // ----- STOCK -----
  const sh = ss.getSheetByName(STOCK);
  sh.setColumnWidths(1,7,95);
  const AZUL='#003893', ROJO='#CE1126', NEGRO='#1a1a1a', AMARILLO='#FCD116';
  const hdr=['Talla','Stock Inicial','Vendidas','Disponible','% Vendido','¿Bajo Stock?','Precio (£)'];

  sh.getRange('A1:G1').merge().setValue('🇨🇴 CAMISETAS COLOMBIA – INVENTARIO')
    .setBackground(AZUL).setFontColor('#fff').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1,36);

  sh.getRange('A4:G4').merge().setValue('👔  HOMBRE').setBackground(AZUL)
    .setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(5,1,1,7).setValues([hdr]).setBackground(AZUL).setFontColor('#fff').setFontWeight('bold');
  HOMBRE.forEach((m,i)=>{ const r=6+i;
    sh.getRange(r,1,1,7).setValues([[m[0],m[1],0,`=B${r}-C${r}`,`=IFERROR(C${r}/B${r},0)`,`=IF(D${r}<=3,"⚠️ PEDIR","✅ OK")`,0]]);
    sh.getRange(r,5).setNumberFormat('0%'); sh.getRange(r,7).setNumberFormat('£#,##0');
  });
  sh.getRange(10,1,1,7).setValues([['TOTAL HOMBRE','=SUM(B6:B9)','=SUM(C6:C9)','=SUM(D6:D9)','','','=SUMPRODUCT(B6:B9,G6:G9)']])
    .setBackground(AZUL).setFontColor('#fff').setFontWeight('bold');
  sh.getRange(10,7).setNumberFormat('£#,##0');

  sh.getRange('A12:G12').merge().setValue('👚  MUJER').setBackground(ROJO)
    .setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(13,1,1,7).setValues([hdr]).setBackground(ROJO).setFontColor('#fff').setFontWeight('bold');
  MUJER.forEach((m,i)=>{ const r=14+i;
    sh.getRange(r,1,1,7).setValues([[m[0],m[1],0,`=B${r}-C${r}`,`=IFERROR(C${r}/B${r},0)`,`=IF(D${r}<=3,"⚠️ PEDIR","✅ OK")`,0]]);
    sh.getRange(r,5).setNumberFormat('0%'); sh.getRange(r,7).setNumberFormat('£#,##0');
  });
  sh.getRange(18,1,1,7).setValues([['TOTAL MUJER','=SUM(B14:B17)','=SUM(C14:C17)','=SUM(D14:D17)','','','=SUMPRODUCT(B14:B17,G14:G17)']])
    .setBackground(ROJO).setFontColor('#fff').setFontWeight('bold');
  sh.getRange(18,7).setNumberFormat('£#,##0');

  sh.getRange(20,1,1,7).setValues([['🏆 TOTAL GENERAL','=B10+B18','=C10+C18','=D10+D18','','','=G10+G18']])
    .setBackground(NEGRO).setFontColor(AMARILLO).setFontWeight('bold');
  sh.getRange(20,7).setNumberFormat('£#,##0');

  // ----- VENTAS -----
  const v = ss.insertSheet(VENTAS);
  v.getRange('A1:J1').merge().setValue('🇨🇴 CAMISETAS COLOMBIA – VENTAS Y ENVÍOS')
    .setBackground(NEGRO).setFontColor('#fff').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  v.getRange(2,1,1,10).setValues([['#','Fecha','Categoría','Talla','Cantidad','Precio (£)','Total (£)','Cliente','Dirección','Estado']])
    .setBackground(NEGRO).setFontColor('#fff').setFontWeight('bold');
  for(let i=0;i<200;i++){ const r=3+i;
    v.getRange(r,1).setValue(i+1);
    v.getRange(r,7).setFormula(`=IFERROR(E${r}*F${r},"")`).setNumberFormat('£#,##0');
    v.getRange(r,2).setNumberFormat('dd/mm/yyyy'); v.getRange(r,6).setNumberFormat('£#,##0');
  }
  const estados=['Pendiente','Enviada','Entregada'];
  const rule=SpreadsheetApp.newDataValidation().requireValueInList(estados,true).build();
  v.getRange(3,10,200,1).setDataValidation(rule);
  v.setColumnWidth(1,40); v.setColumnWidth(2,90); v.setColumnWidth(3,80);
  v.setColumnWidth(4,55); v.setColumnWidth(5,75); v.setColumnWidth(6,85);
  v.setColumnWidth(7,85); v.setColumnWidth(8,140); v.setColumnWidth(9,200); v.setColumnWidth(10,95);

  // ----- GASTOS -----
  buildGastosSheet_(ss);

  SpreadsheetApp.getUi().alert('✅ Listo. Ahora: Deploy > Manage deployments > editar (lápiz) > New version > Deploy.');
}

function buildGastosSheet_(ss){
  const NEGRO='#1a1a1a', VERDE='#1a4731';
  let g = ss.getSheetByName(GASTOS_SHEET);
  if(!g) g = ss.insertSheet(GASTOS_SHEET);
  g.clear();
  g.getRange('A1:E1').merge().setValue('💰 CAMISETAS COLOMBIA – GASTOS')
    .setBackground(NEGRO).setFontColor('#fff').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  g.getRange(2,1,1,5).setValues([['#','Fecha','Descripción','Categoría','Monto (£)']])
    .setBackground(VERDE).setFontColor('#fff').setFontWeight('bold');
  g.setColumnWidth(1,40); g.setColumnWidth(2,100); g.setColumnWidth(3,220);
  g.setColumnWidth(4,100); g.setColumnWidth(5,100);
  for(let i=0;i<200;i++) g.getRange(3+i,5).setNumberFormat('£#,##0.00');
}

// ===== API para la app =====
function doGet(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName(STOCK);
  const out=[];
  for(const k in ROWS){ const r=ROWS[k], p=k.split('|');
    out.push({cat:p[0],size:p[1],initial:sh.getRange(r,2).getValue(),sold:sh.getRange(r,3).getValue(),left:sh.getRange(r,4).getValue(),price:sh.getRange(r,7).getValue()});
  }
  // envíos
  const v=ss.getSheetByName(VENTAS);
  const data=v.getRange(3,1,200,10).getValues();
  const envios=[];
  data.forEach((row,i)=>{
    if(row[1]!==''){
      envios.push({
        row:i+3,
        fecha:row[1]?Utilities.formatDate(new Date(row[1]),'GMT','dd/MM/yyyy'):'',
        cat:row[2],size:row[3],units:row[4],total:row[6],
        cliente:row[7],direccion:row[8],estado:row[9]||'Pendiente'
      });
    }
  });
  // gastos
  let gastosArr=[];
  try{
    const g=ss.getSheetByName(GASTOS_SHEET);
    if(g){
      const gdata=g.getRange(3,1,200,5).getValues();
      gdata.forEach((row,i)=>{
        if(row[2]!==''){
          gastosArr.push({
            row:i+3,
            fecha:row[1]?Utilities.formatDate(new Date(row[1]),'GMT','dd/MM/yyyy'):'',
            desc:row[2],cat:row[3],monto:row[4]
          });
        }
      });
    }
  }catch(ex){}
  return ContentService.createTextOutput(JSON.stringify({ok:true,stock:out,envios:envios,gastos:gastosArr})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const d=JSON.parse(e.postData.contents);

  // actualizar estado de un envío existente
  if(d.action==='estado'){
    const v=ss.getSheetByName(VENTAS);
    v.getRange(Number(d.row),10).setValue(d.estado);
    return js_({ok:true});
  }

  // registrar gasto
  if(d.action==='gasto'){
    let g=ss.getSheetByName(GASTOS_SHEET);
    if(!g){ buildGastosSheet_(ss); g=ss.getSheetByName(GASTOS_SHEET); }
    let row=3; while(row<=202 && g.getRange(row,3).getValue()!=='') row++;
    g.getRange(row,1).setValue(row-2);
    g.getRange(row,2).setValue(new Date()).setNumberFormat('dd/mm/yyyy');
    g.getRange(row,3).setValue(d.desc||'');
    g.getRange(row,4).setValue(d.cat||'Otro');
    g.getRange(row,5).setValue(Number(d.monto)||0);
    return js_({ok:true,row:row});
  }

  // borrar gasto
  if(d.action==='borrarGasto'){
    const g=ss.getSheetByName(GASTOS_SHEET);
    if(g) g.getRange(Number(d.row),1,1,5).clearContent();
    return js_({ok:true});
  }

  // registrar venta nueva
  const r=ROWS[d.cat+'|'+d.size];
  if(!r) return js_({ok:false,error:'Talla inválida'});
  const s=ss.getSheetByName(STOCK);
  s.getRange(r,3).setValue(s.getRange(r,3).getValue()+Number(d.units));
  const v=ss.getSheetByName(VENTAS);
  let row=3; while(row<=202 && v.getRange(row,2).getValue()!=='') row++;
  v.getRange(row,2,1,5).setValues([[new Date(),d.cat,d.size,Number(d.units),Number(d.price)]]);
  v.getRange(row,8).setValue(d.cliente||'');
  v.getRange(row,9).setValue(d.direccion||'');
  v.getRange(row,10).setValue('Pendiente');
  return js_({ok:true});
}

function js_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
