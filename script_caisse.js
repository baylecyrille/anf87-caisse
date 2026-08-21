// ═══════════════════════════════════════════════════════
//  ANF 87 — Google Apps Script CAISSE v2
//  Sheets: Produits | Sites | Categories | Ventes | Utilisateurs
// ═══════════════════════════════════════════════════════

// Actions qui ÉCRIVENT dans le classeur : protégées par un verrou pour éviter
// les conflits quand 2-3 utilisateurs utilisent la caisse en même temps.
var WRITE_ACTIONS = ["saveProduct","deleteProduct","saveSite","deleteSite","saveCategory",
  "deleteCategory","saveUser","deleteUser","saveSale","updateStock","transferStock","initSheets"];

function doGet(e) {
  var p = e.parameter, action = p.action || "getAll", result;
  var lock = null;
  try {
    if (WRITE_ACTIONS.indexOf(action) !== -1) {
      lock = LockService.getScriptLock();
      lock.waitLock(15000); // attend jusqu'à 15s si un autre utilisateur écrit en même temps
    }
    switch(action) {
      case "getAll":          result = getAllData(); break;
      case "saveProduct":     result = saveProduct(e); break;
      case "deleteProduct":   result = deleteProduct(e); break;
      case "saveSite":        result = saveSite(e); break;
      case "deleteSite":      result = deleteSite(e); break;
      case "saveCategory":    result = saveCategory(e); break;
      case "deleteCategory":  result = deleteCategory(e); break;
      case "saveUser":        result = saveUser(e); break;
      case "deleteUser":      result = deleteUser(e); break;
      case "saveSale":        result = saveSale(e); break;
      case "updateStock":     result = updateStock(e); break;
      case "transferStock":   result = transferStock(e); break;
      case "getSales":        result = getSales(e); break;
      case "initSheets":      result = initSheets(); break;
      default: result = {ok:false, error:"Action inconnue: "+action};
    }
  } catch(err) { result = {ok:false, error:err.toString()}; }
  finally { if (lock) lock.releaseLock(); }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Sites
  var sites = getOrCreate(ss,"Sites");
  if(sites.getLastRow()<=1){
    sites.getRange(1,1,1,4).setValues([["ID","Nom","Village","Couleur"]]);
    sites.getRange(1,1,1,4).setFontWeight("bold");
    sites.getRange(2,1,4,4).setValues([
      ["s1","Buvette Nord","Village 1","#CC0000"],
      ["s2","Buvette Est","Village 2","#1565C0"],
      ["s3","Buvette Sud","Village 3","#2E7D32"],
      ["s4","Buvette Ouest","Village 4","#E65100"]
    ]);
  }
  // Categories
  var cats = getOrCreate(ss,"Categories");
  if(cats.getLastRow()<=1){
    cats.getRange(1,1,1,2).setValues([["ID","Nom"]]);
    cats.getRange(1,1,1,2).setFontWeight("bold");
    cats.getRange(2,1,5,2).setValues([["c1","Boissons"],["c2","Bières"],["c3","Snacks"],["c4","Sandwichs"],["c5","Divers"]]);
  }
  // Produits (A=ID|B=Nom|C=Prix|D=Cat|E=Emoji|F=Photo|G=Barcode|H=Boisson|I=VenteAl|J=Unite|K=QteBase|L=Stock_s1|M=Stock_s2|N=Stock_s3|O=Stock_s4|P=PrixAchatTTC)
  var prods = getOrCreate(ss,"Produits");
  if(prods.getLastRow()<=1){
    prods.getRange(1,1,1,16).setValues([["ID","Nom","Prix","Categorie","Emoji","Photo","CodeBarres","EstBoisson","VenteAuLitre","Unite","QteBase","Stock_s1","Stock_s2","Stock_s3","Stock_s4","PrixAchatTTC"]]);
    prods.getRange(1,1,1,16).setFontWeight("bold");
    prods.getRange(2,1,12,16).setValues([
      [1,"Coca-Cola 33cl",2,"Boissons","🥤","","5449000000996",true,false,"",1,48,36,24,12,0.55],
      [2,"Eau plate 50cl",1,"Boissons","💧","","3560070976553",true,false,"",1,60,48,36,24,0.25],
      [3,"Jus d'orange",2,"Boissons","🍊","","",true,false,"",1,24,18,12,6,0.7],
      [4,"Café",1.5,"Boissons","☕","","",true,false,"",1,99,99,99,99,0.3],
      [5,"Bière 33cl",3,"Bières","🍺","","5410228091013",true,false,"",1,72,48,36,24,1.1],
      [6,"Bière pression 25cl",2.5,"Bières","🍺","","",true,true,"cl",25,200,150,100,50,0.9],
      [7,"Vin rouge 15cl",2,"Boissons","🍷","","",true,true,"cl",15,150,100,80,50,0.6],
      [8,"Chips",1.5,"Snacks","🥔","","5053990103525",false,false,"",1,30,20,25,10,0.6],
      [9,"Hot-dog",3.5,"Sandwichs","🌭","","",false,false,"",1,12,10,8,5,1.5],
      [10,"Sandwich jambon",4,"Sandwichs","🥪","","",false,false,"",1,10,8,6,4,1.8],
      [11,"Programme",2,"Divers","📋","","",false,false,"",1,50,50,50,50,0.4],
      [12,"Écharpe ANF",10,"Divers","🧣","","",false,false,"",1,20,10,5,5,4]
    ]);
    prods.setFrozenRows(1);
  }
  // Ventes
  var ventes = getOrCreate(ss,"Ventes");
  if(ventes.getLastRow()<=1){
    ventes.getRange(1,1,1,12).setValues([["ID","Date","Heure","SiteID","SiteNom","Total","Paiement","Articles","Membre","NbArticles","Caissier","PartFacture"]]);
    ventes.getRange(1,1,1,12).setFontWeight("bold"); ventes.setFrozenRows(1);
  }
  // Utilisateurs
  var users = getOrCreate(ss,"Utilisateurs");
  if(users.getLastRow()<=1){
    users.getRange(1,1,1,4).setValues([["ID","Nom","PIN","Role"]]);
    users.getRange(1,1,1,4).setFontWeight("bold");
    users.getRange(2,1,1,4).setValues([["u1","Admin","1234","admin"]]);
  }
  return {ok:true, message:"Feuilles initialisées avec succès !"};
}

function getOrCreate(ss,name){
  var s=ss.getSheetByName(name); if(!s)s=ss.insertSheet(name); return s;
}

function getAllData(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  return {ok:true, products:getProductsData(ss), sites:getSitesData(ss),
    categories:getCategoriesData(ss), users:getUsersData(ss), ts:Date.now()};
}

function getProductsData(ss){
  var sh=ss.getSheetByName("Produits"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,16).getValues()
    .filter(r=>r[0]).map(r=>({
      id:r[0],name:r[1],price:+r[2],cat:r[3],emoji:r[4],photo:r[5]||"",barcode:r[6]||"",
      drink:r[7]===true||r[7]==="TRUE"||r[7]==="true",
      sellByVolume:r[8]===true||r[8]==="TRUE"||r[8]==="true",
      unit:r[9]||"",baseQty:+r[10]||1,
      stock:{s1:+r[11]||0,s2:+r[12]||0,s3:+r[13]||0,s4:+r[14]||0},
      costPrice:+r[15]||0
    }));
}
function getSitesData(ss){
  var sh=ss.getSheetByName("Sites"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0])
    .map(r=>({id:r[0],name:r[1],village:r[2],color:r[3]}));
}
function getCategoriesData(ss){
  var sh=ss.getSheetByName("Categories"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,2).getValues().filter(r=>r[0])
    .map(r=>({id:r[0],name:r[1]}));
}
function getUsersData(ss){
  var sh=ss.getSheetByName("Utilisateurs"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0])
    .map(r=>({id:r[0],name:r[1],pin:r[2].toString(),role:r[3]}));
}

function saveProduct(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), p=e.parameter;
  // Colonnes A:K = infos produit (jamais le stock, colonnes L:O)
  var meta=[p.id,p.name,+p.price,p.cat,p.emoji,p.photo||"",p.barcode||"",
    p.drink==="true",p.sellByVolume==="true",p.unit||"",+p.baseQty||1];
  var costPrice=+p.costPrice||0;
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      // Mise à jour : on NE touche PAS aux colonnes de stock (L:O), pour ne jamais écraser
      // une modification de stock faite entre-temps par un autre appareil/utilisateur.
      sh.getRange(i+2,1,1,11).setValues([meta]);
      sh.getRange(i+2,16).setValue(costPrice); // P = Prix d'achat TTC
      return{ok:true,action:"updated"};
    }}
  }
  // Création d'un nouveau produit : stock initialisé à 0 sur tous les sites
  sh.appendRow(meta.concat([0,0,0,0,costPrice]));
  return{ok:true,action:"created"};
}
function deleteProduct(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), id=e.parameter.id;
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}}
  return{ok:false,error:"Non trouvé"};
}
function saveSite(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Sites"), p=e.parameter;
  var row=[p.id,p.name,p.village||"",p.color||"#CC0000"];
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){sh.getRange(i+2,1,1,4).setValues([row]);return{ok:true,action:"updated"};}}}
  sh.appendRow(row); return{ok:true,action:"created"};
}
function deleteSite(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Sites"), id=e.parameter.id;
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}}
  return{ok:false,error:"Non trouvé"};
}
function saveCategory(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Categories"), p=e.parameter;
  var row=[p.id,p.name];
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){sh.getRange(i+2,1,1,2).setValues([row]);return{ok:true,action:"updated"};}}}
  sh.appendRow(row); return{ok:true,action:"created"};
}
function deleteCategory(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Categories"), id=e.parameter.id;
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}}
  return{ok:false,error:"Non trouvé"};
}
function saveUser(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Utilisateurs"), p=e.parameter;
  var row=[p.id,p.name,p.pin,p.role];
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){sh.getRange(i+2,1,1,4).setValues([row]);return{ok:true,action:"updated"};}}}
  sh.appendRow(row); return{ok:true,action:"created"};
}
function deleteUser(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Utilisateurs"), id=e.parameter.id;
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}}
  return{ok:false,error:"Non trouvé"};
}
function updateStock(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), p=e.parameter;
  var col={s1:12,s2:13,s3:14,s4:15}[p.site]; if(!col)return{ok:false,error:"Site invalide"};
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      var cell=sh.getRange(i+2,col);
      cell.setValue(p.mode==="delta"?Math.max(0,+cell.getValue()+(+p.qty)):Math.max(0,+p.qty));
      return{ok:true,newQty:+cell.getValue()};}}}
  return{ok:false,error:"Produit non trouvé"};
}
function transferStock(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), p=e.parameter;
  var colFrom={s1:12,s2:13,s3:14,s4:15}[p.from], colTo={s1:12,s2:13,s3:14,s4:15}[p.to];
  if(!colFrom||!colTo)return{ok:false,error:"Site invalide"};
  var items; try{items=JSON.parse(decodeURIComponent(p.items||"[]"));}catch{return{ok:false,error:"Items invalides"};}
  var ids=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues():[];
  var results=[];
  items.forEach(function(item){
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===item.id.toString()){
      var qty=Math.min(+item.qty,Math.max(0,+sh.getRange(i+2,colFrom).getValue()));
      sh.getRange(i+2,colFrom).setValue(Math.max(0,+sh.getRange(i+2,colFrom).getValue()-qty));
      sh.getRange(i+2,colTo).setValue(+sh.getRange(i+2,colTo).getValue()+qty);
      results.push({id:item.id,transferred:qty}); break;}}
  });
  return{ok:true,results:results};
}
function saveSale(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Ventes"), p=e.parameter;
  var tz=Session.getScriptTimeZone(), now=new Date(), saleId=now.getTime().toString();
  sh.appendRow([saleId,Utilities.formatDate(now,tz,"dd/MM/yyyy"),Utilities.formatDate(now,tz,"HH:mm:ss"),
    p.site,p.siteName,+p.total,p.payment,p.items,p.member||"",+p.nbItems,p.caissier||"",p.splitPart||""]);
  // Décrémenter stock
  var prodSh=ss.getSheetByName("Produits"), col={s1:12,s2:13,s3:14,s4:15}[p.site];
  var items; try{items=JSON.parse(decodeURIComponent(p.items||"[]"));}catch{items=[];}
  if(col&&prodSh.getLastRow()>1){
    var pIds=prodSh.getRange(2,1,prodSh.getLastRow()-1,1).getValues();
    items.forEach(function(item){
      for(var i=0;i<pIds.length;i++){if(pIds[i][0].toString()===item.id.toString()){
        var cell=prodSh.getRange(i+2,col);
        cell.setValue(Math.max(0,+cell.getValue()-item.qty)); break;}}
    });
  }
  return{ok:true,saleId:saleId};
}
function getSales(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Ventes"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1)return{ok:true,sales:[]};
  var data=sh.getRange(2,1,sh.getLastRow()-1,12).getValues();
  var sales=data.filter(function(r){
    if(!r[0])return false;
    if(p.site&&p.site!=="all"&&r[3]!==p.site)return false;
    if(p.date&&r[1]!==p.date)return false;
    return true;
  }).map(function(r){return{id:r[0],date:r[1],time:r[2],siteId:r[3],siteName:r[4],total:r[5],payment:r[6],member:r[8],caissier:r[10]};});
  return{ok:true,sales:sales};
}
