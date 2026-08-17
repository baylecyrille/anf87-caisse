// ═══════════════════════════════════════════════════════════════════
//  ANF 87 — Google Apps Script CAISSE
//  Feuille "Produits" : A=ID|B=Nom|C=Prix|D=Cat|E=Emoji|F=Barcode|G=Boisson|H=Stock_s1|I=Stock_s2|J=Stock_s3|K=Stock_s4
//  Feuille "Sites"    : A=ID|B=Nom|C=Village|D=Couleur
//  Feuille "Ventes"   : A=ID|B=Date|C=Heure|D=SiteID|E=SiteNom|F=Total|G=Paiement|H=Articles|I=Membre|J=NbArticles
// ═══════════════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e.parameter.action || "getAll");
  var result;
  try {
    switch(action) {
      case "getAll":         result = getAllData();        break;
      case "saveProduct":    result = saveProduct(e);     break;
      case "deleteProduct":  result = deleteProduct(e);   break;
      case "saveSite":       result = saveSite(e);        break;
      case "deleteSite":     result = deleteSite(e);      break;
      case "saveSale":       result = saveSale(e);        break;
      case "updateStock":    result = updateStock(e);     break;
      case "getSales":       result = getSales(e);        break;
      case "initSheets":     result = initSheets();       break;
      default: result = {ok:false, error:"Action inconnue: "+action};
    }
  } catch(err) {
    result = {ok:false, error:err.toString()};
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
//  INITIALISATION DES FEUILLES (1ère utilisation)
// ─────────────────────────────────────────────
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Feuille Sites
  var sites = getOrCreateSheet(ss, "Sites");
  if (sites.getLastRow() <= 1) {
    sites.getRange(1,1,1,4).setValues([["ID","Nom","Village","Couleur"]]);
    sites.getRange(1,1,1,4).setFontWeight("bold");
    sites.getRange(2,1,4,4).setValues([
      ["s1","Buvette Nord","Village 1","#CC0000"],
      ["s2","Buvette Est","Village 2","#1565C0"],
      ["s3","Buvette Sud","Village 3","#2E7D32"],
      ["s4","Buvette Ouest","Village 4","#E65100"],
    ]);
  }

  // Feuille Produits
  var prods = getOrCreateSheet(ss, "Produits");
  if (prods.getLastRow() <= 1) {
    prods.getRange(1,1,1,11).setValues([["ID","Nom","Prix","Categorie","Emoji","CodeBarres","EstBoisson","Stock_s1","Stock_s2","Stock_s3","Stock_s4"]]);
    prods.getRange(1,1,1,11).setFontWeight("bold");
    prods.getRange(2,1,14,11).setValues([
      [1,"Coca-Cola 33cl",2.00,"Boissons","🥤","5449000000996",true,48,36,24,12],
      [2,"Eau plate 50cl",1.00,"Boissons","💧","3560070976553",true,60,48,36,24],
      [3,"Jus d'orange",2.00,"Boissons","🍊","3228882012350",true,24,18,12,6],
      [4,"Limonade",2.00,"Boissons","🍋","",true,24,12,18,8],
      [5,"Café",1.50,"Boissons","☕","",true,99,99,99,99],
      [6,"Bière 33cl",3.00,"Bières","🍺","5410228091013",true,72,48,36,24],
      [7,"Bière sans alcool",2.50,"Bières","🍻","",true,24,24,12,12],
      [8,"Chips",1.50,"Snacks","🥔","5053990103525",false,30,20,25,10],
      [9,"Cacahuètes",1.50,"Snacks","🥜","",false,20,15,18,8],
      [10,"Barre chocolatée",1.50,"Snacks","🍫","",false,24,20,16,8],
      [11,"Sandwich jambon",4.00,"Sandwichs","🥪","",false,10,8,6,4],
      [12,"Hot-dog",3.50,"Sandwichs","🌭","",false,12,10,8,5],
      [13,"Programme match",2.00,"Divers","📋","",false,50,50,50,50],
      [14,"Écharpe ANF 87",10.00,"Divers","🧣","",false,20,10,5,5],
    ]);
  }

  // Feuille Ventes
  var ventes = getOrCreateSheet(ss, "Ventes");
  if (ventes.getLastRow() <= 1) {
    ventes.getRange(1,1,1,10).setValues([["ID","Date","Heure","SiteID","SiteNom","Total","Paiement","Articles","Membre","NbArticles"]]);
    ventes.getRange(1,1,1,10).setFontWeight("bold");
    ventes.setFrozenRows(1);
  }

  return {ok:true, message:"Feuilles initialisées avec succès !"};
}

function getOrCreateSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// ─────────────────────────────────────────────
//  RÉCUPÉRER TOUTES LES DONNÉES
// ─────────────────────────────────────────────
function getAllData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  return {
    ok:       true,
    products: getProductsData(ss),
    sites:    getSitesData(ss),
    ts:       new Date().getTime()
  };
}

function getProductsData(ss) {
  var sh   = ss.getSheetByName("Produits");
  if (!sh || sh.getLastRow() <= 1) return [];
  var data = sh.getRange(2, 1, sh.getLastRow()-1, 11).getValues();
  return data.filter(function(r){ return r[0]; }).map(function(r) {
    return {
      id:      r[0], name: r[1], price: Number(r[2]),
      cat:     r[3], emoji: r[4], barcode: r[5]||"",
      drink:   r[6]===true||r[6]==="TRUE"||r[6]==="true",
      stock:   {s1:Number(r[7]||0), s2:Number(r[8]||0), s3:Number(r[9]||0), s4:Number(r[10]||0)}
    };
  });
}

function getSitesData(ss) {
  var sh = ss.getSheetByName("Sites");
  if (!sh || sh.getLastRow() <= 1) return [];
  var data = sh.getRange(2, 1, sh.getLastRow()-1, 4).getValues();
  return data.filter(function(r){ return r[0]; }).map(function(r) {
    return {id:r[0], name:r[1], village:r[2], color:r[3]};
  });
}

// ─────────────────────────────────────────────
//  PRODUITS — Créer / Modifier
// ─────────────────────────────────────────────
function saveProduct(e) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName("Produits");
  var id   = e.parameter.id;
  var price= parseFloat(e.parameter.price||"0");
  var drink= (e.parameter.drink==="true"||e.parameter.drink===true);
  var row  = [
    id, e.parameter.name||"", price,
    e.parameter.cat||"Divers", e.parameter.emoji||"🏷", e.parameter.barcode||"", drink,
    parseInt(e.parameter.stock_s1||"0"), parseInt(e.parameter.stock_s2||"0"),
    parseInt(e.parameter.stock_s3||"0"), parseInt(e.parameter.stock_s4||"0")
  ];

  // Chercher si le produit existe déjà
  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (var i=0; i<ids.length; i++) {
      if (ids[i][0].toString() === id.toString()) {
        sh.getRange(i+2,1,1,11).setValues([row]);
        return {ok:true, action:"updated", id:id};
      }
    }
  }
  // Nouveau produit
  sh.appendRow(row);
  return {ok:true, action:"created", id:id};
}

// ─────────────────────────────────────────────
//  PRODUITS — Supprimer
// ─────────────────────────────────────────────
function deleteProduct(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Produits");
  var id = e.parameter.id;
  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (var i=0; i<ids.length; i++) {
      if (ids[i][0].toString() === id.toString()) {
        sh.deleteRow(i+2);
        return {ok:true, action:"deleted", id:id};
      }
    }
  }
  return {ok:false, error:"Produit non trouvé"};
}

// ─────────────────────────────────────────────
//  STOCK — Mettre à jour
// ─────────────────────────────────────────────
function updateStock(e) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sh    = ss.getSheetByName("Produits");
  var id    = e.parameter.id;
  var site  = e.parameter.site;  // "s1" à "s4"
  var qty   = parseInt(e.parameter.qty||"0");
  var mode  = e.parameter.mode||"set"; // "set" ou "delta"
  var col   = {s1:8, s2:9, s3:10, s4:11}[site];
  if (!col) return {ok:false, error:"Site invalide"};

  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (var i=0; i<ids.length; i++) {
      if (ids[i][0].toString()===id.toString()) {
        var cell = sh.getRange(i+2, col);
        var newVal = mode==="delta" ? Math.max(0, Number(cell.getValue())+qty) : Math.max(0,qty);
        cell.setValue(newVal);
        return {ok:true, id:id, site:site, newQty:newVal};
      }
    }
  }
  return {ok:false, error:"Produit non trouvé"};
}

// ─────────────────────────────────────────────
//  SITES — Créer / Modifier
// ─────────────────────────────────────────────
function saveSite(e) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName("Sites");
  var id  = e.parameter.id;
  var row = [id, e.parameter.name||"", e.parameter.village||"", e.parameter.color||"#CC0000"];
  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (var i=0; i<ids.length; i++) {
      if (ids[i][0].toString()===id.toString()) {
        sh.getRange(i+2,1,1,4).setValues([row]);
        return {ok:true, action:"updated"};
      }
    }
  }
  sh.appendRow(row);
  return {ok:true, action:"created"};
}

function deleteSite(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Sites");
  var id = e.parameter.id;
  if (sh.getLastRow() > 1) {
    var ids = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (var i=0; i<ids.length; i++) {
      if (ids[i][0].toString()===id.toString()) {
        sh.deleteRow(i+2);
        return {ok:true};
      }
    }
  }
  return {ok:false, error:"Site non trouvé"};
}

// ─────────────────────────────────────────────
//  VENTES — Enregistrer + décrémenter le stock
// ─────────────────────────────────────────────
function saveSale(e) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sh     = ss.getSheetByName("Ventes");
  var tz     = Session.getScriptTimeZone();
  var now    = new Date();
  var saleId = now.getTime().toString();

  var siteId  = e.parameter.site||"s1";
  var siteNom = e.parameter.siteName||"";
  var total   = parseFloat(e.parameter.total||"0");
  var payment = e.parameter.payment||"";
  var items   = e.parameter.items||"[]";
  var membre  = e.parameter.member||"";
  var nbItems = parseInt(e.parameter.nbItems||"0");

  // Enregistrer la vente
  sh.appendRow([
    saleId,
    Utilities.formatDate(now, tz, "dd/MM/yyyy"),
    Utilities.formatDate(now, tz, "HH:mm:ss"),
    siteId, siteNom, total, payment, items, membre, nbItems
  ]);

  // Décrémenter le stock
  var prodSh  = ss.getSheetByName("Produits");
  var col     = {s1:8, s2:9, s3:10, s4:11}[siteId];
  var itemArr;
  try { itemArr = JSON.parse(decodeURIComponent(items)); } catch(ex) { itemArr = []; }
  
  if (col && prodSh.getLastRow()>1 && itemArr.length>0) {
    var prodIds = prodSh.getRange(2,1,prodSh.getLastRow()-1,1).getValues();
    itemArr.forEach(function(item) {
      for (var i=0; i<prodIds.length; i++) {
        if (prodIds[i][0].toString()===item.id.toString()) {
          var cell = prodSh.getRange(i+2, col);
          cell.setValue(Math.max(0, Number(cell.getValue()) - item.qty));
          break;
        }
      }
    });
  }

  return {ok:true, saleId:saleId};
}

// ─────────────────────────────────────────────
//  VENTES — Lire (rapports)
// ─────────────────────────────────────────────
function getSales(e) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName("Ventes");
  var site = e.parameter.site||"";
  var date = e.parameter.date||"";
  if (!sh || sh.getLastRow()<=1) return {ok:true, sales:[]};
  var data = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
  var sales = data.filter(function(r){
    if (!r[0]) return false;
    if (site && r[3]!==site) return false;
    if (date && r[1]!==date) return false;
    return true;
  }).map(function(r){
    return {id:r[0],date:r[1],time:r[2],siteId:r[3],siteName:r[4],total:r[5],payment:r[6],member:r[8]};
  });
  return {ok:true, sales:sales};
}
