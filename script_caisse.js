// ═══════════════════════════════════════════════════════
//  ANF 87 — Google Apps Script CAISSE v2
//  Sheets: Produits | Sites | Categories | Ventes | Utilisateurs
// ═══════════════════════════════════════════════════════

// Actions qui ÉCRIVENT dans le classeur : protégées par un verrou pour éviter
// les conflits quand 2-3 utilisateurs utilisent la caisse en même temps.
var WRITE_ACTIONS = ["saveProduct","deleteProduct","saveSite","deleteSite","saveCategory",
  "deleteCategory","saveUser","deleteUser","saveSale","updateStock","transferStock","uploadPhotoChunk",
  "openContainer","closeContainer","savePortion","deletePortion","uploadPortionPhotoChunk",
  "saveCombo","deleteCombo","uploadComboPhotoChunk","uploadComboItemsChunk","uploadSaleItemsChunk","saveAccompaniment","deleteAccompaniment",
  "saveProductOrder","addToReserve","adjustContainer","setReserveCount","saveConfig","updateSale","saveCameraPref","deleteSale","adjustDeposit","logHappyHour","adjustReadyCombo",
  "savePlat","deletePlat","uploadPlatPhotoChunk","uploadPlatItemsChunk","adjustReadyPlat","migrateToPlats"];
  // initSheets n'est pas verrouillé : c'est une action ponctuelle de 1ère installation,
  // pas de risque de conflit multi-utilisateurs à ce moment-là.

// doPost existe pour compatibilité mais n'est plus utilisé par l'appli : sur ce
// déploiement, les requêtes POST vers l'URL Apps Script renvoient une page de
// connexion Google (HTML) au lieu du JSON attendu. L'appli envoie donc tout en GET,
// et découpe les photos produit (trop volumineuses pour un seul GET) en plusieurs
// petits morceaux envoyés séparément — voir uploadPhotoChunk / saveProduct.
function doPost(e) { return doGet(e); }

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
      case "uploadPhotoChunk":result = uploadPhotoChunk(e); break;
      case "openContainer":   result = openContainer(e); break;
      case "closeContainer":  result = closeContainer(e); break;
      case "adjustContainer": result = adjustContainer(e); break;
      case "setReserveCount":  result = setReserveCount(e); break;
      case "saveConfig":       result = saveConfig(e); break;
      case "logHappyHour":     result = logHappyHour(e); break;
      case "updateSale":       result = updateSale(e); break;
      case "deleteSale":       result = deleteSale(e); break;
      case "savePortion":     result = savePortion(e); break;
      case "deletePortion":   result = deletePortion(e); break;
      case "uploadPortionPhotoChunk": result = uploadPortionPhotoChunk(e); break;
      case "saveCombo":       result = saveCombo(e); break;
      case "deleteCombo":     result = deleteCombo(e); break;
      case "uploadComboPhotoChunk": result = uploadComboPhotoChunk(e); break;
      case "uploadComboItemsChunk": result = uploadComboItemsChunk(e); break;
      case "uploadSaleItemsChunk": result = uploadSaleItemsChunk(e); break;
      case "saveAccompaniment":   result = saveAccompaniment(e); break;
      case "deleteAccompaniment": result = deleteAccompaniment(e); break;
      case "saveProductOrder":    result = saveProductOrder(e); break;
      case "saveCameraPref":      result = saveCameraPref(e); break;
      case "addToReserve":        result = addToReserve(e); break;
      case "adjustDeposit":       result = adjustDeposit(e); break;
      case "adjustReadyCombo":    result = adjustReadyCombo(e); break;
      case "savePlat":            result = savePlat(e); break;
      case "deletePlat":          result = deletePlat(e); break;
      case "uploadPlatPhotoChunk": result = uploadPlatPhotoChunk(e); break;
      case "uploadPlatItemsChunk": result = uploadPlatItemsChunk(e); break;
      case "adjustReadyPlat":     result = adjustReadyPlat(e); break;
      case "migrateToPlats":      result = migrateToPlats(); break;
      default: result = {ok:false, error:"Action inconnue: "+action};
    }
  } catch(err) { result = {ok:false, error:err.toString()}; }
  finally {
    // On force l'écriture à être immédiatement visible côté serveur AVANT de relâcher
    // le verrou et de répondre — sans ça, une lecture (sync automatique) qui arrive
    // juste après peut retomber sur une version pas encore "commitée" de la feuille et
    // donner l'impression que la modification vient d'être effacée.
    if (WRITE_ACTIONS.indexOf(action) !== -1) { try { SpreadsheetApp.flush(); } catch(e2) {} }
    if (lock) lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Sites
  var sites = getOrCreate(ss,"Sites");
  if(sites.getLastRow()<=1){
    sites.getRange(1,1,1,6).setValues([["ID","Nom","Village","Couleur","Actif","PointVente"]]);
    sites.getRange(1,1,1,6).setFontWeight("bold");
    sites.getRange(2,1,4,6).setValues([
      ["s1","Buvette Nord","Village 1","#CC0000",true,true],
      ["s2","Buvette Est","Village 2","#1565C0",true,true],
      ["s3","Buvette Sud","Village 3","#2E7D32",true,true],
      ["s4","Buvette Ouest","Village 4","#E65100",true,true]
    ]);
  }
  // Categories
  var cats = getOrCreate(ss,"Categories");
  if(cats.getLastRow()<=1){
    cats.getRange(1,1,1,2).setValues([["ID","Nom"]]);
    cats.getRange(1,1,1,2).setFontWeight("bold");
    cats.getRange(2,1,5,2).setValues([["c1","Boissons"],["c2","Bières"],["c3","Snacks"],["c4","Sandwichs"],["c5","Divers"]]);
  }
  // Produits (A=ID|B=Nom|C=Prix|D=Cat|E=Emoji|F=Photo|G=Barcode|H=Boisson|I=VenteAl|J=Unite|K=QteBase|L=Stock_s1|M=Stock_s2|N=Stock_s3|O=Stock_s4|P=PrixAchatTTC|Q=Paliers)
  var prods = getOrCreate(ss,"Produits");
  if(prods.getLastRow()<=1){
    prods.getRange(1,1,1,17).setValues([["ID","Nom","Prix","Categorie","Emoji","Photo","CodeBarres","EstBoisson","VenteAuLitre","Unite","QteBase","Stock_s1","Stock_s2","Stock_s3","Stock_s4","PrixAchatTTC","Paliers"]]);
    prods.getRange(1,1,1,17).setFontWeight("bold");
    prods.getRange(2,1,12,17).setValues([
      [1,"Coca-Cola 33cl",2,"Boissons","🥤","","5449000000996",true,false,"",1,48,36,24,12,0.55,""],
      [2,"Eau plate 50cl",1,"Boissons","💧","","3560070976553",true,false,"",1,60,48,36,24,0.25,""],
      [3,"Jus d'orange",2,"Boissons","🍊","","",true,false,"",1,24,18,12,6,0.7,""],
      [4,"Café",1.5,"Boissons","☕","","",true,false,"",1,99,99,99,99,0.3,""],
      [5,"Bière 33cl",3,"Bières","🍺","","5410228091013",true,false,"",1,72,48,36,24,1.1,""],
      [6,"Bière pression 25cl",2.5,"Bières","🍺","","",true,true,"cl",25,200,150,100,50,0.9,"25:2.5,50:5,100:9"],
      [7,"Vin rouge 15cl",2,"Boissons","🍷","","",true,true,"cl",15,150,100,80,50,0.6,"15:2,25:3.2"],
      [8,"Chips",1.5,"Snacks","🥔","","5053990103525",false,false,"",1,30,20,25,10,0.6,""],
      [9,"Hot-dog",3.5,"Sandwichs","🌭","","",false,false,"",1,12,10,8,5,1.5,""],
      [10,"Sandwich jambon",4,"Sandwichs","🥪","","",false,false,"",1,10,8,6,4,1.8,""],
      [11,"Programme",2,"Divers","📋","","",false,false,"",1,50,50,50,50,0.4,""],
      [12,"Écharpe ANF",10,"Divers","🧣","","",false,false,"",1,20,10,5,5,4,""]
    ]);
    prods.setFrozenRows(1);
  }
  // Ventes
  var ventes = getOrCreate(ss,"Ventes");
  if(ventes.getLastRow()<=1){
    ventes.getRange(1,1,1,15).setValues([["ID","Date","Heure","SiteID","SiteNom","Total","Paiement","Articles","Membre","NbArticles","Caissier","PartFacture","MenusPretsConsommes","MenusVendus","PlatsPretsConsommes"]]);
    ventes.getRange(1,1,1,15).setFontWeight("bold"); ventes.setFrozenRows(1);
  }
  // Utilisateurs
  var users = getOrCreate(ss,"Utilisateurs");
  if(users.getLastRow()<=1){
    users.getRange(1,1,1,6).setValues([["ID","Nom","PIN","Role","OrdreProduits","CameraPref"]]);
    users.getRange(1,1,1,6).setFontWeight("bold");
    users.getRange(2,1,1,4).setValues([["u1","Admin","1234","admin"]]);
  }
  // Contenants ouverts (fûts/bouteilles/cubis entamés, un par site+produit)
  var cont = getOrCreate(ss,"Contenants");
  if(cont.getLastRow()<=1){
    cont.getRange(1,1,1,6).setValues([["SiteID","ProduitID","Contenant","TailleCl","RestantCl","DateOuverture"]]);
    cont.getRange(1,1,1,6).setFontWeight("bold");
    cont.setFrozenRows(1);
  }
  // Réserve précise de contenants (par type) reçus via Entrée rapide, pour savoir
  // exactement combien de fûts/bouteilles/cubis de CHAQUE taille restent en stock.
  var res = getOrCreate(ss,"ReserveContenants");
  if(res.getLastRow()<=1){
    res.getRange(1,1,1,5).setValues([["SiteID","ProduitID","Contenant","TailleCl","Nombre"]]);
    res.getRange(1,1,1,5).setFontWeight("bold");
    res.setFrozenRows(1);
  }
  // Réglages globaux du club (clé/valeur) — clé Affiliate SumUp, App ID SumUp...
  var cfg = getOrCreate(ss,"Config");
  if(cfg.getLastRow()<=1){
    cfg.getRange(1,1,1,2).setValues([["Clé","Valeur"]]);
    cfg.getRange(1,1,1,2).setFontWeight("bold");
    cfg.setFrozenRows(1);
  }
  // Consignes actuellement chez les clients (payées, pas encore rendues) — par
  // produit "consigne" et par site.
  var dep = getOrCreate(ss,"ConsignesDehors");
  if(dep.getLastRow()<=1){
    dep.getRange(1,1,1,3).setValues([["SiteID","ProduitID","Nombre"]]);
    dep.getRange(1,1,1,3).setFontWeight("bold");
    dep.setFrozenRows(1);
  }
  // Portions de vente (paliers) : demis/pichets/tailles de gobelet... avec leur
  // propre photo, liées à un produit "liquide" (celui vendu au volume).
  var port = getOrCreate(ss,"Portions");
  if(port.getLastRow()<=1){
    port.getRange(1,1,1,6).setValues([["ID","ProduitID","Nom","TailleCl","Prix","Photo"]]);
    port.getRange(1,1,1,6).setFontWeight("bold");
    port.setFrozenRows(1);
  }
  // Combos/Menus : regroupent plusieurs produits DIFFÉRENTS en une seule vente
  // (ex: Croque + Frites + Boisson). Contenu = liste "produitId:qté,produitId:qté".
  var combo = getOrCreate(ss,"Combos");
  if(combo.getLastRow()<=1){
    combo.getRange(1,1,1,5).setValues([["ID","Nom","Prix","Contenu","Photo"]]);
    combo.getRange(1,1,1,5).setFontWeight("bold");
    combo.setFrozenRows(1);
  }
  // Plats : un niveau intermédiaire entre Produit et Menu (ex: "Croque Monsieur" =
  // un Plat composé de pain+jambon+fromage). Un Menu peut ensuite inclure un Plat
  // comme ingrédient (ex: "Croque Frites" = Plat "Croque Monsieur" + frites +
  // boisson), et un Plat peut aussi se vendre seul directement en caisse. Mêmes
  // colonnes que Combos, mais son "Contenu" ne référence jamais un autre Plat/Menu
  // (uniquement des produits bruts) — c'est ce qui évite l'imbrication à profondeur
  // illimitée qui causait des bugs difficiles à diagnostiquer.
  var plat = getOrCreate(ss,"Plats");
  if(plat.getLastRow()<=1){
    plat.getRange(1,1,1,5).setValues([["ID","Nom","Prix","Contenu","Photo"]]);
    plat.getRange(1,1,1,5).setFontWeight("bold");
    plat.setFrozenRows(1);
  }
  // Accompagnements automatiques : quand on vend le produit ProduitID, on décompte
  // AUSSI, en plus et discrètement (pas de ligne de panier, pas d'impact sur le
  // prix), Qté unités du produit ProduitAssocieID (ex: café -> gobelet, sucre,
  // touillette). Un même produit peut avoir plusieurs lignes d'accompagnement.
  var acc = getOrCreate(ss,"Accompagnements");
  if(acc.getLastRow()<=1){
    acc.getRange(1,1,1,4).setValues([["ID","ProduitID","ProduitAssocieID","Qte"]]);
    acc.getRange(1,1,1,4).setFontWeight("bold");
    acc.setFrozenRows(1);
  }
  return {ok:true, message:"Feuilles initialisées avec succès !"};
}

function safeDecodeItems(raw){
  raw = raw || "[]";
  try { return decodeURIComponent(raw); } catch(e) { return raw; }
}
function getOrCreate(ss,name){
  var s=ss.getSheetByName(name); if(!s)s=ss.insertSheet(name); return s;
}

function getAllData(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  return {ok:true, products:getProductsData(ss), sites:getSitesData(ss),
    categories:getCategoriesData(ss), users:getUsersData(ss), containers:getContainersData(ss),
    portions:getPortionsData(ss), combos:getCombosData(ss), accompaniments:getAccompanimentsData(ss),
    reserve:getReserveData(ss), deposits:getDepositsData(ss), readyCombos:getReadyCombosData(ss),
    plats:getPlatsData(ss), readyPlats:getReadyPlatsData(ss), config:getConfig(ss), ts:Date.now()};
}

// Réglages globaux du club (clé/valeur) — ex: clé Affiliate SumUp, App ID SumUp.
// Contrairement à l'ordre des produits (propre à chaque utilisateur), ces réglages
// sont partagés par tout le monde : une seule config SumUp pour tout le club.
function getConfig(ss){
  var sh=ss.getSheetByName("Config"); if(!sh||sh.getLastRow()<=1)return {};
  var data=sh.getRange(2,1,sh.getLastRow()-1,2).getValues(), cfg={};
  data.forEach(function(r){ if(r[0])cfg[r[0].toString()]=r[1]||""; });
  return cfg;
}
function saveConfig(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Config"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,2).setValues([["Clé","Valeur"]]);
    sh.getRange(1,1,1,2).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  if(!p.key)return{ok:false,error:"Clé manquante"};
  if(sh.getLastRow()>1){
    var keys=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<keys.length;i++){if(keys[i][0].toString()===p.key.toString()){
      sh.getRange(i+2,2).setValue(p.value||"");
      return{ok:true};
    }}
  }
  sh.appendRow([p.key,p.value||""]);
  return{ok:true};
}
// Historique des activations/désactivations du Happy Hour — une ligne par bascule,
// pour garder une trace de quand ça a été activé/désactivé et par qui.
function logHappyHour(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"HappyHourLog"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,4).setValues([["Date","Heure","Action","Utilisateur"]]);
    sh.getRange(1,1,1,4).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var now=new Date(), tz=Session.getScriptTimeZone();
  var dateStr=Utilities.formatDate(now,tz,"dd/MM/yyyy"), heureStr=Utilities.formatDate(now,tz,"HH:mm:ss");
  sh.appendRow([dateStr,heureStr,p.action2==="on"?"Activé":"Désactivé",p.user||""]);
  sh.getRange(sh.getLastRow(),1,1,2).setNumberFormat("@"); // Date/Heure en texte, même raison que pour les ventes
  return{ok:true};
}

function getReserveData(ss){
  var sh=ss.getSheetByName("ReserveContenants"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,5).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({site:r[0],productId:r[1],label:r[2],size:+r[3]||0,count:+r[4]||0}));
}
// Combien d'unités d'un produit "consigne" sont actuellement chez les clients (payées
// mais pas encore rendues) — incrémenté à la vente d'une consigne, décrémenté à la
// vente d'un "retour de consigne" (produit lié via son champ ReturnFor).
function getDepositsData(ss){
  var sh=ss.getSheetByName("ConsignesDehors"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,3).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({site:r[0],productId:r[1],count:+r[2]||0}));
}
function adjustDeposit(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"ConsignesDehors"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,3).setValues([["SiteID","ProduitID","Nombre"]]);
    sh.getRange(1,1,1,3).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var site=p.site, id=p.id, delta=+p.delta||0;
  if(!site||!id||delta===0)return{ok:false,error:"Paramètres manquants"};
  if(sh.getLastRow()>1){
    var data=sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
    for(var i=0;i<data.length;i++){
      if(data[i][0].toString()===site.toString()&&data[i][1].toString()===id.toString()){
        var newCount=Math.max(0,(+data[i][2]||0)+delta);
        sh.getRange(i+2,3).setValue(newCount);
        return{ok:true,count:newCount};
      }
    }
  }
  sh.appendRow([site,id,Math.max(0,delta)]);
  return{ok:true,count:Math.max(0,delta)};
}
// Portions de MENU déjà préparées d'avance (ex: 20 Croque Monsieur montés avant le
// coup d'envoi), par site — décrémenté directement à la vente d'un menu tant qu'il
// en reste, sans re-décompter les ingrédients (déjà fait à la préparation).
function getReadyCombosData(ss){
  var sh=ss.getSheetByName("MenusPrets"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,3).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({site:r[0],comboId:r[1],count:+r[2]||0}));
}
function adjustReadyCombo(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"MenusPrets"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,3).setValues([["SiteID","MenuID","Nombre"]]);
    sh.getRange(1,1,1,3).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var site=p.site, id=p.id, delta=+p.delta||0;
  if(!site||!id||delta===0)return{ok:false,error:"Paramètres manquants"};
  if(sh.getLastRow()>1){
    var data=sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
    for(var i=0;i<data.length;i++){
      if(data[i][0].toString()===site.toString()&&data[i][1].toString()===id.toString()){
        var newCount=Math.max(0,(+data[i][2]||0)+delta);
        sh.getRange(i+2,3).setValue(newCount);
        return{ok:true,count:newCount};
      }
    }
  }
  sh.appendRow([site,id,Math.max(0,delta)]);
  return{ok:true,count:Math.max(0,delta)};
}
// Portions de PLAT déjà préparées d'avance (ex: 20 Croque Monsieur montés avant le
// coup d'envoi, vendus seuls OU utilisés comme ingrédient d'un menu) — même
// mécanisme que MenusPrets/adjustReadyCombo, sur sa propre feuille "PlatsPrets".
function getReadyPlatsData(ss){
  var sh=ss.getSheetByName("PlatsPrets"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,3).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({site:r[0],platId:r[1],count:+r[2]||0}));
}
function adjustReadyPlat(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"PlatsPrets"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,3).setValues([["SiteID","PlatID","Nombre"]]);
    sh.getRange(1,1,1,3).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var site=p.site, id=p.id, delta=+p.delta||0;
  if(!site||!id||delta===0)return{ok:false,error:"Paramètres manquants"};
  if(sh.getLastRow()>1){
    var data=sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
    for(var i=0;i<data.length;i++){
      if(data[i][0].toString()===site.toString()&&data[i][1].toString()===id.toString()){
        var newCount=Math.max(0,(+data[i][2]||0)+delta);
        sh.getRange(i+2,3).setValue(newCount);
        return{ok:true,count:newCount};
      }
    }
  }
  sh.appendRow([site,id,Math.max(0,delta)]);
  return{ok:true,count:Math.max(0,delta)};
}
// Enregistre l'ajout de N contenants d'un type précis (utilisé depuis Entrée rapide),
// pour savoir EXACTEMENT combien de fûts/bouteilles/cubis de CHAQUE taille sont en
// réserve — plus fiable qu'un calcul dérivé du stock total en cl si on mélange
// plusieurs tailles pour un même produit.
function addToReserve(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"ReserveContenants"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,5).setValues([["SiteID","ProduitID","Contenant","TailleCl","Nombre"]]);
    sh.getRange(1,1,1,5).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var site=p.site, id=p.id, label=p.label||"", size=+p.size||0, count=+p.count||0;
  if(!site||!id||size<=0||count===0)return{ok:false,error:"Paramètres manquants"};
  if(sh.getLastRow()>1){
    var data=sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
    for(var i=0;i<data.length;i++){
      if(data[i][0].toString()===site.toString()&&data[i][1].toString()===id.toString()
        &&data[i][2].toString()===label&&(+data[i][3]||0)===size){
        var newCount=Math.max(0,(+data[i][4]||0)+count);
        sh.getRange(i+2,5).setValue(newCount);
        return{ok:true,count:newCount};
      }
    }
  }
  sh.appendRow([site,id,label,size,Math.max(0,count)]);
  return{ok:true,count:Math.max(0,count)};
}
// Fixe directement le nombre de contenants en réserve (correction manuelle depuis
// l'écran Stock), contrairement à addToReserve qui ADDITIONNE (utilisé par Entrée
// rapide). Crée la ligne si elle n'existe pas encore.
function setReserveCount(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"ReserveContenants"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,5).setValues([["SiteID","ProduitID","Contenant","TailleCl","Nombre"]]);
    sh.getRange(1,1,1,5).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var site=p.site, id=p.id, label=p.label||"", size=+p.size||0, count=Math.max(0,+p.count||0);
  if(!site||!id||size<=0)return{ok:false,error:"Paramètres manquants"};
  if(sh.getLastRow()>1){
    var data=sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
    for(var i=0;i<data.length;i++){
      if(data[i][0].toString()===site.toString()&&data[i][1].toString()===id.toString()
        &&data[i][2].toString()===label&&(+data[i][3]||0)===size){
        sh.getRange(i+2,5).setValue(count);
        return{ok:true,count:count};
      }
    }
  }
  sh.appendRow([site,id,label,size,count]);
  return{ok:true,count:count};
}

function getAccompanimentsData(ss){
  var sh=ss.getSheetByName("Accompagnements"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,5).getValues().filter(r=>r[0]&&r[1]&&r[2])
    .map(r=>({id:r[0],productId:r[1],assocId:r[2],qty:+r[3]||1,portionId:(r[4]||"").toString()}));
}
function saveAccompaniment(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Accompagnements"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,5).setValues([["ID","ProduitID","ProduitAssocieID","Qte","PortionID"]]);
    sh.getRange(1,1,1,5).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var row=[p.id,p.productId,p.assocId,+p.qty||1,p.portionId||""];
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      sh.getRange(i+2,1,1,5).setValues([row]);
      return{ok:true,action:"updated"};
    }}
  }
  sh.appendRow(row);
  return{ok:true,action:"created"};
}
function deleteAccompaniment(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Accompagnements"), id=e.parameter.id;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Non trouvé"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}
  return{ok:false,error:"Non trouvé"};
}

function getPortionsData(ss){
  var sh=ss.getSheetByName("Portions"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,7).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({id:r[0],productId:r[1],name:r[2],size:+r[3]||0,price:+r[4]||0,photo:r[5]||"",happyPrice:+r[6]||0}));
}

// Crée ou met à jour une portion de vente (demi, pichet, taille de gobelet...).
function savePortion(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Portions"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,7).setValues([["ID","ProduitID","Nom","TailleCl","Prix","Photo","PrixHappyHour"]]);
    sh.getRange(1,1,1,7).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var photo=(p.photo||"").toString().trim();
  if(photo==="__CHUNKED__"){
    var total=+p.photoChunks||0;
    var cache=CacheService.getScriptCache(), parts=[];
    for(var c=0;c<total;c++){
      var part=cache.get("portionphoto_"+p.id+"_"+c);
      if(part===null)return{ok:false,error:"Photo incomplète (morceau "+(c+1)+"/"+total+" manquant ou expiré), réessayez."};
      parts.push(part);
    }
    photo=parts.join("");
    for(var c2=0;c2<total;c2++)cache.remove("portionphoto_"+p.id+"_"+c2);
  }
  var row=[p.id,p.productId,p.name||"",+p.size||0,+p.price||0,photo,+p.happyPrice||0];
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      if(photo==="__KEEP__")row[5]=sh.getRange(i+2,6).getValue(); // photo inchangée, non renvoyée
      if(row[5]==="__KEEP__")row[5]=sh.getRange(i+2,6).getValue(); // garde-fou : jamais la valeur littérale
      sh.getRange(i+2,1,1,7).setValues([row]);
      return{ok:true,action:"updated"};
    }}
  }
  if(photo==="__KEEP__")row[5]=""; // création : rien à garder
  sh.appendRow(row);
  return{ok:true,action:"created"};
}
function deletePortion(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Portions"), id=e.parameter.id;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Non trouvé"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}
  return{ok:false,error:"Non trouvé"};
}
// Même mécanisme que uploadPhotoChunk (voir plus bas), sous une clé de cache dédiée
// pour ne jamais entrer en collision avec les morceaux de photo produit.
function uploadPortionPhotoChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("portionphoto_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}

function getCombosData(ss){
  var sh=ss.getSheetByName("Combos"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,5).getValues().filter(r=>r[0]).map(r=>{
    var items=(r[3]||"").toString().split(",").map(function(pair){
      var parts=pair.split(":");
      // 4e partie optionnelle : autres produits interchangeables pour cet
      // ingrédient (séparés par "|"), pour pouvoir changer facilement lequel est
      // utilisé cette semaine (ex: jambon 140g vs jambon supérieur 240g) sans
      // rouvrir toute la configuration du menu.
      var alternates=(parts[3]||"").split("|").map(function(a){return a.trim();}).filter(Boolean);
      // 5e partie optionnelle : "combo" si cette ligne référence un AUTRE menu
      // (imbrication, ex: "Croque Frites" qui inclut le menu "Croque Monsieur")
      // plutôt qu'un produit — dans ce cas le 1er champ contient l'ID du menu
      // référencé, pas un ID produit.
      var type=(parts[4]||"").trim();
      var refId=(parts[0]||"").trim();
      if(type==="combo"){
        return{type:"combo",comboId:refId,qty:+parts[1]||1};
      }
      return{type:"product",productId:refId,qty:+parts[1]||1,portionId:(parts[2]||"").trim(),alternates:alternates};
    }).filter(function(it){return it.type==="combo"?it.comboId:it.productId;});
    return{id:r[0],name:r[1],price:+r[2]||0,items:items,photo:r[4]||""};
  });
}

// Crée ou met à jour un combo/menu (plusieurs produits différents en une vente).
function saveCombo(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Combos"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,5).setValues([["ID","Nom","Prix","Contenu","Photo"]]);
    sh.getRange(1,1,1,5).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var photo=(p.photo||"").toString().trim();
  if(photo==="__CHUNKED__"){
    var total=+p.photoChunks||0;
    var cache=CacheService.getScriptCache(), parts=[];
    for(var c=0;c<total;c++){
      var part=cache.get("combophoto_"+p.id+"_"+c);
      if(part===null)return{ok:false,error:"Photo incomplète (morceau "+(c+1)+"/"+total+" manquant ou expiré), réessayez."};
      parts.push(part);
    }
    photo=parts.join("");
    for(var c2=0;c2<total;c2++)cache.remove("combophoto_"+p.id+"_"+c2);
  }
  // Le contenu du menu (ingrédients, alternatives, portions, menus imbriqués...)
  // peut devenir trop long pour tenir dans une seule URL une fois tout cumulé — dans
  // ce cas, envoyé en plusieurs morceaux (même mécanisme que les photos) plutôt que
  // de risquer un échec silencieux de la requête.
  var items=p.items||"";
  if(items==="__CHUNKED__"){
    var totalI=+p.itemsChunks||0;
    var cacheI=CacheService.getScriptCache(), partsI=[];
    for(var ci=0;ci<totalI;ci++){
      var partI=cacheI.get("comboitems_"+p.id+"_"+ci);
      if(partI===null)return{ok:false,error:"Contenu du menu incomplet (morceau "+(ci+1)+"/"+totalI+" manquant ou expiré), réessayez."};
      partsI.push(partI);
    }
    items=partsI.join("");
    for(var ci2=0;ci2<totalI;ci2++)cacheI.remove("comboitems_"+p.id+"_"+ci2);
  }
  var row=[p.id,p.name||"",+p.price||0,items,photo];
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      if(photo==="__KEEP__")row[4]=sh.getRange(i+2,5).getValue(); // photo inchangée, non renvoyée
      if(row[4]==="__KEEP__")row[4]=sh.getRange(i+2,5).getValue(); // garde-fou : jamais la valeur littérale
      sh.getRange(i+2,1,1,5).setValues([row]);
      return{ok:true,action:"updated"};
    }}
  }
  if(photo==="__KEEP__")row[4]=""; // création : rien à garder
  sh.appendRow(row);
  return{ok:true,action:"created"};
}
function deleteCombo(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Combos"), id=e.parameter.id;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Non trouvé"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}
  return{ok:false,error:"Non trouvé"};
}
function uploadComboPhotoChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("combophoto_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}
function uploadComboItemsChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("comboitems_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}
function uploadSaleItemsChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("saleitems_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}

// ── Plats (identiques à Combos dans leur structure, mais leur "Contenu" ne
// référence jamais un autre Plat ni un Menu — uniquement des produits bruts) ──
function getPlatsData(ss){
  var sh=ss.getSheetByName("Plats"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,5).getValues().filter(r=>r[0]).map(r=>{
    var items=(r[3]||"").toString().split(",").map(function(pair){
      var parts=pair.split(":");
      var alternates=(parts[3]||"").split("|").map(function(a){return a.trim();}).filter(Boolean);
      return{type:"product",productId:(parts[0]||"").trim(),qty:+parts[1]||1,portionId:(parts[2]||"").trim(),alternates:alternates};
    }).filter(function(it){return it.productId;});
    return{id:r[0],name:r[1],price:+r[2]||0,items:items,photo:r[4]||""};
  });
}
function savePlat(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Plats"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,5).setValues([["ID","Nom","Prix","Contenu","Photo"]]);
    sh.getRange(1,1,1,5).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var photo=(p.photo||"").toString().trim();
  if(photo==="__CHUNKED__"){
    var total=+p.photoChunks||0;
    var cache=CacheService.getScriptCache(), parts=[];
    for(var c=0;c<total;c++){
      var part=cache.get("platphoto_"+p.id+"_"+c);
      if(part===null)return{ok:false,error:"Photo incomplète (morceau "+(c+1)+"/"+total+" manquant ou expiré), réessayez."};
      parts.push(part);
    }
    photo=parts.join("");
    for(var c2=0;c2<total;c2++)cache.remove("platphoto_"+p.id+"_"+c2);
  }
  var items=p.items||"";
  if(items==="__CHUNKED__"){
    var totalI=+p.itemsChunks||0;
    var cacheI=CacheService.getScriptCache(), partsI=[];
    for(var ci=0;ci<totalI;ci++){
      var partI=cacheI.get("platitems_"+p.id+"_"+ci);
      if(partI===null)return{ok:false,error:"Contenu du plat incomplet (morceau "+(ci+1)+"/"+totalI+" manquant ou expiré), réessayez."};
      partsI.push(partI);
    }
    items=partsI.join("");
    for(var ci2=0;ci2<totalI;ci2++)cacheI.remove("platitems_"+p.id+"_"+ci2);
  }
  var row=[p.id,p.name||"",+p.price||0,items,photo];
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      if(photo==="__KEEP__")row[4]=sh.getRange(i+2,5).getValue();
      if(row[4]==="__KEEP__")row[4]=sh.getRange(i+2,5).getValue(); // garde-fou : jamais la valeur littérale
      sh.getRange(i+2,1,1,5).setValues([row]);
      return{ok:true,action:"updated"};
    }}
  }
  if(photo==="__KEEP__")row[4]="";
  sh.appendRow(row);
  return{ok:true,action:"created"};
}
function deletePlat(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Plats"), id=e.parameter.id;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Non trouvé"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}
  return{ok:false,error:"Non trouvé"};
}
function uploadPlatPhotoChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("platphoto_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}
function uploadPlatItemsChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("platitems_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}
// Migration en un clic depuis Config : convertit chaque Menu (Combo) qui n'est
// utilisé QUE comme ingrédient d'un autre Menu (jamais vendu directement en tant
// que Menu autonome ET ne contenant lui-même aucun autre Menu imbriqué) en un
// Plat — sans quoi les Menus créés avant l'introduction des Plats resteraient
// bloqués dans l'ancien système d'imbrication illimitée. Les Menus qui restent de
// vrais Menus (jamais référencés comme ingrédient d'un autre Menu) ne sont pas
// touchés.
function migrateToPlats(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var combosSh=ss.getSheetByName("Combos");
  if(!combosSh||combosSh.getLastRow()<=1)return{ok:true,migrated:0,names:""};
  var combos=getCombosData(ss);
  // Un menu est "migrable" s'il est référencé comme ingrédient (type "combo") par
  // au moins un autre menu, ET qu'il ne référence lui-même aucun autre menu imbriqué
  // (sinon il faudrait d'abord migrer ses propres sous-menus, cas volontairement
  // non traité automatiquement pour rester simple et sûr).
  var referencedIds={};
  combos.forEach(function(c){
    c.items.forEach(function(it){ if(it.type==="combo")referencedIds[it.comboId]=true; });
  });
  var toMigrate=combos.filter(function(c){
    return referencedIds[c.id]&&!c.items.some(function(it){return it.type==="combo";});
  });
  if(!toMigrate.length)return{ok:true,migrated:0,names:""};
  var platsSh=getOrCreate(ss,"Plats");
  if(platsSh.getLastRow()<=1){
    platsSh.getRange(1,1,1,5).setValues([["ID","Nom","Prix","Contenu","Photo"]]);
    platsSh.getRange(1,1,1,5).setFontWeight("bold"); platsSh.setFrozenRows(1);
  }
  var migratedIds={};
  toMigrate.forEach(function(c){
    var itemsStr=c.items.map(function(it){
      var alts=(it.alternates||[]).join("|");
      return it.productId+":"+it.qty+":"+(it.portionId||"")+":"+alts;
    }).join(",");
    platsSh.appendRow([c.id,c.name,c.price,itemsStr,c.photo||""]);
    migratedIds[c.id]=true;
  });
  // Retire les menus migrés de la feuille Combos (ils vivent désormais dans Plats),
  // en repartant du bas pour que les suppressions de ligne ne décalent pas les
  // index restants à traiter.
  var ids=combosSh.getRange(2,1,combosSh.getLastRow()-1,1).getValues();
  for(var i=ids.length-1;i>=0;i--){
    if(migratedIds[ids[i][0]])combosSh.deleteRow(i+2);
  }
  // Dans les menus restants, change le marqueur de type "combo" en "plat" pour
  // toute référence à un menu qui vient d'être migré — sans ça, ces menus
  // pointeraient vers un ID qui n'existe plus dans Combos.
  var remainingData=combosSh.getRange(2,1,combosSh.getLastRow()-1,5).getValues();
  for(var r=0;r<remainingData.length;r++){
    var content=(remainingData[r][3]||"").toString();
    if(!content)continue;
    var changed=false;
    var newContent=content.split(",").map(function(pair){
      var parts=pair.split(":");
      if(parts[4]&&parts[4].trim()==="combo"&&migratedIds[(parts[0]||"").trim()]){
        changed=true;
        return parts[0]+":"+parts[1]+"::"+":plat";
      }
      return pair;
    }).join(",");
    if(changed)combosSh.getRange(r+2,4).setValue(newContent);
  }
  return{ok:true,migrated:toMigrate.length,names:toMigrate.map(function(c){return c.name;}).join(", ")};
}

// Les 4 colonnes de stock (L à O) correspondent aux 4 premiers sites de l'onglet
// "Sites", DANS L'ORDRE où ils y sont — pas à un identifiant fixe "s1/s2/s3/s4".
// Ça permet de renommer ou recréer un site depuis Admin > Sites sans casser le stock
// (avant ce correctif, un site dont l'ID n'était pas exactement s1/s2/s3/s4 provoquait
// l'erreur "Site invalide" sur toute modification de stock).
function siteColumn(ss, siteId){
  if(!siteId)return null;
  var sh=ss.getSheetByName("Sites");
  if(!sh||sh.getLastRow()<=1)return null;
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){
    if(ids[i][0].toString()===siteId.toString()){
      // Les 4 premiers sites gardent leurs colonnes historiques L:O (12-15), pour ne
      // jamais avoir à déplacer les données déjà enregistrées. Le 5e site et les
      // suivants utilisent de nouvelles colonnes à partir de R (18) — après
      // PrixAchatTTC (P) et Paliers (Q), donc sans jamais les décaler non plus.
      return i<4 ? 12+i : 18+(i-4);
    }
  }
  return null;
}

function getProductsData(ss){
  var sh=ss.getSheetByName("Produits"); if(!sh||sh.getLastRow()<=1)return [];
  var sites=getSitesData(ss); // pour associer chaque site à SA colonne de stock, dans l'ordre
  // Colonne Y (25) réservée à "Contenants", Z (26) à "Couleur", AA (27) à "ReturnFor"
  // (retour de consigne), AB (28) à "HappyPrice" (prix Happy Hour), AC (29) à
  // "SitesVendus" (IDs de sites séparés par virgule où le produit est vendu — vide =
  // tous les sites) — volontairement loin des colonnes de stock des sites au-delà du
  // 4e (qui démarrent à R=18 et peuvent grandir), pour ne jamais entrer en collision
  // avec elles.
  // Colonne AD (30) : "AfficherBarreContenants" (vide/TRUE = affiché par défaut dans
  // la barre de contenants en haut de la Caisse ; FALSE = exclu volontairement).
  var CONTAINERS_COL=25, COLOR_COL=26, RETURNFOR_COL=27, HAPPYPRICE_COL=28, SITESVENDUS_COL=29, SHOWBAR_COL=30;
  var maxCol=Math.max(SHOWBAR_COL, sites.length>4 ? 18+(sites.length-4)-1 : 17);
  return sh.getRange(2,1,sh.getLastRow()-1,maxCol).getValues()
    .filter(r=>r[0]).map(r=>{
      var stock={};
      sites.forEach(function(s,i){
        var col=i<4?12+i:18+(i-4);
        stock[s.id]=+r[col-1]||0;
      });
      return {
        id:r[0],name:r[1],price:+r[2],cat:r[3],emoji:r[4],photo:r[5]||"",barcode:(r[6]!=null&&r[6]!==""?r[6].toString():""),
        drink:r[7]===true||r[7]==="TRUE"||r[7]==="true",
        sellByVolume:r[8]===true||r[8]==="TRUE"||r[8]==="true",
        unit:r[9]||"",baseQty:+r[10]||1,
        stock:stock,
        costPrice:+r[15]||0,
        presets:(r[16]||"").toString(),
        containers:(r[CONTAINERS_COL-1]||"").toString(),
        color:(r[COLOR_COL-1]||"").toString(),
        returnFor:(r[RETURNFOR_COL-1]||"").toString(),
        happyPrice:+r[HAPPYPRICE_COL-1]||0,
        siteAvailability:(r[SITESVENDUS_COL-1]||"").toString(),
        showInContainerBar:r[SHOWBAR_COL-1]!==false&&r[SHOWBAR_COL-1]!=="FALSE"&&r[SHOWBAR_COL-1]!=="false"
      };
    });
}
function getSitesData(ss){
  var sh=ss.getSheetByName("Sites"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues().filter(r=>r[0])
    .map(r=>({id:r[0],name:r[1],village:r[2],color:r[3],
      active:r[4]!==false&&r[4]!=="FALSE"&&r[4]!=="false",
      isSalesPoint:r[5]!==false&&r[5]!=="FALSE"&&r[5]!=="false"}));
}
function getCategoriesData(ss){
  var sh=ss.getSheetByName("Categories"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,2).getValues().filter(r=>r[0])
    .map(r=>({id:r[0],name:r[1]}));
}
function getUsersData(ss){
  var sh=ss.getSheetByName("Utilisateurs"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues().filter(r=>r[0])
    .map(r=>({id:r[0],name:r[1],pin:r[2].toString(),role:r[3],productOrder:(r[4]||"").toString(),cameraPref:(r[5]||"").toString()}));
}
// Sauvegarde l'ordre personnalisé des produits sur l'écran caisse pour un utilisateur
// donné (colonne E "OrdreProduits" de la feuille Utilisateurs) — ne touche à rien
// d'autre sur sa ligne (nom, PIN, rôle inchangés).
function saveProductOrder(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Utilisateurs"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1||!p.userId)return{ok:false,error:"Utilisateur introuvable"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.userId.toString()){
    sh.getRange(i+2,5).setValue(p.order||"");
    return{ok:true};
  }}
  return{ok:false,error:"Utilisateur introuvable"};
}
// Sauvegarde la caméra préférée pour le scan (colonne F "CameraPref") — en plus du
// localStorage sur l'appareil, pour survivre à un vidage de cache et suivre la
// personne si elle réutilise le même appareil après une réinstallation.
function saveCameraPref(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Utilisateurs"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1||!p.userId)return{ok:false,error:"Utilisateur introuvable"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.userId.toString()){
    sh.getRange(i+2,6).setValue(p.cameraId||"");
    return{ok:true};
  }}
  return{ok:false,error:"Utilisateur introuvable"};
}

function getContainersData(ss){
  var sh=ss.getSheetByName("Contenants"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({site:r[0],productId:r[1],label:r[2],size:+r[3]||0,remaining:+r[4]||0,openedAt:r[5]}));
}

// Ouvre un nouveau contenant (fût/bouteille/cubi) pour un produit sur un site donné.
// S'il y en avait déjà un ouvert pour ce produit+site, il est remplacé (l'ancien est
// considéré abandonné — l'appli doit normalement demander de "Terminer" avant d'en
// ouvrir un nouveau, mais on ne bloque pas côté serveur pour rester simple).
function openContainer(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Contenants"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,6).setValues([["SiteID","ProduitID","Contenant","TailleCl","RestantCl","DateOuverture"]]);
    sh.getRange(1,1,1,6).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var size=+p.size||0; if(!p.site||!p.id||size<=0)return{ok:false,error:"Paramètres manquants"};
  var tz=Session.getScriptTimeZone(), now=Utilities.formatDate(new Date(),tz,"dd/MM/yyyy HH:mm");
  var row=[p.site,p.id,p.label||"",size,size,now];
  if(sh.getLastRow()>1){
    var data=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
    for(var i=0;i<data.length;i++){
      if(data[i][0].toString()===p.site.toString()&&data[i][1].toString()===p.id.toString()){
        sh.getRange(i+2,1,1,6).setValues([row]); decrementReserve(ss,p.site,p.id,p.label||"",size); return{ok:true,action:"updated"};
      }
    }
  }
  sh.appendRow(row);
  decrementReserve(ss,p.site,p.id,p.label||"",size);
  return{ok:true,action:"created"};
}
// Décompte 1 unité de la réserve précise (voir addToReserve) quand ce contenant est
// ouvert — ne fait rien si aucune réserve n'a été enregistrée pour ce type (repli
// silencieux, l'écran Contenants utilisera alors le calcul dérivé du stock total).
function decrementReserve(ss,site,id,label,size){
  var sh=ss.getSheetByName("ReserveContenants"); if(!sh||sh.getLastRow()<=1)return;
  var data=sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
  for(var i=0;i<data.length;i++){
    if(data[i][0].toString()===site.toString()&&data[i][1].toString()===id.toString()
      &&data[i][2].toString()===label&&(+data[i][3]||0)===size){
      sh.getRange(i+2,5).setValue(Math.max(0,(+data[i][4]||0)-1));
      return;
    }
  }
}

// Termine un contenant : ce qu'il restait dans la jauge est déduit du stock du
// produit (pour corriger les pertes/verres renversés et recaler la réalité), puis
// la ligne est supprimée.
function closeContainer(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Contenants"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Aucun contenant ouvert"};
  var data=sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
  for(var i=0;i<data.length;i++){
    if(data[i][0].toString()===p.site.toString()&&data[i][1].toString()===p.id.toString()){
      var remaining=+data[i][4]||0;
      if(remaining>0){
        var prodSh=ss.getSheetByName("Produits"), col=siteColumn(ss,p.site);
        if(col){
          var pIds=prodSh.getRange(2,1,prodSh.getLastRow()-1,1).getValues();
          for(var j=0;j<pIds.length;j++){
            if(pIds[j][0].toString()===p.id.toString()){
              var cell=prodSh.getRange(j+2,col);
              cell.setValue(Math.max(0,+cell.getValue()-remaining));
              break;
            }
          }
        }
      }
      sh.deleteRow(i+2);
      return{ok:true,reconciled:remaining};
    }
  }
  return{ok:false,error:"Contenant introuvable"};
}
// Corrige directement la jauge d'un contenant ouvert (ex: erreur de saisie, écart
// constaté à l'inventaire) — ajuste AUSSI le stock global du produit du même écart,
// pour que jauge et stock restent cohérents entre eux.
function adjustContainer(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Contenants"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Aucun contenant ouvert"};
  var newRemaining=+p.remaining; if(isNaN(newRemaining)||newRemaining<0)return{ok:false,error:"Valeur invalide"};
  var data=sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
  for(var i=0;i<data.length;i++){
    if(data[i][0].toString()===p.site.toString()&&data[i][1].toString()===p.id.toString()){
      var oldRemaining=+data[i][4]||0;
      var delta=newRemaining-oldRemaining;
      sh.getRange(i+2,5).setValue(newRemaining);
      if(delta!==0){
        var prodSh=ss.getSheetByName("Produits"), col=siteColumn(ss,p.site);
        if(col){
          var pIds=prodSh.getRange(2,1,prodSh.getLastRow()-1,1).getValues();
          for(var j=0;j<pIds.length;j++){
            if(pIds[j][0].toString()===p.id.toString()){
              var cell=prodSh.getRange(j+2,col);
              cell.setValue(Math.max(0,+cell.getValue()+delta));
              break;
            }
          }
        }
      }
      return{ok:true,newRemaining:newRemaining};
    }
  }
  return{ok:false,error:"Contenant introuvable"};
}

function saveProduct(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), p=e.parameter;
  // .trim() en sécurité : un espace ou un caractère invisible ajouté quelque part en
  // route (encodage d'URL, copier-coller...) ferait échouer une comparaison stricte
  // avec "__KEEP__" sans que ce soit visible autrement.
  var photo=(p.photo||"").toString().trim();
  if(photo==="__CHUNKED__"){
    // La photo a été envoyée en plusieurs morceaux via uploadPhotoChunk (POST ne
    // fonctionne pas de façon fiable sur ce déploiement, et une photo en un seul
    // GET dépasse la longueur d'URL maximale). On la reconstitue ici.
    var total=+p.photoChunks||0;
    var cache=CacheService.getScriptCache(), parts=[];
    for(var c=0;c<total;c++){
      var part=cache.get("photo_"+p.id+"_"+c);
      if(part===null)return{ok:false,error:"Photo incomplète (morceau "+(c+1)+"/"+total+" manquant ou expiré), réessayez."};
      parts.push(part);
    }
    photo=parts.join("");
    for(var c2=0;c2<total;c2++)cache.remove("photo_"+p.id+"_"+c2); // nettoyage
  }
  var costPrice=+p.costPrice||0;
  var presets=(p.presets||"").toString();
  var containers=(p.containers||"").toString();
  var color=(p.color||"").toString();
  var returnFor=(p.returnFor||"").toString();
  var happyPrice=+p.happyPrice||0;
  var siteAvailability=(p.siteAvailability||"").toString();
  var showInContainerBar=p.showInContainerBar===undefined?true:(p.showInContainerBar==="true"||p.showInContainerBar===true);
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      // "__KEEP__" : la photo n'a pas changé côté appli, on ne l'a donc pas renvoyée
      // (gain de temps important) — on garde simplement celle déjà enregistrée.
      var finalPhoto = photo==="__KEEP__" ? sh.getRange(i+2,6).getValue() : photo;
      // Garde-fou : quelle qu'en soit la cause, la valeur littérale "__KEEP__" ne
      // doit JAMAIS finir stockée comme si c'était une vraie photo — dans ce cas on
      // garde la valeur déjà en place plutôt que d'écraser une bonne photo.
      if(finalPhoto==="__KEEP__")finalPhoto=sh.getRange(i+2,6).getValue();
      var meta=[p.id,p.name,+p.price,p.cat,p.emoji,finalPhoto,p.barcode||"",
        p.drink==="true",p.sellByVolume==="true",p.unit||"",+p.baseQty||1];
      // Mise à jour : on NE touche PAS aux colonnes de stock (L:O), pour ne jamais écraser
      // une modification de stock faite entre-temps par un autre appareil/utilisateur.
      sh.getRange(i+2,7).setNumberFormat("@"); // Codebarres en texte : empêche Sheets de
      // le convertir en nombre et de supprimer un éventuel zéro de tête (cause du
      // "produit non trouvé" au rescan d'un code-barres déjà enregistré).
      sh.getRange(i+2,1,1,11).setValues([meta]);
      sh.getRange(i+2,16).setValue(costPrice); // P = Prix d'achat TTC
      sh.getRange(i+2,17).setValue(presets);   // Q = Paliers de vente au volume
      sh.getRange(i+2,25).setValue(containers);// Y = Tailles de contenant (jauge fûts/bouteilles/cubis)
      sh.getRange(i+2,26).setValue(color);     // Z = Couleur de la jauge (comme les sites)
      sh.getRange(i+2,27).setValue(returnFor); // AA = Retour de consigne pour (ID du produit "consigne")
      sh.getRange(i+2,28).setValue(happyPrice);// AB = Prix Happy Hour
      sh.getRange(i+2,29).setValue(siteAvailability);// AC = Sites où le produit est vendu (vide = tous)
      sh.getRange(i+2,30).setValue(showInContainerBar);// AD = Affiché dans la barre de contenants (Caisse)
      return{ok:true,action:"updated"};
    }}
  }
  // Création d'un nouveau produit : "__KEEP__" n'a pas de sens ici (rien à garder), on
  // traite comme une photo vide dans ce cas précis (ne devrait normalement pas arriver).
  var createPhoto = photo==="__KEEP__" ? "" : photo;
  var meta=[p.id,p.name,+p.price,p.cat,p.emoji,createPhoto,p.barcode||"",
    p.drink==="true",p.sellByVolume==="true",p.unit||"",+p.baseQty||1];
  // Stock initialisé à 0 sur tous les sites (les 4 colonnes historiques L:O, plus une
  // colonne par site au-delà du 4e, jusqu'aux colonnes Y=25/Z=26/AA=27/AB=28/AC=29/AD=30
  // réservées à Contenants/Couleur/ReturnFor/HappyPrice/SitesVendus/AfficherBarre —
  // donc jusqu'à 11 sites au total sans collision).
  var nbSites=getSitesData(ss).length;
  var row=meta.slice();               // colonnes 1-11 (A:K)
  row.push(0,0,0,0);                  // colonnes 12-15 (L:O) — toujours réservées aux 4 premiers sites
  row.push(costPrice);                // colonne 16 (P)
  row.push(presets);                  // colonne 17 (Q)
  var extra=nbSites>4?nbSites-4:0;
  for(var s=0;s<extra;s++)row.push(0);// colonnes 18+ (R, S...) pour le 5e site et au-delà
  while(row.length<29)row.push("");   // comble jusqu'à la colonne 29 si besoin
  row[24]=containers;                 // colonne 25 (Y)
  row[25]=color;                      // colonne 26 (Z)
  row[26]=returnFor;                  // colonne 27 (AA)
  row[27]=happyPrice;                 // colonne 28 (AB)
  row[28]=siteAvailability;           // colonne 29 (AC)
  row[29]=showInContainerBar;         // colonne 30 (AD)
  sh.appendRow(row);
  sh.getRange(sh.getLastRow(),7).setNumberFormat("@"); // Codebarres en texte, même raison que ci-dessus
  return{ok:true,action:"created"};
}
// Reçoit un morceau de photo (base64) et le stocke temporairement (10 min) en attendant
// que tous les morceaux soient arrivés ; saveProduct les réassemble ensuite via photoChunks.
function uploadPhotoChunk(e){
  var p=e.parameter;
  if(!p.id||p.idx===undefined||!p.chunk)return{ok:false,error:"Paramètres manquants"};
  CacheService.getScriptCache().put("photo_"+p.id+"_"+p.idx, p.chunk, 600);
  return{ok:true};
}
function deleteProduct(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), id=e.parameter.id;
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===id.toString()){sh.deleteRow(i+2);return{ok:true};}}}
  return{ok:false,error:"Non trouvé"};
}
function saveSite(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Sites"), p=e.parameter;
  var active=p.active===undefined?true:(p.active==="true"||p.active===true);
  var isSalesPoint=p.isSalesPoint===undefined?true:(p.isSalesPoint==="true"||p.isSalesPoint===true);
  var row=[p.id,p.name,p.village||"",p.color||"#CC0000",active,isSalesPoint];
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){sh.getRange(i+2,1,1,6).setValues([row]);return{ok:true,action:"updated"};}}}
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
  var col=siteColumn(ss,p.site); if(!col)return{ok:false,error:"Site invalide"};
  if(sh.getLastRow()>1){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      var cell=sh.getRange(i+2,col);
      cell.setValue(p.mode==="delta"?Math.max(0,+cell.getValue()+(+p.qty)):Math.max(0,+p.qty));
      return{ok:true,newQty:+cell.getValue()};}}}
  return{ok:false,error:"Produit non trouvé"};
}
function transferStock(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), p=e.parameter;
  var colFrom=siteColumn(ss,p.from), colTo=siteColumn(ss,p.to);
  if(!colFrom||!colTo)return{ok:false,error:"Site invalide"};
  var items; try{items=JSON.parse(safeDecodeItems(p.items));}catch(err){return{ok:false,error:"Items invalides: "+err};}
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
  // Le détail des articles ("items") peut devenir long pour un panier chargé
  // (beaucoup de produits différents) — au-delà d'un certain seuil, envoyé en
  // plusieurs morceaux (même mécanisme que les photos et les menus) plutôt que de
  // risquer un échec de la requête (URL trop longue).
  var itemsRaw=p.items||"";
  if(itemsRaw==="__CHUNKED__"){
    var totalIt=+p.itemsChunks||0;
    var cacheIt=CacheService.getScriptCache(), partsIt=[];
    for(var ii=0;ii<totalIt;ii++){
      var partIt=cacheIt.get("saleitems_"+p.id+"_"+ii);
      if(partIt===null)return{ok:false,error:"Détail de la vente incomplet (morceau "+(ii+1)+"/"+totalIt+" manquant ou expiré), réessayez."};
      partsIt.push(partIt);
    }
    itemsRaw=partsIt.join("");
    for(var ii2=0;ii2<totalIt;ii2++)cacheIt.remove("saleitems_"+p.id+"_"+ii2);
  }
  var tz=Session.getScriptTimeZone(), now=new Date(), saleId=now.getTime().toString();
  var dateStr=Utilities.formatDate(now,tz,"dd/MM/yyyy"), heureStr=Utilities.formatDate(now,tz,"HH:mm:ss");
  sh.appendRow([saleId,dateStr,heureStr,
    p.site,p.siteName,+p.total,p.payment,itemsRaw,p.member||"",+p.nbItems,p.caissier||"",p.splitPart||"",p.readyCombos||"",p.comboLines||"",p.readyPlats||""]);
  // Empêche Google Sheets de convertir les colonnes Date/Heure en vraies dates (ce qui
  // cassait le filtre "Aujourd'hui" et les totaux dans les Rapports : la date revenait
  // au format ISO complet au lieu du "dd/MM/yyyy" attendu par l'appli, donc plus rien
  // ne correspondait) — on force le format texte puis on réécrit la valeur pour que
  // Sheets ne la réinterprète plus comme une date.
  var lastRow=sh.getLastRow();
  sh.getRange(lastRow,2).setNumberFormat("@").setValue(dateStr);
  sh.getRange(lastRow,3).setNumberFormat("@").setValue(heureStr);
  // Décrémenter stock
  var prodSh=ss.getSheetByName("Produits"), col=siteColumn(ss,p.site);
  var items; try{items=JSON.parse(safeDecodeItems(itemsRaw));}catch(err){items=[];}
  // Diagnostic : avant, un article non trouvé (mauvais ID, site invalide...) était
  // ignoré en silence et la vente répondait quand même "ok" — impossible de savoir
  // pourquoi un stock ne bougeait pas. On renvoie maintenant explicitement ce qui a
  // été décompté et ce qui ne l'a pas été.
  var decremented=[], notFound=[];
  if(col&&prodSh.getLastRow()>1){
    var pIds=prodSh.getRange(2,1,prodSh.getLastRow()-1,1).getValues();
    items.forEach(function(item){
      var found=false;
      for(var i=0;i<pIds.length;i++){if(pIds[i][0].toString()===item.id.toString()){
        var cell=prodSh.getRange(i+2,col);
        cell.setValue(Math.max(0,+cell.getValue()-item.qty));
        decremented.push(item.id+":-"+item.qty);
        found=true; break;}}
      if(!found)notFound.push(String(item.id));
    });
  } else if(items.length){
    notFound=items.map(function(it){return String(it.id);});
  }
  // Décrémenter aussi la jauge du contenant ouvert pour ce produit+site, s'il y en a
  // un (fût/bouteille/cubi entamé) — en plus du stock global, sans jamais aller sous 0.
  var contSh=ss.getSheetByName("Contenants");
  if(contSh&&contSh.getLastRow()>1&&items.length){
    var cData=contSh.getRange(2,1,contSh.getLastRow()-1,6).getValues();
    items.forEach(function(item){
      for(var k=0;k<cData.length;k++){
        if(cData[k][0].toString()===p.site.toString()&&cData[k][1].toString()===item.id.toString()){
          var rCell=contSh.getRange(k+2,5);
          rCell.setValue(Math.max(0,+rCell.getValue()-item.qty));
          break;
        }
      }
    });
  }
  var result={ok:true,saleId:saleId,stockUpdated:decremented};
  if(notFound.length)result.stockWarning="Stock NON décompté (site="+p.site+", colonne="+col+") pour ID(s): "+notFound.join(", ");
  return result;
}
// Modifie une vente déjà enregistrée EN PLACE (même ID, même date/heure/site) — ne
// touche que le total et le détail des articles. Utilisé pour corriger une commande
// passée sans créer une nouvelle vente à la date du jour. Le stock doit être ajusté
// séparément côté client (delta entre l'ancien et le nouveau contenu).
function updateSale(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Ventes"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Aucune vente enregistrée"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){
    if(ids[i][0].toString()===p.id.toString()){
      sh.getRange(i+2,6).setValue(+p.total||0);   // F = Total
      if(p.payment!==undefined)sh.getRange(i+2,7).setValue(p.payment); // G = Paiement (facultatif)
      sh.getRange(i+2,8).setValue(p.items||"");   // H = Articles
      sh.getRange(i+2,10).setValue(+p.nbItems||0);// J = NbArticles
      return{ok:true};
    }
  }
  return{ok:false,error:"Vente introuvable"};
}
// Supprime définitivement une vente. Le stock des articles qu'elle contenait doit
// être restitué CÔTÉ CLIENT avant cet appel (via updateStock delta), car ce script
// ne connaît pas le site à recréditer sans reparser les articles ici — le client a
// déjà cette info sous la main.
function deleteSale(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Ventes"), id=e.parameter.id;
  if(!sh||sh.getLastRow()<=1)return{ok:false,error:"Aucune vente enregistrée"};
  var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ids.length;i++){
    if(ids[i][0].toString()===id.toString()){
      sh.deleteRow(i+2);
      return{ok:true};
    }
  }
  return{ok:false,error:"Vente introuvable"};
}
function getSales(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Ventes"), p=e.parameter;
  if(!sh||sh.getLastRow()<=1)return{ok:true,sales:[]};
  var tz=Session.getScriptTimeZone();
  // Certaines lignes déjà enregistrées ont été converties en vraies dates par Google
  // Sheets (voir saveSale) : on les reformate à la volée en "dd/MM/yyyy"/"HH:mm:ss"
  // pour que les anciennes ventes redeviennent lisibles dans les Rapports, pas
  // seulement les nouvelles.
  var fmtDate=function(v){ return v instanceof Date ? Utilities.formatDate(v,tz,"dd/MM/yyyy") : v; };
  var fmtTime=function(v){ return v instanceof Date ? Utilities.formatDate(v,tz,"HH:mm:ss") : v; };
  var data=sh.getRange(2,1,sh.getLastRow()-1,15).getValues();
  var sales=data.filter(function(r){
    if(!r[0])return false;
    if(p.site&&p.site!=="all"&&r[3]!==p.site)return false;
    if(p.date&&fmtDate(r[1])!==p.date)return false;
    return true;
  }).map(function(r){return{id:r[0],date:fmtDate(r[1]),time:fmtTime(r[2]),siteId:r[3],siteName:r[4],total:r[5],payment:r[6],items:r[7],member:r[8],caissier:r[10],readyCombos:r[12]||"",comboLines:r[13]||"",readyPlats:r[14]||""};});
  return{ok:true,sales:sales};
}
