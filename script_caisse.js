// ═══════════════════════════════════════════════════════
//  ANF 87 — Google Apps Script CAISSE v2
//  Sheets: Produits | Sites | Categories | Ventes | Utilisateurs
// ═══════════════════════════════════════════════════════

// Actions qui ÉCRIVENT dans le classeur : protégées par un verrou pour éviter
// les conflits quand 2-3 utilisateurs utilisent la caisse en même temps.
var WRITE_ACTIONS = ["saveProduct","deleteProduct","saveSite","deleteSite","saveCategory",
  "deleteCategory","saveUser","deleteUser","saveSale","updateStock","transferStock","uploadPhotoChunk",
  "openContainer","closeContainer","savePortion","deletePortion","uploadPortionPhotoChunk"];
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
      case "savePortion":     result = savePortion(e); break;
      case "deletePortion":   result = deletePortion(e); break;
      case "uploadPortionPhotoChunk": result = uploadPortionPhotoChunk(e); break;
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
  // Contenants ouverts (fûts/bouteilles/cubis entamés, un par site+produit)
  var cont = getOrCreate(ss,"Contenants");
  if(cont.getLastRow()<=1){
    cont.getRange(1,1,1,6).setValues([["SiteID","ProduitID","Contenant","TailleCl","RestantCl","DateOuverture"]]);
    cont.getRange(1,1,1,6).setFontWeight("bold");
    cont.setFrozenRows(1);
  }
  // Portions de vente (paliers) : demis/pichets/tailles de gobelet... avec leur
  // propre photo, liées à un produit "liquide" (celui vendu au volume).
  var port = getOrCreate(ss,"Portions");
  if(port.getLastRow()<=1){
    port.getRange(1,1,1,6).setValues([["ID","ProduitID","Nom","TailleCl","Prix","Photo"]]);
    port.getRange(1,1,1,6).setFontWeight("bold");
    port.setFrozenRows(1);
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
    portions:getPortionsData(ss), ts:Date.now()};
}

function getPortionsData(ss){
  var sh=ss.getSheetByName("Portions"); if(!sh||sh.getLastRow()<=1)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues().filter(r=>r[0]&&r[1])
    .map(r=>({id:r[0],productId:r[1],name:r[2],size:+r[3]||0,price:+r[4]||0,photo:r[5]||""}));
}

// Crée ou met à jour une portion de vente (demi, pichet, taille de gobelet...).
function savePortion(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=getOrCreate(ss,"Portions"), p=e.parameter;
  if(sh.getLastRow()<=1){
    sh.getRange(1,1,1,6).setValues([["ID","ProduitID","Nom","TailleCl","Prix","Photo"]]);
    sh.getRange(1,1,1,6).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  var photo=p.photo||"";
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
  var row=[p.id,p.productId,p.name||"",+p.size||0,+p.price||0,photo];
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
      sh.getRange(i+2,1,1,6).setValues([row]);
      return{ok:true,action:"updated"};
    }}
  }
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
  // Colonne Y (25) réservée à "Contenants", Z (26) à "Couleur" — volontairement loin
  // des colonnes de stock des sites au-delà du 4e (qui démarrent à R=18 et peuvent
  // grandir), pour ne jamais entrer en collision avec elles.
  var CONTAINERS_COL=25, COLOR_COL=26;
  var maxCol=Math.max(COLOR_COL, sites.length>4 ? 18+(sites.length-4)-1 : 17);
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
        color:(r[COLOR_COL-1]||"").toString()
      };
    });
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
        sh.getRange(i+2,1,1,6).setValues([row]); return{ok:true,action:"updated"};
      }
    }
  }
  sh.appendRow(row);
  return{ok:true,action:"created"};
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

function saveProduct(e){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Produits"), p=e.parameter;
  var photo=p.photo||"";
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
  // Colonnes A:K = infos produit (jamais le stock, colonnes L:O)
  var meta=[p.id,p.name,+p.price,p.cat,p.emoji,photo,p.barcode||"",
    p.drink==="true",p.sellByVolume==="true",p.unit||"",+p.baseQty||1];
  var costPrice=+p.costPrice||0;
  var presets=(p.presets||"").toString();
  var containers=(p.containers||"").toString();
  var color=(p.color||"").toString();
  if(sh.getLastRow()>1){
    var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<ids.length;i++){if(ids[i][0].toString()===p.id.toString()){
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
      return{ok:true,action:"updated"};
    }}
  }
  // Création d'un nouveau produit : stock initialisé à 0 sur tous les sites
  // (les 4 colonnes historiques L:O, plus une colonne par site au-delà du 4e, jusqu'à
  // les colonnes Y=25/Z=26 réservées à Contenants/Couleur — donc jusqu'à 11 sites
  // au total sans collision).
  var nbSites=getSitesData(ss).length;
  var row=meta.slice();               // colonnes 1-11 (A:K)
  row.push(0,0,0,0);                  // colonnes 12-15 (L:O) — toujours réservées aux 4 premiers sites
  row.push(costPrice);                // colonne 16 (P)
  row.push(presets);                  // colonne 17 (Q)
  var extra=nbSites>4?nbSites-4:0;
  for(var s=0;s<extra;s++)row.push(0);// colonnes 18+ (R, S...) pour le 5e site et au-delà
  while(row.length<24)row.push("");   // comble jusqu'à la colonne 24 si besoin
  row[24]=containers;                 // colonne 25 (Y)
  row[25]=color;                      // colonne 26 (Z)
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
  var tz=Session.getScriptTimeZone(), now=new Date(), saleId=now.getTime().toString();
  sh.appendRow([saleId,Utilities.formatDate(now,tz,"dd/MM/yyyy"),Utilities.formatDate(now,tz,"HH:mm:ss"),
    p.site,p.siteName,+p.total,p.payment,p.items,p.member||"",+p.nbItems,p.caissier||"",p.splitPart||""]);
  // Décrémenter stock
  var prodSh=ss.getSheetByName("Produits"), col=siteColumn(ss,p.site);
  var items; try{items=JSON.parse(safeDecodeItems(p.items));}catch(err){items=[];}
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
