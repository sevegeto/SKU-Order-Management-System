const CONFIG_COMPRAS = {
  sheets: {
    compras: "Compras" // o el nombre de tu hoja
  }
};

// =============================================================================
// PACK COMPLETO: ZLOOKUP + UPC Lookup Service v8.1.3 DEPURADO
// =============================================================================
const CONFIG_BUSCA_UPC = {
  CACHE_DURATION: 21600,
  SHEET_NAME: "BuscaUPCs",
  HEADER_ROW: 5,
  START_ROW: 6,
  UPC_COLUMN: 1,
  MAX_RETRIES: 2,
  INITIAL_DELAY: 1000,
  MAX_BACKOFF_MS: 10000,
  BATCH_SIZE: 5,
  BATCH_PAUSE_MS: 2500,
  SEARCH_PAUSE_MS: 750,
  MAX_VENDORS: 6,
  PROPERTY_KEY: 'UPC_API_KEY',
  MIN_SEARCH_LENGTH: 3,
  MAX_SEARCH_LENGTH: 80,
  ENABLE_PROGRESSIVE_SEARCH: true
};

// --- Para no romper tu otro onEdit de compras ---
function isCustomFunctionCall() {
  try {
    // This will fail if called from a custom function, which doesn't have UI permissions.
    SpreadsheetApp.getUi();
    return false;
  } catch (e) {
    return true;
  }
}

/**
 * A universal onEdit trigger that delegates to other specific onEdit functions
 * based on the sheet name.
 */
function onEdit(e) {
  try {
    const hoja = e.range.getSheet();
    const nombreHoja = hoja.getName();
    if (typeof CONFIG_COMPRAS!== 'undefined' && CONFIG_COMPRAS && CONFIG_COMPRAS.sheets && nombreHoja === CONFIG_COMPRAS.sheets.compras && typeof onEditShippo_ === 'function') {
      onEditShippo_(e);
    }
    if (nombreHoja === CONFIG_BUSCA_UPC.SHEET_NAME && typeof updateProductInfoFromUPCs_onEdit_ === 'function') {
      updateProductInfoFromUPCs_onEdit_(e);
    }
  } catch (err) { Logger.log(`Error onEdit universal: ${err.message}`); }
}

const REGEX = {
  GTIN: /^\d{12,14}$/, COLUMN_LETTER: /^[A-Z]+$/, ONLY_NUMBERS_SYMBOLS: /^[\d\s\-\.\#\(\)\/\*\&]+$/,
  ONLY_NUMBERS: /^\d+$/, SKU_ONLY: /^sku[\s\-:]?\d+$/i, EBAY_PREFIXES: /^(US!|NEW!|HOT!|SALE!|BEST!|\*+|#+)\s*/gi,
  VARIANT_SUFFIX: /,\s*(for|fits|compatible|w\/|with)\s+.*/gi,
  COLOR_PARENTHESES: /\s*\([^)]*(?:version|color|black|white|grey|gray|blue|red|green|camo|fde|od|tan)\s*[^)]*\)\s*$/gi,
  THREAD_SPECS: /\b\d+\/\d+x\d+\b/g, CALIBERS: /\b\d{2,3}(BLK|WIN|WSM|REM|ACP|AUTO|MAG|S&W|SPECIAL)\b/gi,
  MEASUREMENTS: /\b\d+(\.\d+)?["']?\s*(inch|in|inches|mm|cm|oz|lbs?|rd)\b/gi,
  ROUNDS_PACKS: /\b\d+\s*\/\s*rd\b/gi, COUNT_PACKS: /\b\d+\s*-?\s*(pack|pc|pcs|piece|count)\b/gi,
  MULTIPLE_SPACES: /\s+/g, DASHES: /\s*[-–—]\s*/g, TRAILING_PUNCT: /\s*[,;]\s*$/
};
const BAD_NAME_PATTERNS = [/\b(shipping|handling|delivery|freight|postage|courier|mail)\b/, /\b(fedex|usps|ups|dhl)\b/, /\b(fee|charge|adjustment|discount|credit|refund|tax|insurance)\b/, /\b(test|prueba|ejemplo|sample|demo)\b/, /^(n\/a|tbd|pending|various|none|null|unknown)$/, /\b(free\s+gift|promo\s+item|bonus)\b/];

const BASE_HEADERS = ["upc","ean","title","description","brand","model","color","size","dimension","weight","category","currency","lowest_recorded_price","highest_recorded_price","images","offers","asin","elid","Error"];

function generateVendorHeaders(max) {
  const h = [];
  for (let i = 1; i <= max; i++) {
    h.push(`vendor_${i}_merchant`, `vendor_${i}_price`, `vendor_${i}_currency`, `vendor_${i}_condition`, `vendor_${i}_domain`, `vendor_${i}_availability`, `vendor_${i}_link`);
  }
  return h;
}

const VENDOR_HEADERS = generateVendorHeaders(CONFIG_BUSCA_UPC.MAX_VENDORS);
const FINAL_HEADERS = [...BASE_HEADERS,...VENDOR_HEADERS];
const IDX = FINAL_HEADERS.reduce((map, header, i) => (map[header] = i, map), {});
const CACHE_NS_PROP = 'UPC_CACHE_NS_V4';

class APIMetrics {
  constructor() {
    this.reset();
  }

  reset() {
    this.requests = 0;
    this.cacheHits = 0;
    this.cacheMisses = 0;
    this.errors = 0;
    this.rateLimits = 0;
    this.progressiveSearches = 0;
    this.startTime = new Date();
  }

  logRequest() { this.requests++; }
  logCacheHit() { this.cacheHits++; }
  logCacheMiss() { this.cacheMisses++; }
  logError() { this.errors++; }
  logRateLimit() { this.rateLimits++; }
  logProgressiveSearch() { this.progressiveSearches++; }

  getSummary() {
    const elapsed = (new Date() - this.startTime) / 1000;
    const total = this.cacheHits + this.cacheMisses;
    const cacheRate = total > 0 ? ((this.cacheHits / total) * 100).toFixed(1) : "0.0";
    const rps = elapsed > 0 ? (this.requests / elapsed).toFixed(2) : "0.00";
    return {
      requests: this.requests,
      cacheHits: this.cacheHits,
      cacheMisses: this.cacheMisses,
      cacheHitRate: `${cacheRate}%`,
      errors: this.errors,
      rateLimits: this.rateLimits,
      progressiveSearches: this.progressiveSearches,
      elapsedSeconds: elapsed.toFixed(1),
      requestsPerSecond: rps,
    };
  }

  logSummary() {
    const s = this.getSummary();
    log(`📊 === MÉTRICAS ===`);
    log(`⏱ ${s.elapsedSeconds}s | 📡 ${s.requests} (${s.requestsPerSecond}/s) | 💾 ${s.cacheHits} (${s.cacheHitRate}) | 🔄 ${s.cacheMisses} | 🔍 ${s.progressiveSearches} | ⚠ ${s.rateLimits} | ❌ ${s.errors}`);
  }
}

const metrics = new APIMetrics();
let consecutiveRateLimits = 0;
const MAX_CONSECUTIVE_RATE_LIMITS = 3;

function resetRateLimitCounter() {
  consecutiveRateLimits = 0;
}

function trackRateLimit() {
  consecutiveRateLimits++;
  if (consecutiveRateLimits >= MAX_CONSECUTIVE_RATE_LIMITS) {
    log(`🛑 ${consecutiveRateLimits} rate limits, esperando 60s...`, "WARN");
    delay(60000);
    resetRateLimitCounter();
  }
  return false;
}

// =============================================================================
// MENÚ Y FUNCIONES DE ENTRADA
// =============================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🔍 UPC Lookup v8.1.3")
    .addItem("📋 Actualizar por UPC (hoja BuscaUPCs)", "updateProductInfoFromUPCs_Safe")
    .addSeparator()
    .addItem("🔎 Buscar UPCs por NOMBRE", "buscarUPCsPorNombre_Safe")
    .addItem("🔎 Búsqueda avanzada", "buscarYCompletarInteractivoAvanzado_Safe")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("⚙ Configuración")
      .addItem("🔑 Configurar API Key", "setupAPIKey")
      .addItem("🩺 Diagnóstico", "diagnosticoSistema")
      .addItem("📊 Métricas", "mostrarMetricas")
      .addItem("🧪 Test búsqueda", "testBusquedaNombre"))
    .addSeparator()
    .addItem("🧹 Limpiar Caché", "clearUPCCache")
    .addToUi();
}

function withScriptLock(fn) {
  return function(...args) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(7000)) {
      try {
        SpreadsheetApp.getUi().alert("⏳ Proceso en curso", "Otro proceso en ejecución.");
      } catch (e) {
        log("Bloqueado por LockService", "WARN");
      }
      return;
    }
    try {
      metrics.reset();
      return fn.apply(this, args);
    } catch (e) {
      log(`Error: ${e.message}\n${e.stack}`, "ERROR");
      try {
        SpreadsheetApp.getUi().alert("❌ Error", e.message);
      } catch (uiError) {
        // UI might not be available
      }
    } finally {
      metrics.logSummary();
      lock.releaseLock();
    }
  };
}

const updateProductInfoFromUPCs_Safe = withScriptLock(updateProductInfoFromUPCs);
const buscarUPCsPorNombre_Safe = withScriptLock(buscarUPCsPorNombre);
const buscarYCompletarInteractivoAvanzado_Safe = withScriptLock(buscarYCompletarInteractivoAvanzado);

function log(msg, level = "INFO") { console.log(`[${new Date().toISOString()}] ${level}: ${msg}`); }
function delay(ms) { const MAX_MS = 5 * 60 * 1000; if (ms > MAX_MS) ms = MAX_MS; Utilities.sleep(ms); }
function getAPIKey() { const key = PropertiesService.getScriptProperties().getProperty(CONFIG_BUSCA_UPC.PROPERTY_KEY); return key && key.trim() ? key.trim() : ''; }
function getSheetByName(name) { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); if (!sheet) { log(`❌ Hoja "${name}" no encontrada.`, "ERROR"); } return sheet; }
function columnLetterToIndex(letter) { if (!letter) return null; letter = letter.toUpperCase().trim(); if (!REGEX.COLUMN_LETTER.test(letter)) return null; let col = 0; for (let i = 0; i < letter.length; i++) col = col * 26 + (letter.charCodeAt(i) - 64); return col > 0 ? col : null; }
function isValidGTIN(gtin) { const s = String(gtin).trim(); if (!REGEX.GTIN.test(s)) return false; const digits = s.split('').map(Number); const checkDigit = digits.pop(); let sum = 0; for (let i = 0; i < digits.length; i++) sum += digits[i] * ((i % 2) !== (digits.length % 2) ? 3 : 1); return checkDigit === (10 - (sum % 10)) % 10; }
function isValidUPCFormat(upc) { return upc && typeof upc === 'string' && REGEX.GTIN.test(upc.trim()); }
function isSearchableProductName(name) { if (!name || typeof name !== 'string') return false; const clean = String(name).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); if (clean.length < CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH) return false; if (REGEX.ONLY_NUMBERS_SYMBOLS.test(clean)) return false; if (REGEX.ONLY_NUMBERS.test(clean)) return false; if (REGEX.SKU_ONLY.test(clean)) return false; if (BAD_NAME_PATTERNS.some(p => p.test(clean))) return false; const letterCount = (clean.match(/[a-z]/g) || []).length; return letterCount >= 3; }

// =============================================================================
// LÓGICA DE BÚSQUEDA Y PROCESAMIENTO DE NOMBRES
// =============================================================================
const KNOWN_BRANDS=['ruger','glock','smith & wesson','s&w','sig sauer','beretta','colt','remington','winchester','browning','mossberg','savage','henry','springfield','kimber','taurus','walther','hk','fn','magpul','promag','mec-gar','wilson combat','eds','egw','leupold','vortex','bushnell','nikon','redfield','weaver','burris','holosun','eotech','aimpoint','trijicon','streamlight','surefire','olight','hogue','pachmayr','uncle mike',"uncle mike's",'blackhawk','safariland','crosman','daisy','benjamin','gamo','umarex','bear archery','mathews','hoyt','pse','bowtech','diamond','barnett','tenpoint','ravin','excalibur','killer instinct',"tink's",'code blue','wildlife research','primos','hunters specialties','thermacell','shimano','daiwa','penn','abu garcia','pflueger','ugly stik','berkley','rapala','strike king','zoom','yamamoto','sony','samsung','lg','hp','dell','lenovo','asus','acer','apple','logitech','corsair','razer','steelseries','akchoer','szco','szco supplies','cold steel','buck','gerber','kershaw','benchmade','spyderco','crkt','sog','ka-bar'];
const BRAND_REGEX_CACHE=new Map();
function getBrandRegex(brand){if(!BRAND_REGEX_CACHE.has(brand)){const escaped=brand.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');BRAND_REGEX_CACHE.set(brand,new RegExp(`(^|[\\s\\-\\.])${escaped}([\\s\\-\\.]|$)`, 'i'));}return BRAND_REGEX_CACHE.get(brand);}
function extractBrand(productName){const lower=productName.toLowerCase();for(const brand of KNOWN_BRANDS){if(getBrandRegex(brand).test(lower))return brand;}return null;}
function optimizeProductName(rawName){if(!rawName)return"";let name=String(rawName).trim();name=name.replace(REGEX.EBAY_PREFIXES,'');name=name.replace(REGEX.VARIANT_SUFFIX,'');name=name.replace(REGEX.COLOR_PARENTHESES,'');name=name.replace(REGEX.THREAD_SPECS,'');name=name.replace(REGEX.CALIBERS,'');name=name.replace(REGEX.MEASUREMENTS,'');name=name.replace(REGEX.ROUNDS_PACKS,'');name=name.replace(REGEX.COUNT_PACKS,'');name=name.replace(REGEX.MULTIPLE_SPACES,' ');name=name.replace(REGEX.DASHES,' ');name=name.replace(REGEX.TRAILING_PUNCT,'');if(name.length>CONFIG_BUSCA_UPC.MAX_SEARCH_LENGTH){const truncated=name.substring(0,CONFIG_BUSCA_UPC.MAX_SEARCH_LENGTH);const lastSpace=truncated.lastIndexOf(' ');name=lastSpace>20?truncated.substring(0,lastSpace):truncated;}return name.trim();}
function getShortProductName(rawName){const optimized=optimizeProductName(rawName);const words=optimized.split(/\s+/);const brand=extractBrand(rawName);if(brand){const brandFirstWord=brand.split(' ')[0];const brandIndex=words.findIndex(w=>w.toLowerCase().includes(brandFirstWord));if(brandIndex>=0){return words.slice(brandIndex,brandIndex+4).join(' ');}}const significant=words.filter(w=>{const lower=w.toLowerCase();return w.length>2&&!/^(the|for|and|with|new|hot|set|kit)$/i.test(lower)&&!REGEX.ONLY_NUMBERS.test(w);});return significant.slice(0,5).join(' ');}

// =============================================================================
// CACHING
// =============================================================================
function getCacheNs(){const props=PropertiesService.getScriptProperties();let ns=props.getProperty(CACHE_NS_PROP);if(!ns){ns=`v${Math.random().toString(36).substring(2,8)}`;props.setProperty(CACHE_NS_PROP,ns);}return ns;}
function makeCacheKey(kind,id){const safeId=String(id).toLowerCase().replace(/\s+/g,'_').substring(0,200);return `${getCacheNs()}:${kind}:${safeId}`;}
function clearUPCCache(){const props=PropertiesService.getScriptProperties();const newNs=`v${Math.random().toString(36).substring(2,8)}`;props.setProperty(CACHE_NS_PROP,newNs);log(`🧹 Caché invalidada: ${newNs}`);SpreadsheetApp.getActiveSpreadsheet().toast("Caché limpiada","Info",5);}
function shrinkItemForCache(item){if(!item||typeof item!=='object')return item;const copy=JSON.parse(JSON.stringify(item));if(Array.isArray(copy.images))copy.images=copy.images.slice(0,3);if(Array.isArray(copy.offers))copy.offers=copy.offers.slice(0,CONFIG_BUSCA_UPC.MAX_VENDORS);if(copy.description&&copy.description.length>500)copy.description=copy.description.substring(0,500)+"...";return copy;}

// =============================================================================
// FORMATEO DE DATOS
// =============================================================================
function flattenJSON(data,prefix=""){const result={};for(const key in data){if(!Object.prototype.hasOwnProperty.call(data,key))continue;const val=data[key];if(val===null||val===undefined){result[prefix+key]="";}else if(typeof val==="object"&&!Array.isArray(val)){Object.assign(result,flattenJSON(val,prefix+key+"."));}else if(Array.isArray(val)){result[prefix+key]=val.length?JSON.stringify(val):"[]";}else{result[prefix+key]=val;}}return result;}
function formatEmptyRow(identifier,msg="No se encontraron datos"){const row=new Array(FINAL_HEADERS.length).fill("");row[IDX.upc]=identifier;row[IDX.Error]=msg;return row;}
function formatProductData(data,upc){const flat=flattenJSON(data);const row=new Array(FINAL_HEADERS.length).fill("");FINAL_HEADERS.forEach((header,i)=>{if(header.startsWith('vendor_'))return;if(header==="upc")row[i]=upc||data.upc||"";else if(header==="offers"&&Array.isArray(data.offers))row[i]=`Count: ${data.offers.length}`;else row[i]=flat[header]||data[header]||"";});const offers=Array.isArray(data.offers)?data.offers:[];offers.slice(0,CONFIG_BUSCA_UPC.MAX_VENDORS).forEach((offer,j)=>{const base=BASE_HEADERS.length+j*7;row[base+0]=offer.merchant||"";row[base+1]=offer.price||"";row[base+2]=offer.currency||"";row[base+3]=offer.condition||"";row[base+4]=offer.domain||"";row[base+5]=offer.availability||"";row[base+6]=offer.link||"";});return row;}
function computeBackoffMs(attempt,response=null){if(response){try{const retryAfter=Number(response.getHeaders()['Retry-After']||0)*1000;if(retryAfter>0)return Math.min(retryAfter,60000);}catch(e){}}const baseDelay=CONFIG_BUSCA_UPC.INITIAL_DELAY*Math.pow(2,attempt);const jitter=baseDelay*(Math.random()*0.3);return Math.min(baseDelay+jitter,CONFIG_BUSCA_UPC.MAX_BACKOFF_MS);}
function scoreCandidate(query,candidate){const queryLower=String(query).toLowerCase();const qTokens=new Set(queryLower.split(/\s+/).filter(t=>t.length>2));const cTitle=(candidate.title||"").toLowerCase();const cBrand=(candidate.brand||"").toLowerCase();const cModel=(candidate.model||"").toLowerCase();const cTokens=new Set([...cTitle.split(/\s+/),cBrand,cModel].filter(t=>t&&t.length>2));let intersection=0;qTokens.forEach(token=>{if(cTokens.has(token))intersection++;cTokens.forEach(ct=>{if(ct.includes(token)||token.includes(ct))intersection+=0.5;});});const union=new Set([...qTokens,...cTokens]).size;const jaccard=union===0?0:intersection/union;let brandBonus=0;if(cBrand){if(queryLower.includes(cBrand))brandBonus=0.4;else if(cBrand.split(' ').some(b=>queryLower.includes(b)))brandBonus=0.2;}const modelBonus=(cModel&&queryLower.includes(cModel))?0.3:0;const upcBonus=(candidate.upc||candidate.ean)?0.2:0;return jaccard+brandBonus+modelBonus+upcBonus;}

// =============================================================================
// CLIENTE API (UPCITEMDB)
// =============================================================================
function makeAPIRequest(endpoint,paramsString){const apiKey=getAPIKey();const isFreePlan=!apiKey;const API_BASE=isFreePlan?'https://api.upcitemdb.com/prod/trial':'https://api.upcitemdb.com/prod/v1';const fullUrl=API_BASE+endpoint+'?'+paramsString;metrics.logRequest();const options={method:'get',headers:isFreePlan?{}:{'user_key':apiKey},muteHttpExceptions:true};try{const response=UrlFetchApp.fetch(fullUrl,options);return{response,code:response.getResponseCode()};}catch(error){log(`Error API: ${error.message}`,"ERROR");return{response:null,code:0};}}
function handleAPIResponse(code,response,attempt,retryCallback){if(code===200){resetRateLimitCounter();return JSON.parse(response.getContentText());}if(code===429){metrics.logRateLimit();if(trackRateLimit())delay(60000);if(attempt<CONFIG_BUSCA_UPC.MAX_RETRIES){delay(computeBackoffMs(attempt,response));return retryCallback();}}if(code>=500&&attempt<CONFIG_BUSCA_UPC.MAX_RETRIES){delay(computeBackoffMs(attempt,response));return retryCallback();}if(code===401){log(`❌ Auth 401: Verifica API key upcitemdb`,"ERROR");}else log(`❌ API ${code}`,"ERROR");metrics.logError();return null;}
function searchProductInfoByUPC(upc){const cleanUpc=String(upc||"").replace(/\D+/g,"");if(!REGEX.GTIN.test(cleanUpc)){metrics.logError();return formatEmptyRow(upc,"Formato inválido");}const cache=CacheService.getScriptCache();const cacheKey=makeCacheKey('upc',cleanUpc);const cached=cache.get(cacheKey);if(cached){metrics.logCacheHit();return formatProductData(JSON.parse(cached),cleanUpc);}metrics.logCacheMiss();for(let attempt=0;attempt<=CONFIG_BUSCA_UPC.MAX_RETRIES;attempt++){try{const {response,code}=makeAPIRequest('/lookup',`upc=${encodeURIComponent(cleanUpc)}`);if(code===200){const json=JSON.parse(response.getContentText());if(json.items?.length>0){const item=shrinkItemForCache(json.items[0]);cache.put(cacheKey,JSON.stringify(item),CONFIG_BUSCA_UPC.CACHE_DURATION);return formatProductData(item,cleanUpc);}return formatEmptyRow(upc,"Sin resultados");}const result=handleAPIResponse(code,response,attempt,()=>searchProductInfoByUPC(upc));if(result!==null)return result;if(code!==429&&code<500)return formatEmptyRow(upc,`HTTP ${code}`);}catch(e){if(attempt<CONFIG_BUSCA_UPC.MAX_RETRIES){delay(computeBackoffMs(attempt));continue;}metrics.logError();return formatEmptyRow(upc,`Error: ${e.message}`);}}return formatEmptyRow(upc,`Fallo tras ${CONFIG_BUSCA_UPC.MAX_RETRIES+1} intentos`);}
function callSearchAPI(searchQuery){for(let attempt=0;attempt<=CONFIG_BUSCA_UPC.MAX_RETRIES;attempt++){try{const {response,code}=makeAPIRequest('/search',`s=${encodeURIComponent(searchQuery)}&match_mode=0&type=product`);const result=handleAPIResponse(code,response,attempt,()=>callSearchAPI(searchQuery));if(result){const items=Array.isArray(result.items)?result.items:[];if(items.length===0)return null;let best=items[0],max=-1;items.slice(0,15).forEach(it=>{const sc=scoreCandidate(searchQuery,it);if(sc>max){max=sc;best=it;}});return best;}}catch(e){if(attempt<CONFIG_BUSCA_UPC.MAX_RETRIES){delay(computeBackoffMs(attempt));continue;}metrics.logError();return null;}}return null;}
function searchProductByNameRaw(productName,useOptimization=true){const rawName=String(productName).trim();if(rawName.length<CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH||!isSearchableProductName(rawName))return null;const searchName=useOptimization?optimizeProductName(rawName):rawName;if(searchName.length<CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH)return null;const cache=CacheService.getScriptCache();const cacheKey=makeCacheKey('name',searchName);const cached=cache.get(cacheKey);if(cached){metrics.logCacheHit();return JSON.parse(cached);}metrics.logCacheMiss();const result=callSearchAPI(searchName);if(result){cache.put(cacheKey,JSON.stringify(shrinkItemForCache(result)),CONFIG_BUSCA_UPC.CACHE_DURATION);return result;}return null;}
function smartSearchByName(productName){const rawName=String(productName).trim();if(!isSearchableProductName(rawName))return null;const optimizedName=optimizeProductName(rawName);let result=searchProductByNameRaw(optimizedName,false);if(result&&(result.upc||result.ean))return result;if(CONFIG_BUSCA_UPC.ENABLE_PROGRESSIVE_SEARCH){const shortName=getShortProductName(rawName);if(shortName&&shortName!==optimizedName&&shortName.length>=CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH){metrics.logProgressiveSearch();delay(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);result=callSearchAPI(shortName);if(result&&(result.upc||result.ean)){CacheService.getScriptCache().put(makeCacheKey('name',optimizedName),JSON.stringify(shrinkItemForCache(result)),CONFIG_BUSCA_UPC.CACHE_DURATION);return result;}}const brand=extractBrand(rawName);if(brand&&brand.length>=3){const words=optimizedName.split(/\s+/);const brandFirstWord=brand.split(' ')[0];const brandIdx=words.findIndex(w=>w.toLowerCase().includes(brandFirstWord));if(brandIdx>=0&&words[brandIdx+1]){const brandPlusOne=`${brand} ${words[brandIdx+1]}`.trim();if(brandPlusOne!==shortName&&brandPlusOne.length>=5){metrics.logProgressiveSearch();delay(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);result=callSearchAPI(brandPlusOne);if(result&&(result.upc||result.ean)){CacheService.getScriptCache().put(makeCacheKey('name',optimizedName),JSON.stringify(shrinkItemForCache(result)),CONFIG_BUSCA_UPC.CACHE_DURATION);return result;}}}}log(`❌ No encontrado: "${rawName.substring(0,50)}..."`);return null;}
function searchProductByName(productName){const result=smartSearchByName(productName);return result?formatProductData(result,result.upc||result.ean||""):formatEmptyRow("","Sin resultados");}

// =============================================================================
// FUNCIONES INTERACTIVAS (UI)
// =============================================================================
function promptUserInput(ui,title,message){const response=ui.prompt(title,message,ui.ButtonSet.OK_CANCEL);return{confirmed:response.getSelectedButton()===ui.Button.OK,value:response.getResponseText().trim()};}
function processRow(sheet,currentRow,productName,destCol,currentDest){if(!productName)return{skipped:true};if(currentDest&&currentDest!=="NO_ENCONTRADO"&&currentDest!=="ERROR"){return{skipped:true};}if(!isSearchableProductName(productName))return{skipped:true,invalid:true};try{const item=smartSearchByName(productName);if(item&&(item.upc||item.ean)){sheet.getRange(currentRow,destCol).setValue(String(item.upc||item.ean).trim());return{success:true};}else{sheet.getRange(currentRow,destCol).setValue("NO_ENCONTRADO");return{notFound:true};}}catch(err){sheet.getRange(currentRow,destCol).setValue("ERROR");return{error:true};}}
function buscarUPCsPorNombre(){const ui=SpreadsheetApp.getUi();const ss=SpreadsheetApp.getActive();const sheets=ss.getSheets().map(s=>s.getName());const sheetList="Selecciona número:\n\n"+sheets.map((n,i)=>`${i+1}. ${n}`).join('\n');const sheetInput=promptUserInput(ui,"📑 Hoja",sheetList);if(!sheetInput.confirmed)return;const sheetIndex=Number(sheetInput.value);if(!sheetIndex||sheetIndex<1||sheetIndex>sheets.length)return ui.alert("Error","Número inválido");const sheet=ss.getSheetByName(sheets[sheetIndex-1]);const originInput=promptUserInput(ui,"📝 Origen","Letra columna NOMBRE (ej: B):");if(!originInput.confirmed)return;const originCol=columnLetterToIndex(originInput.value);if(!originCol)return ui.alert("Error","Columna origen inválida");const destInput=promptUserInput(ui,"🎯 Destino","Letra columna UPC (ej: C):");if(!destInput.confirmed)return;const destCol=columnLetterToIndex(destInput.value);if(!destCol)return ui.alert("Error","Columna destino inválida");const startInput=promptUserInput(ui,"📍 Fila inicial","¿Desde qué fila? (ej: 2):");if(!startInput.confirmed)return;const startRow=Number(startInput.value);if(!startRow||startRow<1)return ui.alert("Error","Fila inválida");const lastRow=sheet.getLastRow();if(lastRow<startRow)return ui.alert("Aviso","No hay datos");const numRows=lastRow-startRow+1;const confirmResp=ui.alert("Confirmar",`Hoja: ${sheets[sheetIndex-1]}\nOrigen: ${originInput.value.toUpperCase()}\nDestino: ${destInput.value.toUpperCase()}\nFilas: ${startRow}-${lastRow} (${numRows})\n¿Continuar?`,ui.ButtonSet.YES_NO);if(confirmResp!==ui.Button.YES)return;SpreadsheetApp.getActiveSpreadsheet().toast("Iniciado","🔍 UPC",10);const maxCol=Math.max(originCol,destCol);const data=sheet.getRange(startRow,1,numRows,maxCol).getValues();let procesados=0,encontrados=0,noEncontrados=0,saltados=0,errores=0;for(let i=0;i<data.length;i++){const currentRow=startRow+i;const productName=String(data[i][originCol-1]||"").trim();const currentDest=String(data[i][destCol-1]||"").trim();const result=processRow(sheet,currentRow,productName,destCol,currentDest);if(result.skipped){if(result.invalid)saltados++;continue;}procesados++;if(result.success)encontrados++;else if(result.notFound)noEncontrados++;else if(result.error)errores++;if(i<data.length-1)delay(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);if(procesados%5===0)SpreadsheetApp.flush();}const summary=`📊 RESUMEN\nTiempo: procesados ${procesados}\n✅ Encontrados: ${encontrados}\n❌ No: ${noEncontrados}\n⏭ Saltados: ${saltados}\n🔴 Errores: ${errores}`;ui.alert("Terminado",summary,ui.ButtonSet.OK);}

// =============================================================================
// FUNCIONES DE HOJA DE CÁLCULO (TRIGGERS Y PROCESOS)
// =============================================================================
function updateProductInfoFromUPCs_onEdit_(e){try{const hoja=e.range.getSheet();const fila=e.range.getRow();const col=e.range.getColumn();if(fila<CONFIG_BUSCA_UPC.START_ROW)return;if(col!==CONFIG_BUSCA_UPC.UPC_COLUMN)return;const upc=String(e.value||'').trim();if(!upc||!isValidGTIN(upc))return;const data=searchProductInfoByUPC(upc);if(data&&data[IDX.Error]===''){hoja.getRange(fila,CONFIG_BUSCA_UPC.UPC_COLUMN+1,1,data.length-1).setValues([data.slice(1)]);}}catch(err){log(`onEdit UPC: ${err.message}`);}}
function updateProductInfoFromUPCs(){const sheet=getSheetByName(CONFIG_BUSCA_UPC.SHEET_NAME);if(!sheet)return;const lastRow=sheet.getLastRow();if(lastRow<CONFIG_BUSCA_UPC.START_ROW)return;const numRows=lastRow-CONFIG_BUSCA_UPC.START_ROW+1;const upcData=sheet.getRange(CONFIG_BUSCA_UPC.START_ROW,CONFIG_BUSCA_UPC.UPC_COLUMN,numRows,1).getValues();let procesados=0;for(let i=0;i<upcData.length;i++){const upc=String(upcData[i][0]||'').trim();if(!upc)continue;const fila=CONFIG_BUSCA_UPC.START_ROW+i;const data=searchProductInfoByUPC(upc);if(data){sheet.getRange(fila,CONFIG_BUSCA_UPC.UPC_COLUMN,1,data.length).setValues([data]);procesados++;}if(procesados%5===0)delay(CONFIG_BUSCA_UPC.BATCH_PAUSE_MS);}log(`Actualización: ${procesados} UPCs`);}
function buscarYCompletarInteractivoAvanzado(){SpreadsheetApp.getUi().alert("Avanzado","Usa 'Buscar UPCs por NOMBRE' por ahora.");}
function setupAPIKey(){const ui=SpreadsheetApp.getUi();const result=ui.prompt('API Key UPCItemDB','Pega tu key (vacío = gratis):',ui.ButtonSet.OK_CANCEL);if(result.getSelectedButton()===ui.Button.OK){PropertiesService.getScriptProperties().setProperty(CONFIG_BUSCA_UPC.PROPERTY_KEY,result.getResponseText().trim());ui.alert('Guardada','Key configurada');}}
function diagnosticoSistema(){const ui=SpreadsheetApp.getUi();const apiKey=getAPIKey();ui.alert('Diagnóstico',`API: ${apiKey?'Configurada':'Gratis'}\nHoja: ${CONFIG_BUSCA_UPC.SHEET_NAME}\nCache: ${getCacheNs()}`);}
function mostrarMetricas(){const s=metrics.getSummary();SpreadsheetApp.getUi().alert('Métricas',`Tiempo: ${s.elapsedSeconds}s\nRequests: ${s.requests}\nCache: ${s.cacheHits} (${s.cacheHitRate})\nProgresivas: ${s.progressiveSearches}\nRate limits: ${s.rateLimits}\nErrores: ${s.errors}`);}
function testBusquedaNombre(){const ui=SpreadsheetApp.getUi();const result=ui.prompt('Test','Nombre producto:',ui.ButtonSet.OK_CANCEL);if(result.getSelectedButton()===ui.Button.OK){const item=smartSearchByName(result.getResponseText().trim());ui.alert('Resultado',item?`UPC: ${item.upc||item.ean}\n${item.title}`:'No encontrado');}}

// =============================================================================
// FUNCIÓN ESPECÍFICA (obsoleto?)
// =============================================================================

const CONFIG_HOJA_ACTUAL = {
  START_ROW: 7,
  PRODUCT_COL: 3,
  DEST_COL: 15,
  HEADER_ROW: 6,
  HEADER_TEXT: "UPC ENCONTRADO"
};

function obtenerUPCsParaHojaActual(){
  const ui=SpreadsheetApp.getUi();
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow=sheet.getLastRow();
  
  if(lastRow < FILA_INICIO) {
    ui.alert("Sin datos");
    return; // ✅ Return vacío, sin código después
  }
  
  // ✅ Este código SOLO se ejecuta si NO hay return
  sheet.getRange(CONFIG_HOJA_ACTUAL.HEADER_ROW, CONFIG_HOJA_ACTUAL.DEST_COL).setValue(CONFIG_HOJA_ACTUAL.HEADER_TEXT).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
  const numRows=lastRow - CONFIG_HOJA_ACTUAL.START_ROW + 1;
  const nombres=sheet.getRange(CONFIG_HOJA_ACTUAL.START_ROW, CONFIG_HOJA_ACTUAL.PRODUCT_COL, numRows, 1).getValues();
  const resultados=[];
  SpreadsheetApp.getActive().toast("Buscando UPCs...","🔍");
  
  for(let i=0;i<nombres.length;i++){
    const nombre=nombres[i][0];
    if(nombre && String(nombre).trim() !== ""){
      try{
        const info=smartSearchByName(nombre);
        resultados.push([info ? (info.upc||info.ean||"NO ENCONTRADO") : "NO ENCONTRADO"]);
      }catch(e){
        resultados.push(["ERROR API"]);
      }
    } else {
      resultados.push([""]);
    }
    Utilities.sleep(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);
    if(i % 5 === 0) SpreadsheetApp.getActive().toast(`Procesando: ${i+1} de ${numRows}`,"🔍");
  }
  
  sheet.getRange(CONFIG_HOJA_ACTUAL.START_ROW, CONFIG_HOJA_ACTUAL.DEST_COL, resultados.length, 1).setValues(resultados);
  ui.alert("Completado","UPCs en columna O");
} // ✅ LLAVE DE CIERRE CORRECTA


// =============================================================================
// ZLOOKUP FAMILIA - DEPURADO PARA CUSTOM FUNCTION (sin LanguageApp en fórmula)
// =============================================================================
function translateToSpanishSafe(text){
  if(!text || /^\d+$/.test(String(text).replace(/\s/g,''))) return null;
  if(isCustomFunctionCall()) return null; // No traducir en =ZLOOKUP()
  try{
    const cache=CacheService.getScriptCache();
    const key='tr_es_'+text.substring(0,150);
    const cached=cache.get(key);
    if(cached) return cached==='__null__'?null:cached;
    const tr=LanguageApp.translate(text,'','es');
    const result=(tr&&tr.toLowerCase()!==text.toLowerCase())?tr:null;
    cache.put(key,result===null?'__null__':result,21600);
    return result;
  }catch(e){return null;}
}
function norm(s){return String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9\s]+/g,'').replace(/\s+/g,' ').trim();}
function countCommonTokens(a,b){const setA=new Set(a.split(' '));let c=0;for(const t of b.split(' '))if(t.length>2&&setA.has(t))c++;return c;}
function similarityFast(a,b){if(a===b)return 100;if(!a||!b)return 0;const maxLen=Math.max(a.length,b.length);if(maxLen>80){a=a.slice(0,80);b=b.slice(0,80);}return ((maxLen-levenshtein2Rows(a,b))/maxLen)*100;}
function levenshtein2Rows(a,b){if(a.length<b.length){let t=a;a=b;b=t;}if(b.length===0)return a.length;let prev=new Array(b.length+1),curr=new Array(b.length+1);for(let j=0;j<=b.length;j++)prev[j]=j;for(let i=1;i<=a.length;i++){curr[0]=i;const ca=a.charAt(i-1);for(let j=1;j<=b.length;j++){const cost=ca===b.charAt(j-1)?0:1;curr[j]=Math.min(prev[j]+1,curr[j-1]+1,prev[j-1]+cost);}const t=prev;prev=curr;curr=t;}return prev[b.length];}

function FIND_PRODUCT_CODE(value, lookupRange, returnCol, threshold){
  const eff=threshold!=null?threshold:75;
  const result=ZLOOKUP(value, lookupRange, returnCol, eff, null);
  if(result && result!=="" && result!=="No encontrado") return result;
  return `Not found. Try '🔍 UPC Lookup > 🔎 Buscar UPCs por NOMBRE'`;
}

function ZLOOKUP(value, lookupRange, returnCol, threshold, boostRange){
  try{
    if(value==null||value==="") return "No encontrado";
    value=String(value).trim();if(!value) return "No encontrado";
    threshold=threshold!=null?Number(threshold):65; returnCol=Number(returnCol)||2;
    if(!lookupRange||lookupRange.length===0) return "No encontrado";
    const searchNorm=norm(value);
    if(!searchNorm) return "No encontrado";
    const searchTokens=new Set(searchNorm.split(' ').filter(t=>t.length>1));
    let translated=null, searchNormTranslated=null, translatedTokens=null;
    if(!isCustomFunctionCall()){
      translated=translateToSpanishSafe(value);
      if(translated){searchNormTranslated=norm(translated); translatedTokens=new Set(searchNormTranslated.split(' ').filter(t=>t.length>1));}
    }
    for(let i=0;i<lookupRange.length;i++){
      if(!lookupRange[i]||!lookupRange[i][0]) continue;
      const n=norm(String(lookupRange[i][0]));
      if(n===searchNorm || (searchNormTranslated&&n===searchNormTranslated)) return lookupRange[i][returnCol-1]||"No encontrado";
    }
    if(threshold>=100) return "No encontrado";
    let bestScore=-1,bestIdx=-1;
    for(let i=0;i<lookupRange.length;i++){
      const cell=lookupRange[i]?lookupRange[i][0]:null; if(!cell) continue;
      const cellNorm=norm(String(cell)); if(!cellNorm) continue;
      if(searchNorm.length>5&&cellNorm.length>5){
        let hasOverlap=false;
        for(const t of searchTokens){if(t.length>=3&&cellNorm.includes(t)){hasOverlap=true;break;}}
        if(!hasOverlap&&translatedTokens){for(const t of translatedTokens){if(t.length>=3&&cellNorm.includes(t)){hasOverlap=true;break;}}}
        if(!hasOverlap&&Math.abs(cellNorm.length-searchNorm.length)>searchNorm.length*0.8) continue;
      }
      const scoreOriginal=similarityFast(searchNorm,cellNorm);
      const scoreTranslated=searchNormTranslated?similarityFast(searchNormTranslated,cellNorm):0;
      let score=Math.max(scoreOriginal,scoreTranslated);
      let boost=countCommonTokens(searchNorm,cellNorm)*5;
      if(boostRange&&boostRange[i]&&boostRange[i][0]!=''){boost+=parseFloat(boostRange[i][0])||0;}
      const finalScore=Math.min(score+boost,100);
      if(finalScore>bestScore){bestScore=finalScore;bestIdx=i;if(finalScore===100) break;}
    }
    if(bestIdx!==-1&&bestScore>=threshold){return lookupRange[bestIdx][returnCol-1]||"No encontrado";}
    return "No encontrado";
  }catch(e){return "No encontrado";}
}

function ZLOOKUP_BATCH(searchRange, lookupRange, returnCol, threshold){
  if(!Array.isArray(searchRange)) searchRange=[[searchRange]];
  threshold=threshold!=null?Number(threshold):65; 
  returnCol=Number(returnCol)||2;
  let last=searchRange.length-1; 
  while(last>=0 && !searchRange[last][0]) last--; 
  if(last<0) return [[""]];
  
  searchRange=searchRange.slice(0,last+1);
  
  const lookupPre=[];
  for(let i=0;i<lookupRange.length;i++){
    if(!lookupRange[i]||!lookupRange[i][0]) continue;
    const n=norm(String(lookupRange[i][0]));
    if(!n) continue;
    lookupPre.push({n:n, idx:i, ret:lookupRange[i][returnCol-1]||"No encontrado"});
  }
  
  const exactMap=new Map(); 
  for(const item of lookupPre){
    if(!exactMap.has(item.n)) exactMap.set(item.n,item);
  }
  
  const out=[];
  for(let r=0;r<searchRange.length;r++){
    const val=searchRange[r][0]; 
    if(!val){
      out.push([""]);
      continue;
    }
    
    const searchNorm=norm(String(val));
    const searchTokens = new Set(searchNorm.split(' ').filter(t => t.length > 1));
    let searchNormTranslated=null;
    let translatedTokens = null;
    
    if(!isCustomFunctionCall()){
      const tr=translateToSpanishSafe(String(val)); 
      if(tr) {
        searchNormTranslated=norm(tr);
        translatedTokens = new Set(searchNormTranslated.split(' ').filter(t => t.length > 1));
      }
    }
    
    if(exactMap.has(searchNorm)){
      out.push([exactMap.get(searchNorm).ret]);
      continue;
    }
    if(searchNormTranslated && exactMap.has(searchNormTranslated)){
      out.push([exactMap.get(searchNormTranslated).ret]);
      continue;
    }
    
    let bestScore=-1;
    let bestRet="No encontrado";
    
    for(const item of lookupPre){
      const cellNorm = item.n;
      
      // 🔥 FILTRO RÁPIDO DE TOKENS
      if(searchNorm.length > 5 && cellNorm.length > 5){
        let hasOverlap = false;
        for(const t of searchTokens){
          if(t.length >= 3 && cellNorm.includes(t)){
            hasOverlap = true;
            break;
          }
        }
        if(!hasOverlap && translatedTokens){
          for(const t of translatedTokens){
            if(t.length >= 3 && cellNorm.includes(t)){
              hasOverlap = true;
              break;
            }
          }
        }
        if(!hasOverlap && Math.abs(cellNorm.length - searchNorm.length) > searchNorm.length * 0.8){
          continue;
        }
      }
      
      const s1=similarityFast(searchNorm, item.n);
      const s2=searchNormTranslated ? similarityFast(searchNormTranslated, item.n) : 0;
      const score=Math.max(s1, s2);
      
      if(score > bestScore){
        bestScore = score;
        bestRet = item.ret;
        if(score === 100) break;
      }
    }
    
    out.push([bestScore >= threshold ? bestRet : "No encontrado"]);
  }
  
  return out;
} // ← ✅ ESTA LLAVE CIERRA LA FUNCIÓN ZLOOKUP_BATCH
