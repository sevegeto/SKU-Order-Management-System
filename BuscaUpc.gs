// =============================================================================
// UPC Lookup Service - Versión Mejorada
// Versión: 8.1.3
// Última actualización: enero 2026
// =============================================================================
//
// Incorpora soporte para plan gratuito (trial) basado en resultados de búsqueda web.
// Para más info sobre planes y API:
// - upcitemdb.com/api
// - devs.upcitemdb.com/
// - www.upcitemdb.com/wp/docs/main/bulk-lookup/google-sheets-add-on/

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

/**
 * ═══════════════════════════════════════════
 * TRIGGER UNIVERSAL onEdit - Deriva según la hoja editada
 * ═══════════════════════════════════════════════════════════════════════════
 */
function onEdit(e) {
    try {
        const hoja = e.range.getSheet();
        const nombreHoja = hoja.getName();

        // === 1. DERIVAR A RASTREO SHIPPO ===
        if (nombreHoja === CONFIG_COMPRAS.sheets.compras) {
            onEditShippo_(e);
            return;
        }

        // === 2. DERIVAR A BÚSQUEDA UPC ===
        if (nombreHoja === CONFIG_BUSCA_UPC.SHEET_NAME) {
            updateProductInfoFromUPCs_onEdit_(e);
            return;
        }

    } catch (err) {
        Logger.log(`Error en onEdit universal: ${err.message}`);
    }
}

/**
 * ═══════════════════════════════════════════
 * CONSTANTES Y REGEX PRE-COMPILADOS
 * ═══════════════════════════════════════════
 */
const REGEX = {
  GTIN: /^\d{12,14}$/,
  COLUMN_LETTER: /^[A-Z]+$/,
  ONLY_NUMBERS_SYMBOLS: /^[\d\s\-\.\#\(\)\/\*\&]+$/,
  ONLY_NUMBERS: /^\d+$/,
  SKU_ONLY: /^sku[\s\-:]?\d+$/i,
  EBAY_PREFIXES: /^(US!|NEW!|HOT!|SALE!|BEST!|\*+|#+)\s*/gi,
  VARIANT_SUFFIX: /,\s*(for|fits|compatible|w\/|with)\s+.*/gi,
  COLOR_PARENTHESES: /\s*\([^)]*(?:version|color|black|white|grey|gray|blue|red|green|camo|fde|od|tan)\s*[^)]*\)\s*$/gi,
  THREAD_SPECS: /\b\d+\/\d+x\d+\b/g,
  CALIBERS: /\b\d{2,3}(BLK|WIN|WSM|REM|ACP|AUTO|MAG|S&W|SPECIAL)\b/gi,
  MEASUREMENTS: /\b\d+(\.\d+)?["']?\s*(inch|in|inches|mm|cm|oz|lbs?|rd)\b/gi,
  ROUNDS_PACKS: /\b\d+\s*\/\s*rd\b/gi,
  COUNT_PACKS: /\b\d+\s*-?\s*(pack|pc|pcs|piece|count)\b/gi,
  MULTIPLE_SPACES: /\s+/g,
  DASHES: /\s*[-–—]\s*/g,
  TRAILING_PUNCT: /\s*[,;]\s*$/
};

const BAD_NAME_PATTERNS = [
  /\b(shipping|handling|delivery|freight|postage|courier|mail)\b/,
  /\b(fedex|usps|ups|dhl)\b/,
  /\b(fee|charge|adjustment|discount|credit|refund|tax|insurance)\b/,
  /\b(test|prueba|ejemplo|sample|demo)\b/,
  /^(n\/a|tbd|pending|various|none|null|unknown)$/,
  /\b(free\s+gift|promo\s+item|bonus)\b/
];

/**
 * ═══════════════════════════════════════════
 * CABECERAS Y ÍNDICES
 * ═══════════════════════════════════════════
 */
const BASE_HEADERS = [
  "upc", "ean", "title", "description", "brand", "model", "color", "size",
  "dimension", "weight", "category", "currency", "lowest_recorded_price",
  "highest_recorded_price", "images", "offers", "asin", "elid", "Error"
];

function generateVendorHeaders(max) {
  const headers = [];
  for (let i = 1; i <= max; i++) {
    headers.push(
      `vendor_${i}_merchant`, `vendor_${i}_price`, `vendor_${i}_currency`,
      `vendor_${i}_condition`, `vendor_${i}_domain`, `vendor_${i}_availability`,
      `vendor_${i}_link`
    );
  }
  return headers;
}

const VENDOR_HEADERS = generateVendorHeaders(CONFIG_BUSCA_UPC.MAX_VENDORS);
const FINAL_HEADERS = [...BASE_HEADERS,...VENDOR_HEADERS];
const IDX = FINAL_HEADERS.reduce((map, header, i) => (map[header] = i, map), {});
const CACHE_NS_PROP = 'UPC_CACHE_NS_V4';

/**
 * ═══════════════════════════════════════════
 * CLASE DE MÉTRICAS
 * ═══════════════════════════════════════════════════════════════════════════
 */
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
    const cacheRate = total > 0? ((this.cacheHits / total) * 100).toFixed(1) : "0.0";
    const rps = elapsed > 0? (this.requests / elapsed).toFixed(2) : "0.00";

    return {
      requests: this.requests,
      cacheHits: this.cacheHits,
      cacheMisses: this.cacheMisses,
      cacheHitRate: `${cacheRate}%`,
      errors: this.errors,
      rateLimits: this.rateLimits,
      progressiveSearches: this.progressiveSearches,
      elapsedSeconds: elapsed.toFixed(1),
      requestsPerSecond: rps
    };
  }

  logSummary() {
    const s = this.getSummary();
    log(`📊 === MÉTRICAS DE EJECUCIÓN ===`);
    log(`⏱️ Tiempo: ${s.elapsedSeconds}s`);
    log(`📡 Requests API: ${s.requests} (${s.requestsPerSecond}/s)`);
    log(`💾 Cache hits: ${s.cacheHits} (${s.cacheHitRate})`);
    log(`🔄 Cache misses: ${s.cacheMisses}`);
    log(`🔍 Búsquedas progresivas: ${s.progressiveSearches}`);
    log(`⚠️ Rate limits: ${s.rateLimits}`);
    log(`❌ Errores: ${s.errors}`);
    log(`================================`);
  }
}

const metrics = new APIMetrics();

/**
 * ═══════════════════════════════════════════
 * RATE LIMIT TRACKING
 * ═══════════════════════════════════════════
 */
let consecutiveRateLimits = 0;
const MAX_CONSECUTIVE_RATE_LIMITS = 3;

function resetRateLimitCounter() {
  consecutiveRateLimits = 0;
}

function trackRateLimit() {
  consecutiveRateLimits++;
  if (consecutiveRateLimits >= MAX_CONSECUTIVE_RATE_LIMITS) {
    log(`🛑 ${consecutiveRateLimits} rate limits consecutivos. Esperando 60 segundos...`, "WARN");
    delay(60000);
    resetRateLimitCounter();
  }
  return false;
}

/**
 * ═══════════════════════════════════════════
 * MENÚ PERSONALIZADO
 * ═══════════════════════════════════════════════════════════════════════════
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🔍 UPC Lookup v8.1.3")
  .addItem("📋 Actualizar por UPC (hoja BuscaUPCs)", "updateProductInfoFromUPCs_Safe")
  .addSeparator()
  .addItem("🔎 Buscar UPCs por NOMBRE (recomendado)", "buscarUPCsPorNombre_Safe")
  .addItem("🔎 Búsqueda avanzada (modo/columnas)", "buscarYCompletarInteractivoAvanzado_Safe")
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu("⚙️ Configuración")
    .addItem("🔑 Configurar API Key", "setupAPIKey")
    .addItem("🩺 Diagnóstico del sistema", "diagnosticoSistema")
    .addItem("📊 Ver métricas de sesión", "mostrarMetricas")
    .addItem("🧪 Test búsqueda por nombre", "testBusquedaNombre"))
  .addSeparator()
  .addItem("🧹 Limpiar Caché", "clearUPCCache")
  .addToUi();
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * LOCK SERVICE
 * ═══════════════════════════════════════════
 */
function withScriptLock(fn) {
  return function(...args) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(7000)) {
      try {
        SpreadsheetApp.getUi().alert(
          "⏳ Proceso en curso",
          "Otro proceso ya está en ejecución. Espera a que termine.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (e) {
        log("Ejecución bloqueada por LockService (sin UI).", "WARN");
      }
      return;
    }
    try {
      metrics.reset();
      return fn.apply(this, args);
    } catch (e) {
      log(`Error no capturado: ${e.message}\n${e.stack}`, "ERROR");
      try {
        SpreadsheetApp.getUi().alert("❌ Error", `Error inesperado: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (uiError) {
        log("No se pudo mostrar error en UI", "WARN");
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

/**
 * ═══════════════════════════════════════════
 * UTILIDADES BÁSICAS
 * ═══════════════════════════════════════════
 */
function log(msg, level = "INFO") {
  console.log(`[${new Date().toISOString()}] ${level}: ${msg}`);
}

function delay(ms) {
  const MAX_MS = 5 * 60 * 1000;
  if (ms > MAX_MS) {
    log(`⚠️ delay limitado a ${MAX_MS}ms`, "WARN");
    ms = MAX_MS;
  }
  Utilities.sleep(ms);
}

function getAPIKey() {
  const key = PropertiesService.getScriptProperties().getProperty(CONFIG_BUSCA_UPC.PROPERTY_KEY);
  return key && key.trim()? key.trim() : '';
}

function getSheetByName(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    log(`❌ Hoja "${name}" no encontrada.`, "ERROR");
    try {
      SpreadsheetApp.getUi().alert("❌ Error", `Hoja "${name}" no encontrada.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
  }
  return sheet;
}

function columnLetterToIndex(letter) {
  if (!letter) return null;
  letter = letter.toUpperCase().trim();
  if (!REGEX.COLUMN_LETTER.test(letter)) return null;

  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col > 0? col : null;
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * VALIDACIONES
 * ═══════════════════════════════════════════
 */
function isValidGTIN(gtin) {
  const s = String(gtin).trim();
  if (!REGEX.GTIN.test(s)) return false;

  const digits = s.split('').map(Number);
  const checkDigit = digits.pop();
  let sum = 0;

  for (let i = 0; i < digits.length; i++) {
    sum += digits[i] * ((i % 2)!== (digits.length % 2)? 3 : 1);
  }

  return checkDigit === (10 - (sum % 10)) % 10;
}

function isValidUPCFormat(upc) {
  return upc && typeof upc === 'string' && REGEX.GTIN.test(upc.trim());
}

function isSearchableProductName(name) {
  if (!name || typeof name!== 'string') return false;

  const clean = String(name)
  .trim()
  .toLowerCase()
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "");

  if (clean.length < CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH) return false;
  if (REGEX.ONLY_NUMBERS_SYMBOLS.test(clean)) return false;
  if (REGEX.ONLY_NUMBERS.test(clean)) return false;
  if (REGEX.SKU_ONLY.test(clean)) return false;
  if (BAD_NAME_PATTERNS.some(pattern => pattern.test(clean))) return false;

  const letterCount = (clean.match(/[a-z]/g) || []).length;
  return letterCount >= 3;
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * MARCAS CONOCIDAS Y OPTIMIZACIÓN DE NOMBRES
 * ═══════════════════════════════════════════════════════════════════════════
 */
const KNOWN_BRANDS = [
  'ruger', 'glock', 'smith & wesson', 's&w', 'sig sauer', 'beretta', 'colt',
  'remington', 'winchester', 'browning', 'mossberg', 'savage', 'henry',
  'springfield', 'kimber', 'taurus', 'walther', 'hk', 'fn', 'magpul',
  'promag', 'mec-gar', 'wilson combat', 'eds', 'egw', 'leupold', 'vortex',
  'bushnell', 'nikon', 'redfield', 'weaver', 'burris', 'holosun', 'eotech',
  'aimpoint', 'trijicon', 'streamlight', 'surefire', 'olight', 'hogue',
  'pachmayr', 'uncle mike', "uncle mike's", 'blackhawk', 'safariland',
  'crosman', 'daisy', 'benjamin', 'gamo', 'umarex', 'bear archery',
  'mathews', 'hoyt', 'pse', 'bowtech', 'diamond', 'barnett', 'tenpoint',
  'ravin', 'excalibur', 'killer instinct', "tink's", 'code blue',
  'wildlife research', 'primos', 'hunters specialties', 'thermacell',
  'shimano', 'daiwa', 'penn', 'abu garcia', 'pflueger', 'ugly stik',
  'berkley', 'rapala', 'strike king', 'zoom', 'yamamoto',
  'sony', 'samsung', 'lg', 'hp', 'dell', 'lenovo', 'asus', 'acer',
  'apple', 'logitech', 'corsair', 'razer', 'steelseries',
  'akchoer', 'szco', 'szco supplies', 'cold steel', 'buck', 'gerber',
  'kershaw', 'benchmade', 'spyderco', 'crkt', 'sog', 'ka-bar'
];

const BRAND_REGEX_CACHE = new Map();

function getBrandRegex(brand) {
  if (!BRAND_REGEX_CACHE.has(brand)) {
    const escaped = brand.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    BRAND_REGEX_CACHE.set(brand, new RegExp(`(^|[\\s\\-\\.])${escaped}([\\s\\-\\.]|$)`, 'i'));
  }
  return BRAND_REGEX_CACHE.get(brand);
}

function extractBrand(productName) {
  const lower = productName.toLowerCase();

  for (const brand of KNOWN_BRANDS) {
    if (getBrandRegex(brand).test(lower)) {
      return brand;
    }
  }

  return null;
}

function optimizeProductName(rawName) {
  if (!rawName) return "";

  let name = String(rawName).trim();

  name = name.replace(REGEX.EBAY_PREFIXES, '');
  name = name.replace(REGEX.VARIANT_SUFFIX, '');
  name = name.replace(REGEX.COLOR_PARENTHESES, '');
  name = name.replace(REGEX.THREAD_SPECS, '');
  name = name.replace(REGEX.CALIBERS, '');
  name = name.replace(REGEX.MEASUREMENTS, '');
  name = name.replace(REGEX.ROUNDS_PACKS, '');
  name = name.replace(REGEX.COUNT_PACKS, '');
  name = name.replace(REGEX.MULTIPLE_SPACES, ' ');
  name = name.replace(REGEX.DASHES, ' ');
  name = name.replace(REGEX.TRAILING_PUNCT, '');

  if (name.length > CONFIG_BUSCA_UPC.MAX_SEARCH_LENGTH) {
    const truncated = name.substring(0, CONFIG_BUSCA_UPC.MAX_SEARCH_LENGTH);
    const lastSpace = truncated.lastIndexOf(' ');
    name = lastSpace > 20? truncated.substring(0, lastSpace) : truncated;
  }

  return name.trim();
}

function getShortProductName(rawName) {
  const optimized = optimizeProductName(rawName);
  const words = optimized.split(/\s+/);
  const brand = extractBrand(rawName);

  if (brand) {
    const brandFirstWord = brand.split(' ')[0];
    const brandIndex = words.findIndex(w => w.toLowerCase().includes(brandFirstWord));
    if (brandIndex >= 0) {
      return words.slice(brandIndex, brandIndex + 4).join(' ');
    }
  }

  const significant = words.filter(w => {
    const lower = w.toLowerCase();
    return w.length > 2 &&
         !/^(the|for|and|with|new|hot|set|kit)$/i.test(lower) &&
         !REGEX.ONLY_NUMBERS.test(w);
  });

  return significant.slice(0, 5).join(' ');
}

/**
 * ═══════════════════════════════════════════
 * UTILIDADES DE CACHÉ
 * ═══════════════════════════════════════════════════════════════════════════
 */
function getCacheNs() {
  const props = PropertiesService.getScriptProperties();
  let ns = props.getProperty(CACHE_NS_PROP);
  if (!ns) {
    ns = `v${Math.random().toString(36).substring(2, 8)}`;
    props.setProperty(CACHE_NS_PROP, ns);
  }
  return ns;
}

function makeCacheKey(kind, id) {
  const safeId = String(id).toLowerCase().replace(/\s+/g, '_').substring(0, 200);
  return `${getCacheNs()}:${kind}:${safeId}`;
}

function clearUPCCache() {
  const props = PropertiesService.getScriptProperties();
  const newNs = `v${Math.random().toString(36).substring(2, 8)}`;
  props.setProperty(CACHE_NS_PROP, newNs);
  log(`🧹 Caché invalidada. Nuevo namespace: ${newNs}`);
  SpreadsheetApp.getActiveSpreadsheet().toast("Caché limpiada con éxito.", "Info", 5);
}

function shrinkItemForCache(item) {
  if (!item || typeof item!== 'object') return item;
  const copy = JSON.parse(JSON.stringify(item));
  if (Array.isArray(copy.images)) copy.images = copy.images.slice(0, 3);
  if (Array.isArray(copy.offers)) copy.offers = copy.offers.slice(0, CONFIG_BUSCA_UPC.MAX_VENDORS);
  if (copy.description && copy.description.length > 500) {
    copy.description = copy.description.substring(0, 500) + "...";
  }
  return copy;
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * UTILIDADES DE FORMATO
 * ═══════════════════════════════════════════
 */
function flattenJSON(data, prefix = "") {
  const result = {};
  for (const key in data) {
    if (!Object.prototype.hasOwnProperty.call(data, key)) continue;
    const val = data[key];
    if (val === null || val === undefined) {
      result[prefix + key] = "";
    } else if (typeof val === "object" &&!Array.isArray(val)) {
      Object.assign(result, flattenJSON(val, prefix + key + "."));
    } else if (Array.isArray(val)) {
      result[prefix + key] = val.length? JSON.stringify(val) : "[]";
    } else {
      result[prefix + key] = val;
    }
  }
  return result;
}

function formatEmptyRow(identifier, msg = "No se encontraron datos") {
  const row = new Array(FINAL_HEADERS.length).fill("");
  row[IDX.upc] = identifier;
  row[IDX.Error] = msg;
  return row;
}

function formatProductData(data, upc) {
  const flat = flattenJSON(data);
  const row = new Array(FINAL_HEADERS.length).fill("");

  FINAL_HEADERS.forEach((header, i) => {
    if (header.startsWith('vendor_')) return;
    if (header === "upc") row[i] = upc || data.upc || "";
    else if (header === "offers" && Array.isArray(data.offers)) row[i] = `Count: ${data.offers.length}`;
    else row[i] = flat[header] || data[header] || "";
  });

  const offers = Array.isArray(data.offers)? data.offers : [];
  offers.slice(0, CONFIG_BUSCA_UPC.MAX_VENDORS).forEach((offer, j) => {
    const base = BASE_HEADERS.length + j * 7;
    row[base + 0] = offer.merchant || "";
    row[base + 1] = offer.price || "";
    row[base + 2] = offer.currency || "";
    row[base + 3] = offer.condition || "";
    row[base + 4] = offer.domain || "";
    row[base + 5] = offer.availability || "";
    row[base + 6] = offer.link || "";
  });

  return row;
}

/**
 * ═══════════════════════════════════════════
 * BACKOFF Y SCORING
 * ═══════════════════════════════════════════
 */
function computeBackoffMs(attempt, response = null) {
  if (response) {
    try {
      const retryAfter = Number(response.getHeaders()['Retry-After'] || 0) * 1000;
      if (retryAfter > 0) {
        log(`⏳ API solicitó esperar ${retryAfter/1000}s`, "WARN");
        return Math.min(retryAfter, 60000);
      }
    } catch (e) {}
  }

  const baseDelay = CONFIG_BUSCA_UPC.INITIAL_DELAY * Math.pow(2, attempt);
  const jitter = baseDelay * (Math.random() * 0.3);
  const finalDelay = Math.min(baseDelay + jitter, CONFIG_BUSCA_UPC.MAX_BACKOFF_MS);

  log(`⏳ Esperando ${(finalDelay/1000).toFixed(1)}s (intento ${attempt + 1})`, "INFO");
  return finalDelay;
}

function scoreCandidate(query, candidate) {
  const queryLower = String(query).toLowerCase();
  const qTokens = new Set(queryLower.split(/\s+/).filter(t => t.length > 2));

  const cTitle = (candidate.title || "").toLowerCase();
  const cBrand = (candidate.brand || "").toLowerCase();
  const cModel = (candidate.model || "").toLowerCase();
  const cTokens = new Set([...cTitle.split(/\s+/), cBrand, cModel].filter(t => t && t.length > 2));

  let intersection = 0;
  qTokens.forEach(token => {
    if (cTokens.has(token)) intersection++;
    cTokens.forEach(ct => {
      if (ct.includes(token) || token.includes(ct)) intersection += 0.5;
    });
  });
  const union = new Set([...qTokens,...cTokens]).size;
  const jaccard = union === 0? 0 : intersection / union;

  let brandBonus = 0;
  if (cBrand) {
    if (queryLower.includes(cBrand)) brandBonus = 0.4;
    else if (cBrand.split(' ').some(b => queryLower.includes(b))) brandBonus = 0.2;
  }

  const modelBonus = (cModel && queryLower.includes(cModel))? 0.3 : 0;
  const upcBonus = (candidate.upc || candidate.ean)? 0.2 : 0;

  return jaccard + brandBonus + modelBonus + upcBonus;
}

/**
 * ═══════════════════════════════════════════
 * API HELPERS
 * ═══════════════════════════════════════════
 */
function makeAPIRequest(endpoint, paramsString) {
  const apiKey = getAPIKey();
  const isFreePlan =!apiKey;

  const API_BASE = isFreePlan? 'https://api.upcitemdb.com/prod/trial' : 'https://api.upcitemdb.com/prod/v1';
  const fullUrl = API_BASE + endpoint + '?' + paramsString;

  metrics.logRequest();

  const options = {
    method: 'get',
    headers: isFreePlan? {} : { 'user_key': apiKey },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const code = response.getResponseCode();
    return { response, code };
  } catch (error) {
    log(`Error en API request: ${error.message}`, "ERROR");
    return { response: null, code: 0 };
  }
}

function handleAPIResponse(code, response, attempt, retryCallback) {
  if (code === 200) {
    resetRateLimitCounter();
    return JSON.parse(response.getContentText());
  }

  if (code === 429) {
    metrics.logRateLimit();
    if (trackRateLimit()) {
      delay(60000);
    }
    if (attempt < CONFIG_BUSCA_UPC.MAX_RETRIES) {
      const wait = computeBackoffMs(attempt, response);
      delay(wait);
      return retryCallback();
    }
  }

  if (code >= 500 && attempt < CONFIG_BUSCA_UPC.MAX_RETRIES) {
    const wait = computeBackoffMs(attempt, response);
    log(`⚠️ HTTP ${code} (error servidor). Reintentando...`, "WARN");
    delay(wait);
    return retryCallback();
  }

  if (code === 401) {
    log(`❌ API Auth Error (401): Verifica tu API key de UPCItemDB. Si la key parece de OpenAI (sk-... o empieza con F3A), es incorrecta. Obtén una real de upcitemdb.com.`, "ERROR");
  } else {
    log(`❌ API Error: HTTP ${code}`, "ERROR");
  }
  metrics.logError();
  return null;
}

/**
 * ═══════════════════════════════════════════
 * BÚSQUEDA POR UPC
 * ═══════════════════════════════════════════
 */
function searchProductInfoByUPC(upc) {
  const cleanUpc = String(upc || "").replace(/\D+/g, "");

  if (!isValidGTIN(cleanUpc) &&!isValidUPCFormat(cleanUpc)) {
    metrics.logError();
    return formatEmptyRow(upc, "Formato UPC/GTIN inválido");
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = makeCacheKey('upc', cleanUpc);
  const cached = cache.get(cacheKey);

  if (cached) {
    metrics.logCacheHit();
    log(`📦 Cache hit UPC: ${cleanUpc}`);
    return formatProductData(JSON.parse(cached), cleanUpc);
  }

  metrics.logCacheMiss();

  for (let attempt = 0; attempt <= CONFIG_BUSCA_UPC.MAX_RETRIES; attempt++) {
    try {
      log(`🔍 Consulta UPC: ${cleanUpc} (Intento ${attempt + 1})`);

      const { response, code } = makeAPIRequest('/lookup', `upc=${encodeURIComponent(cleanUpc)}`);

      if (code === 200) {
        const json = JSON.parse(response.getContentText());
        if (json.items?.length > 0) {
          const item = shrinkItemForCache(json.items[0]);
          cache.put(cacheKey, JSON.stringify(item), CONFIG_BUSCA_UPC.CACHE_DURATION);
          return formatProductData(item, cleanUpc);
        }
        return formatEmptyRow(upc, "Sin resultados");
      }

      const result = handleAPIResponse(code, response, attempt, () => searchProductInfoByUPC(upc));
      if (result!== null) return result;
      if (code!== 429 && code < 500) {
        return formatEmptyRow(upc, `Error API: HTTP ${code}`);
      }

    } catch (e) {
      if (attempt < CONFIG_BUSCA_UPC.MAX_RETRIES) {
        const wait = computeBackoffMs(attempt);
        log(`⚠️ Error conexión: ${e.message}. Reintentando...`, "WARN");
        delay(wait);
        continue;
      }
      metrics.logError();
      return formatEmptyRow(upc, `Error API: ${e.message}`);
    }
  }

  metrics.logError();
  return formatEmptyRow(upc, `Fallo tras ${CONFIG_BUSCA_UPC.MAX_RETRIES + 1} intentos`);
}

/**
 * ═══════════════════════════════════════════
 * BÚSQUEDA POR NOMBRE
 * ═══════════════════════════════════════════
 */
function callSearchAPI(searchQuery) {
  for (let attempt = 0; attempt <= CONFIG_BUSCA_UPC.MAX_RETRIES; attempt++) {
    try {
      log(`🔎 API Search: "${searchQuery.substring(0, 50)}..." (Intento ${attempt + 1})`);

      const { response, code } = makeAPIRequest('/search', `s=${encodeURIComponent(searchQuery)}&match_mode=0&type=product`);
      const result = handleAPIResponse(code, response, attempt, () => callSearchAPI(searchQuery));

      if (result) {
        const items = Array.isArray(result.items)? result.items : [];

        if (items.length === 0) {
          log(`📭 Sin resultados para: "${searchQuery.substring(0, 40)}..."`);
          return null;
        }

        let bestCandidate = items[0];
        let maxScore = -1;

        items.slice(0, 15).forEach(item => {
          const score = scoreCandidate(searchQuery, item);
          if (score > maxScore) {
            maxScore = score;
            bestCandidate = item;
          }
        });

        log(`🏅 Mejor candidato: "${(bestCandidate.title || "").substring(0, 50)}..." (Score: ${maxScore.toFixed(2)})`);
        return bestCandidate;
      }

    } catch (e) {
      if (attempt < CONFIG_BUSCA_UPC.MAX_RETRIES) {
        const wait = computeBackoffMs(attempt);
        log(`⚠️ Error conexión: ${e.message}. Reintentando...`, "WARN");
        delay(wait);
        continue;
      }
      metrics.logError();
      log(`❌ Error fatal: ${e.message}`, "ERROR");
      return null;
    }
  }

  return null;
}

function searchProductByNameRaw(productName, useOptimization = true) {
  const rawName = String(productName).trim();

  if (rawName.length < CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH ||!isSearchableProductName(rawName)) {
    return null;
  }

  const searchName = useOptimization? optimizeProductName(rawName) : rawName;

  if (searchName.length < CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH) {
    log(`⚠️ Nombre muy corto después de optimizar: "${searchName}"`, "WARN");
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = makeCacheKey('name', searchName);
  const cached = cache.get(cacheKey);

  if (cached) {
    metrics.logCacheHit();
    log(`📦 Cache hit nombre: "${searchName.substring(0, 40)}..."`);
    return JSON.parse(cached);
  }

  metrics.logCacheMiss();
  const result = callSearchAPI(searchName);

  if (result) {
    const itemToCache = shrinkItemForCache(result);
    cache.put(cacheKey, JSON.stringify(itemToCache), CONFIG_BUSCA_UPC.CACHE_DURATION);
    return itemToCache;
  }

  return null;
}

function smartSearchByName(productName) {
  const rawName = String(productName).trim();

  if (!isSearchableProductName(rawName)) {
    log(`⏭️ Nombre no válido: "${rawName.substring(0, 40)}..."`);
    return null;
  }

  const optimizedName = optimizeProductName(rawName);
  log(`🔍 Búsqueda optimizada: "${optimizedName}"`);

  let result = searchProductByNameRaw(optimizedName, false);

  if (result && (result.upc || result.ean)) {
    log(`✅ Encontrado con nombre optimizado`);
    return result;
  }

  if (CONFIG_BUSCA_UPC.ENABLE_PROGRESSIVE_SEARCH) {
    const shortName = getShortProductName(rawName);

    if (shortName && shortName!== optimizedName && shortName.length >= CONFIG_BUSCA_UPC.MIN_SEARCH_LENGTH) {
      metrics.logProgressiveSearch();
      log(`🔄 Búsqueda progresiva: "${shortName}"`);
      delay(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);

      result = callSearchAPI(shortName);

      if (result && (result.upc || result.ean)) {
        const cache = CacheService.getScriptCache();
        const itemToCache = shrinkItemForCache(result);
        cache.put(makeCacheKey('name', optimizedName), JSON.stringify(itemToCache), CONFIG_BUSCA_UPC.CACHE_DURATION);
        log(`✅ Encontrado con nombre corto`);
        return result;
      }
    }

    const brand = extractBrand(rawName);
    if (brand && brand.length >= 3) {
      const words = optimizedName.split(/\s+/);
      const brandFirstWord = brand.split(' ')[0];
      const brandIdx = words.findIndex(w => w.toLowerCase().includes(brandFirstWord));

      if (brandIdx >= 0 && words[brandIdx + 1]) {
        const brandPlusOne = `${brand} ${words[brandIdx + 1]}`.trim();

        if (brandPlusOne!== shortName && brandPlusOne.length >= 5) {
          metrics.logProgressiveSearch();
          log(`🔄 Búsqueda por marca: "${brandPlusOne}"`);
          delay(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);

          result = callSearchAPI(brandPlusOne);

          if (result && (result.upc || result.ean)) {
            const cache = CacheService.getScriptCache();
            const itemToCache = shrinkItemForCache(result);
            cache.put(makeCacheKey('name', optimizedName), JSON.stringify(itemToCache), CONFIG_BUSCA_UPC.CACHE_DURATION);
            log(`✅ Encontrado con marca`);
            return result;
          }
        }
      }
    }
  }

  log(`❌ No encontrado: "${rawName.substring(0, 50)}..."`);
  return null;
}

function searchProductByName(productName) {
  const result = smartSearchByName(productName);
  return result ? formatProductData(result, result.upc || result.ean || "") : formatEmptyRow("", "Sin resultados");
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * FUNCIONES PRINCIPALES DE PROCESAMIENTO
 * ═══════════════════════════════════════════
 */
function promptUserInput(ui, title, message) {
  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  return {
    confirmed: response.getSelectedButton() === ui.Button.OK,
    value: response.getResponseText().trim()
  };
}

function processRow(sheet, currentRow, productName, destCol, currentDest) {
  if (!productName) return { skipped: true };

  if (currentDest && currentDest !== "NO_ENCONTRADO" && currentDest !== "ERROR") {
    log(`⏭️ Fila ${currentRow}: Ya tiene UPC: ${currentDest}`);
    return { skipped: true };
  }

  if (!isSearchableProductName(productName)) {
    log(`⏭️ Fila ${currentRow}: Nombre no válido: "${productName.substring(0, 40)}..."`);
    return { skipped: true, invalid: true };
  }

  log(`\n📍 Fila ${currentRow}: "${productName.substring(0, 60)}..."`);

  try {
    const item = smartSearchByName(productName);

    if (item && (item.upc || item.ean)) {
      const resolvedUPC = String(item.upc || item.ean).trim();
      sheet.getRange(currentRow, destCol).setValue(resolvedUPC);
      log(`✅ Fila ${currentRow}: UPC → ${resolvedUPC} (${(item.title || "").substring(0, 40)}...)`);
      return { success: true };
    } else {
      sheet.getRange(currentRow, destCol).setValue("NO_ENCONTRADO");
      log(`❌ Fila ${currentRow}: Sin resultados → NO_ENCONTRADO`);
      return { notFound: true };
    }
  } catch (err) {
    sheet.getRange(currentRow, destCol).setValue("ERROR");
    log(`🔴 Fila ${currentRow}: ${err.message}`, "ERROR");
    return { error: true };
  }
}

function buscarUPCsPorNombre() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  getAPIKey();

  const sheets = ss.getSheets().map(s => s.getName());
  const sheetList = "Selecciona el número de la hoja:\n\n" +
    sheets.map((name, i) => `${i + 1}. ${name}`).join('\n');

  const sheetInput = promptUserInput(ui, "📑 Selección de hoja", sheetList);
  if (!sheetInput.confirmed) return;

  const sheetIndex = Number(sheetInput.value);
  if (!sheetIndex || sheetIndex < 1 || sheetIndex > sheets.length) {
    return ui.alert("❌ Error", "Número de hoja inválido", ui.ButtonSet.OK);
  }
  const sheet = ss.getSheetByName(sheets[sheetIndex - 1]);

  const originInput = promptUserInput(ui, "📝 Columna ORIGEN", "Letra de la columna con NOMBRE del producto (ej: A, B, C):");
  if (!originInput.confirmed) return;
  const originCol = columnLetterToIndex(originInput.value);
  if (!originCol) return ui.alert("❌ Error", "Columna origen inválida", ui.ButtonSet.OK);

  const destInput = promptUserInput(ui, "🎯 Columna DESTINO", "Letra de la columna donde escribir el UPC (ej: B, C, D):");
  if (!destInput.confirmed) return;
  const destCol = columnLetterToIndex(destInput.value);
  if (!destCol) return ui.alert("❌ Error", "Columna destino inválida", ui.ButtonSet.OK);

  const startInput = promptUserInput(ui, "📍 Fila inicial", "¿Desde qué fila quieres empezar? (ej: 2 para saltar encabezado):");
  if (!startInput.confirmed) return;
  const startRow = Number(startInput.value);
  if (!startRow || startRow < 1) return ui.alert("❌ Error", "Fila inicial inválida", ui.ButtonSet.OK);

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return ui.alert("⚠️ Aviso", "No hay datos para procesar", ui.ButtonSet.OK);

  const numRows = lastRow - startRow + 1;
  const confirmResp = ui.alert(
    "🚀 Confirmar proceso",
    `Hoja: ${sheets[sheetIndex - 1]}\nColumna origen: ${originInput.value.toUpperCase()}\nColumna destino: ${destInput.value.toUpperCase()}\nFilas: ${startRow} a ${lastRow} (${numRows} filas)\n\n¿Continuar?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResp !== ui.Button.YES) return;

  log(`🚀 Iniciando búsqueda: ${sheets[sheetIndex - 1]}, filas ${startRow}-${lastRow}`);
  SpreadsheetApp.getActiveSpreadsheet().toast("Proceso iniciado. Revisa el log para progreso.", "🔍 UPC Lookup", 10);

  const maxCol = Math.max(originCol, destCol);
  const data = sheet.getRange(startRow, 1, numRows, maxCol).getValues();

  let procesados = 0, encontrados = 0, noEncontrados = 0, saltados = 0, errores = 0;
  const startTime = new Date();

  for (let i = 0; i < data.length; i++) {
    const currentRow = startRow + i;
    const productName = String(data[i][originCol - 1] || "").trim();
    const currentDest = String(data[i][destCol - 1] || "").trim();

    const result = processRow(sheet, currentRow, productName, destCol, currentDest);

    if (result.skipped) {
      if (result.invalid) saltados++;
      continue;
    }

    procesados++;
    if (result.success) encontrados++;
    else if (result.notFound) noEncontrados++;
    else if (result.error) errores++;

    if (i < data.length - 1) delay(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS);

    if (procesados % 5 === 0) {
      SpreadsheetApp.flush();
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Progreso: ${procesados} procesados, ${encontrados} encontrados...`,
        "🔍 UPC Lookup",
        5
      );
    }
  }

  const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  const successRate = procesados > 0 ? ((encontrados / procesados) * 100).toFixed(1) : 0;

  const summary = [
    `📊 RESUMEN`,
    ``,
    `⏱️ Tiempo: ${elapsed} segundos`,
    ``,
    `• Total filas: ${numRows}`,
    `• Procesados: ${procesados}`,
    `• ✅ UPCs encontrados: ${encontrados} (${successRate}%)`,
    `• ❌ No encontrados: ${noEncontrados}`,
    `• ⏭️ Saltados: ${saltados}`,
    `• 🔴 Errores: ${errores}`,
    ``,
    `Búsquedas progresivas: ${metrics.getSummary().progressiveSearches}`
  ].join('\n');

  log(`\n${summary}`);
  ui.alert("✅ Proceso terminado", summary, ui.ButtonSet.OK);
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * FUNCIÓN onEdit ESPECÍFICA PARA UPC - Renombrada para evitar conflictos
 * ═══════════════════════════════════════════
 */
function updateProductInfoFromUPCs_onEdit_(e) {
  try {
    const hoja = e.range.getSheet();
    const fila = e.range.getRow();
    const col = e.range.getColumn();

    // Solo procesar si es columna UPC y fila >= START_ROW
    if (fila < CONFIG_BUSCA_UPC.START_ROW) return;
    if (col !== CONFIG_BUSCA_UPC.UPC_COLUMN) return;

    const upc = String(e.value || '').trim();
    if (!upc || !isValidGTIN(upc)) return;

    log(`onEdit UPC: Detectado UPC en fila ${fila}: ${upc}`);
    
    const data = searchProductInfoByUPC(upc);
    if (data && data[IDX.Error] === '') {
      // Escribir datos en la misma fila, columnas siguientes
      hoja.getRange(fila, CONFIG_BUSCA_UPC.UPC_COLUMN + 1, 1, data.length - 1)
        .setValues([data.slice(1)]);
      log(`onEdit UPC: Fila ${fila} actualizada`);
    }

  } catch (err) {
    log(`Error en updateProductInfoFromUPCs_onEdit_: ${err.message}`);
  }
}

/**
 * ═══════════════════════════════════════════
 * FUNCIONES AUXILIARES FALTANTES
 * ═══════════════════════════════════════════════════════════════════════════
 */
function updateProductInfoFromUPCs() {
  const sheet = getSheetByName(CONFIG_BUSCA_UPC.SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_BUSCA_UPC.START_ROW) {
    log("No hay datos para procesar", "WARN");
    return;
  }

  const numRows = lastRow - CONFIG_BUSCA_UPC.START_ROW + 1;
  const upcData = sheet.getRange(CONFIG_BUSCA_UPC.START_ROW, CONFIG_BUSCA_UPC.UPC_COLUMN, numRows, 1).getValues();

  let procesados = 0;
  for (let i = 0; i < upcData.length; i++) {
    const upc = String(upcData[i][0] || '').trim();
    if (!upc) continue;

    const fila = CONFIG_BUSCA_UPC.START_ROW + i;
    const data = searchProductInfoByUPC(upc);
    
    if (data) {
      sheet.getRange(fila, CONFIG_BUSCA_UPC.UPC_COLUMN, 1, data.length).setValues([data]);
      procesados++;
    }

    if (procesados % 5 === 0) delay(CONFIG_BUSCA_UPC.BATCH_PAUSE_MS);
  }

  log(`Actualización completada: ${procesados} UPCs procesados`);
}

function buscarYCompletarInteractivoAvanzado() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Función avanzada", "Esta función requiere configuración adicional. Usa 'Buscar UPCs por NOMBRE' por ahora.", ui.ButtonSet.OK);
}

function setupAPIKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Configurar API Key de UPCItemDB',
    'Pega tu API key de upcitemdb.com:\n(Deja vacío para usar plan gratuito)',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty(CONFIG_BUSCA_UPC.PROPERTY_KEY, key);
    ui.alert('API Key guardada', key ? 'Key configurada correctamente' : 'Usando plan gratuito', ui.ButtonSet.OK);
  }
}

function diagnosticoSistema() {
  const ui = SpreadsheetApp.getUi();
  const apiKey = getAPIKey();
  const msg = [
    `Diagnóstico del Sistema:`,
    ``,
    `API Key: ${apiKey ? 'Configurada' : 'No configurada (plan gratuito)'}`,
    `Hoja configurada: ${CONFIG_BUSCA_UPC.SHEET_NAME}`,
    `Fila inicio: ${CONFIG_BUSCA_UPC.START_ROW}`,
    `Columna UPC: ${CONFIG_BUSCA_UPC.UPC_COLUMN}`,
    ``,
    `Cache namespace: ${getCacheNs()}`
  ].join('\n');
  ui.alert('Diagnóstico', msg, ui.ButtonSet.OK);
}

function mostrarMetricas() {
  const ui = SpreadsheetApp.getUi();
  const s = metrics.getSummary();
  const msg = [
    `Métricas de Sesión:`,
    ``,
    `Tiempo: ${s.elapsedSeconds}s`,
    `Requests: ${s.requests}`,
    `Cache hits: ${s.cacheHits} (${s.cacheHitRate})`,
    `Cache misses: ${s.cacheMisses}`,
    `Búsquedas progresivas: ${s.progressiveSearches}`,
    `Rate limits: ${s.rateLimits}`,
    `Errores: ${s.errors}`
  ].join('\n');
  ui.alert('Métricas', msg, ui.ButtonSet.OK);
}

function testBusquedaNombre() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Test Búsqueda', 'Ingresa un nombre de producto para probar:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const nombre = result.getResponseText().trim();
    const item = smartSearchByName(nombre);
    const msg = item 
      ? `Encontrado:\nUPC: ${item.upc || item.ean}\nTítulo: ${item.title}\nMarca: ${item.brand}`
      : 'No encontrado';
    ui.alert('Resultado', msg, ui.ButtonSet.OK);
  }
}



/**
 * Obtiene los UPC de todos los productos en la hoja actual 
 * y los escribe en la columna O (empezando en O6).
 */
function obtenerUPCsParaHojaActual() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Configuración de columnas según la imagen
  const FILA_INICIO = 7; // Los datos empiezan en la fila 7
  const COL_PRODUCTO = 3; // Columna C
  const COL_DESTINO = 15; // Columna O
  const FILA_HEADER = 6; // Para poner el título en O6

  const lastRow = sheet.getLastRow();
  
  if (lastRow < FILA_INICIO) {
    ui.alert("No se detectaron datos en la hoja.");
    return;
  }

  // 1. Ponemos el encabezado en O6
  sheet.getRange(FILA_HEADER, COL_DESTINO).setValue("UPC ENCONTRADO")
       .setBackground("#4a86e8") // Azul opcional para que combine
       .setFontColor("white")
       .setFontWeight("bold");

  // 2. Obtenemos todos los nombres de productos de la columna C
  const numRows = lastRow - FILA_INICIO + 1;
  const nombres = sheet.getRange(FILA_INICIO, COL_PRODUCTO, numRows, 1).getValues();
  const resultadosUPC = [];

  log(`Iniciando búsqueda masiva para ${numRows} productos...`);
  SpreadsheetApp.getActive().toast("Buscando UPCs... por favor espera.", "🔍 UPC Service");

  // 3. Iteramos y buscamos
  for (let i = 0; i < nombres.length; i++) {
    const nombreProducto = nombres[i][0];
    
    if (nombreProducto && String(nombreProducto).trim() !== "") {
      try {
        // Llamamos a la función de búsqueda del sistema
        const info = smartSearchByName(nombreProducto);
        const upc = info ? (info.upc || info.ean || "NO ENCONTRADO") : "NO ENCONTRADO";
        
        resultadosUPC.push([upc]);
        log(`Fila ${FILA_INICIO + i}: ${upc}`);
      } catch (e) {
        resultadosUPC.push(["ERROR API"]);
      }
    } else {
      resultadosUPC.push([""]); // Celda vacía si no hay nombre
    }

    // Pausa de seguridad para evitar bloqueos de la API (0.5 segundos)
    Utilities.sleep(CONFIG_BUSCA_UPC.SEARCH_PAUSE_MS || 500);
    
    // Actualizar progreso en la barra inferior cada 5 productos
    if (i % 5 === 0) {
      SpreadsheetApp.getActive().toast(`Procesando: ${i + 1} de ${numRows}`, "🔍 UPC Service");
    }
  }

  // 4. Escribimos todos los resultados en la columna O de un solo golpe (Batch Write)
  sheet.getRange(FILA_INICIO, COL_DESTINO, resultadosUPC.length, 1).setValues(resultadosUPC);

  ui.alert("✅ Proceso completado. Los UPC se han guardado en la columna O.");
}