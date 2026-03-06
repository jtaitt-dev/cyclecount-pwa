const $ = id => document.getElementById(id)

// === Toast Notifications ===
const toast = (msg, type = 'info') => {
  const el = document.createElement('div')
  el.className = `toast toast-${type}`
  el.textContent = msg
  $('toasts').appendChild(el)
  setTimeout(() => {
    el.classList.add('fade-out')
    setTimeout(() => el.remove(), 300)
  }, 3500)
}

// === Error Handling ===
window.addEventListener('error', e => {
  console.error(e)
  toast('Error: ' + (e?.message || e), 'error')
})
window.addEventListener('unhandledrejection', e => {
  console.error(e)
  toast('Error: ' + (e?.reason?.message || e?.reason || e), 'error')
})

// === UI Helpers ===
const setBadge = (id, txt, cls) => {
  const el = $(id)
  if (!el) return
  el.textContent = txt
  el.className = (id.startsWith('chip') ? 'chip-val ' : 'badge-val ') + (cls || '')
}

const setStep = (num, done, statusText) => {
  const numEl = $('step' + num + 'Num')
  const statusEl = $('step' + num + 'Status')
  if (numEl && done) {
    numEl.textContent = '\u2713'
    numEl.className = 'step-num done'
  }
  if (statusEl) {
    statusEl.textContent = statusText || ''
    statusEl.className = 'step-status' + (done ? ' good' : '')
  }
}

const withLoading = (btn, fn, loadingText) => async () => {
  const origHTML = btn.innerHTML
  btn.disabled = true
  btn.classList.add('loading')
  if (loadingText) btn.textContent = loadingText
  try {
    await fn()
  } catch (e) {
    toast(e?.message || String(e), 'error')
  } finally {
    btn.disabled = false
    btn.classList.remove('loading')
    btn.innerHTML = origHTML
  }
}

const showConfirm = (title, msg, okText = 'OK', okStyle = '') => new Promise(resolve => {
  $('confirmTitle').textContent = title
  $('confirmMsg').textContent = msg
  const okBtn = $('confirmOk')
  okBtn.textContent = okText
  okBtn.style.background = okStyle || 'var(--primary)'
  $('confirmModal').classList.remove('hidden')
  const cleanup = result => {
    $('confirmModal').classList.add('hidden')
    resolve(result)
  }
  okBtn.onclick = () => cleanup(true)
  $('confirmCancel').onclick = () => cleanup(false)
  $('confirmModal').onclick = e => {
    if (e.target === $('confirmModal')) cleanup(false)
  }
})

const timeAgo = iso => {
  const s = Math.floor((Date.now() - new Date(iso).getTime()) / 1000)
  if (s < 60) return 'just now'
  if (s < 3600) return Math.floor(s / 60) + 'm ago'
  if (s < 86400) return Math.floor(s / 3600) + 'h ago'
  return Math.floor(s / 86400) + 'd ago'
}

const nowIso = () => new Date().toISOString()

const formatDateLocal = iso => {
  const d = new Date(iso)
  if (isNaN(d.getTime())) return ''
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  const dd = String(d.getDate()).padStart(2, '0')
  const yyyy = d.getFullYear()
  let h = d.getHours()
  const ampm = h >= 12 ? 'PM' : 'AM'
  h = h % 12 || 12
  const min = String(d.getMinutes()).padStart(2, '0')
  return mm + '-' + dd + '-' + yyyy + ' ' + h + ':' + min + ' ' + ampm
}

// === Theme ===
const applyTheme = (theme) => {
  document.documentElement.setAttribute('data-theme', theme)
  localStorage.setItem('theme', theme)
  const meta = document.querySelector('meta[name="theme-color"]')
  if (meta) meta.content = theme === 'light' ? '#f0f2f5' : '#0a0e1a'
  const btn = $('themeToggle')
  if (btn) btn.textContent = theme === 'light' ? '\uD83C\uDF19' : '\u2600\uFE0F'
}

const toggleTheme = () => {
  const current = localStorage.getItem('theme') || 'dark'
  applyTheme(current === 'dark' ? 'light' : 'dark')
}

// === Config ===
const CFG = {
  clientId: '5591a292-6fe4-4ddd-9754-4b55efbce9be',
  authority: 'https://login.microsoftonline.com/common',
  scopes: ['User.Read', 'Files.ReadWrite.All', 'offline_access']
}

// === State ===
let msalApp = null
let account = null
let accessToken = null
let driveId = null
let itemId = null
let inventoryCache = new Map()
let scanner = null
let scanning = false
let setupCollapsed = false
let editingIndex = -1
let photoLookupTimer = null

const DBK = 'cycleCounts.v1'
const db = {
  get: () => JSON.parse(localStorage.getItem(DBK) || '[]'),
  set: x => localStorage.setItem(DBK, JSON.stringify(x))
}

// === Helpers ===
const base64Url = s => {
  const b = btoa(unescape(encodeURIComponent(s)))
  return b.replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '')
}
const shareIdFromLink = url => 'u!' + base64Url(url)

const getBaseUrl = () => {
  let p = window.location.pathname
  if (p.endsWith('index.html')) p = p.substring(0, p.lastIndexOf('/') + 1)
  if (!p.endsWith('/')) p += '/'
  return window.location.origin + p
}

// === MSAL ===
const requireMsal = () => {
  if (typeof msal === 'undefined') {
    toast('MSAL library not loaded. Check your network.', 'error')
    return false
  }
  return true
}

const msalInit = async () => {
  if (!requireMsal()) return
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: CFG.authority,
      redirectUri: getBaseUrl()
    },
    cache: {
      cacheLocation: 'localStorage',
      storeAuthStateInCookie: false
    }
  })
  await msalApp.initialize()
}

const pickAccount = () => {
  const accts = msalApp.getAllAccounts()
  if (accts?.length) account = accts[0]
}

const handleRedirect = async () => {
  const r = await msalApp.handleRedirectPromise()
  if (r?.account) account = r.account
  if (!account) pickAccount()
  if (account) {
    setBadge('chipAuth', 'Signed in', 'good')
    setStep(1, true, account.username || 'Signed in')
  } else {
    setBadge('chipAuth', 'Signed out', 'bad')
  }
}

const ensureToken = async () => {
  if (!account) pickAccount()
  if (!account) throw new Error('Not signed in. Complete Step 1 first.')
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: CFG.scopes, account })
    accessToken = r.accessToken
    return accessToken
  } catch (e) {
    toast('Redirecting for login...', 'warning')
    await msalApp.acquireTokenRedirect({ scopes: CFG.scopes })
    throw new Error('Redirecting for token')
  }
}

// === Graph API ===
const graph = async (method, url, body) => {
  if (!accessToken) await ensureToken()
  const h = { 'Authorization': 'Bearer ' + accessToken }
  if (body) h['Content-Type'] = 'application/json'
  const r = await fetch('https://graph.microsoft.com/v1.0' + url, {
    method, headers: h,
    body: body ? JSON.stringify(body) : undefined
  })
  const t = await r.text()
  if (!r.ok) throw new Error('Graph ' + r.status + ': ' + t.substring(0, 200))
  return t ? JSON.parse(t) : {}
}

const graphRaw = async (method, url, headers, body) => {
  if (!accessToken) await ensureToken()
  const h = { 'Authorization': 'Bearer ' + accessToken, ...(headers || {}) }
  const r = await fetch('https://graph.microsoft.com/v1.0' + url, { method, headers: h, body })
  const t = await r.text()
  if (!r.ok) throw new Error('Graph ' + r.status + ': ' + t.substring(0, 200))
  return t ? JSON.parse(t) : {}
}

// Safe graph call — returns null on error instead of throwing
const graphSafe = async (method, url, body) => {
  try { return await graph(method, url, body) } catch (e) { console.warn('graphSafe:', e.message); return null }
}

// === Workbook ===
const wbUrl = () => '/drives/' + driveId + '/items/' + itemId + '/workbook'

const resolveWorkbook = async () => {
  const link = $('shareLink').value.trim()
  if (!link) throw new Error('Paste a OneDrive share link first (Step 2)')
  const sid = shareIdFromLink(link)
  const di = await graph('GET', '/shares/' + sid + '/driveItem')
  itemId = di.id
  driveId = di.parentReference?.driveId
  if (!driveId || !itemId) throw new Error('Unable to resolve drive/item id')
  setStep(2, true, 'Workbook connected')
  toast('Workbook connected', 'success')
}

// =========================================================
// FEATURE 1: Ensure OnHand table has extended columns
// =========================================================
const ensureOnHandColumns = async () => {
  const hdrRange = await graphSafe('GET', wbUrl() + '/tables/OnHand/headerRowRange')
  if (!hdrRange) return
  const headers = (hdrRange.values?.[0] || []).map(h => (h ?? '').toString().trim())
  const needed = ['SalesPadQty', 'LastCountedQty', 'LastCountDate', 'LastCountedBy']
  const missing = needed.filter(n => !headers.includes(n))
  if (missing.length === 0) return

  for (const colName of missing) {
    try {
      await graph('POST', wbUrl() + '/tables/OnHand/columns', { name: colName })
    } catch (e) {
      console.warn('Could not add column ' + colName + ':', e.message)
    }
  }
  toast('OnHand table updated with new columns', 'info')
}

// === Inventory Cache (reads extended columns) ===
const loadInventoryCache = async () => {
  inventoryCache.clear()
  await ensureOnHandColumns()

  const rows = await graph('GET', wbUrl() + '/tables/OnHand/rows?$top=100000')
  const vals = rows.value?.map(x => x.values?.[0]).filter(Boolean) || []
  let ok = 0
  for (const v of vals) {
    const sku = (v[0] ?? '').toString().trim()
    const bin = (v[1] ?? '').toString().trim()
    const exp = Number(v[2] ?? 0)
    const salespadQty = (v[3] !== undefined && v[3] !== null && v[3] !== '') ? Number(v[3]) : null
    const lastCountedQty = (v[4] !== undefined && v[4] !== null && v[4] !== '') ? Number(v[4]) : null
    const lastCountDate = (v[5] ?? '').toString().trim() || null
    const lastCountedBy = (v[6] ?? '').toString().trim() || null
    if (!sku) continue
    inventoryCache.set(sku, { bin, expectedQty: exp, salespadQty, lastCountedQty, lastCountDate, lastCountedBy })
    ok++
  }
  setBadge('chipCache', ok + ' SKUs', 'good')
  setStep(3, true, ok + ' SKUs loaded')
  toast('Loaded ' + ok + ' SKUs', 'success')
  validateAndCompute()
  if (account && driveId && inventoryCache.size > 0 && !setupCollapsed) {
    toggleSetup()
  }
}

// =========================================================
// FEATURE 4: Update OnHand row on sync
// =========================================================
const updateOnHandRow = async (sku, countedQty, timestamp) => {
  const hdrRange = await graphSafe('GET', wbUrl() + '/tables/OnHand/headerRowRange')
  if (!hdrRange) return
  const headers = (hdrRange.values?.[0] || []).map(h => (h ?? '').toString().trim())
  const colLastCounted = headers.indexOf('LastCountedQty')
  const colLastDate = headers.indexOf('LastCountDate')
  const colLastBy = headers.indexOf('LastCountedBy')
  if (colLastCounted < 0 && colLastDate < 0 && colLastBy < 0) return

  const rows = await graph('GET', wbUrl() + '/tables/OnHand/rows?$top=100000')
  const allRows = rows.value || []
  let matchIdx = -1
  for (let i = 0; i < allRows.length; i++) {
    const rowSku = (allRows[i].values?.[0]?.[0] ?? '').toString().trim()
    if (rowSku === sku) { matchIdx = i; break }
  }
  if (matchIdx < 0) return

  const existing = allRows[matchIdx].values[0].slice()
  const maxCol = Math.max(colLastCounted, colLastDate, colLastBy)
  while (existing.length <= maxCol) existing.push('')
  if (colLastCounted >= 0) existing[colLastCounted] = countedQty
  if (colLastDate >= 0) existing[colLastDate] = formatDateLocal(timestamp)
  if (colLastBy >= 0) existing[colLastBy] = account?.name || account?.username || 'Unknown'

  await graph('PATCH', wbUrl() + '/tables/OnHand/rows/itemAt(index=' + matchIdx + ')', {
    values: [existing]
  })
}

// =========================================================
// FEATURE: Add new SKU to OnHand table (inline)
// =========================================================
const addSkuToInventory = async () => {
  const sku = $('sku').value.trim()
  if (!sku) { toast('Enter a SKU first', 'warning'); return }
  if (inventoryCache.has(sku)) { toast('SKU already exists in inventory', 'info'); return }
  if (!driveId || !itemId) { toast('Connect workbook first', 'warning'); return }

  const bin = $('bin').value.trim()
  const expectedQty = Number($('countedQty').value || 0)

  const ok = await showConfirm(
    'Add SKU to Inventory?',
    'Add "' + sku + '" to OnHand table with Bin: ' + (bin || '—') + ', Expected Qty: ' + expectedQty + '?',
    'Confirm'
  )
  if (!ok) return

  try {
    await graph('POST', wbUrl() + '/tables/OnHand/rows/add', {
      values: [[sku, bin, expectedQty, '', '', '', '']]
    })

    // Update local cache
    inventoryCache.set(sku, {
      bin, expectedQty, salespadQty: null, lastCountedQty: null, lastCountDate: null, lastCountedBy: null
    })
    setBadge('chipCache', inventoryCache.size + ' SKUs', 'good')
    toast('SKU "' + sku + '" added to inventory', 'success')
    validateAndCompute()
  } catch (e) {
    toast('Failed to add SKU: ' + (e?.message || e), 'error')
  }
}

// =========================================================
// FEATURE: Move SKU to different bin (inline)
// =========================================================
const moveSkuBin = async () => {
  const sku = $('sku').value.trim()
  if (!sku) return
  const hit = inventoryCache.get(sku)
  if (!hit) { toast('SKU not found in inventory', 'warning'); return }
  if (!driveId || !itemId) { toast('Connect workbook first', 'warning'); return }

  const newBin = prompt('Move "' + sku + '" to which bin?', hit.bin || '')
  if (newBin === null || newBin.trim() === '') return
  const trimBin = newBin.trim()
  if (trimBin === hit.bin) { toast('Same bin — no change', 'info'); return }

  try {
    // Find the row in OnHand table
    const rows = await graph('GET', wbUrl() + '/tables/OnHand/rows?$top=100000')
    const allRows = rows.value || []
    let matchIdx = -1
    for (let i = 0; i < allRows.length; i++) {
      const rowSku = (allRows[i].values?.[0]?.[0] ?? '').toString().trim()
      if (rowSku === sku) { matchIdx = i; break }
    }
    if (matchIdx < 0) { toast('SKU not found in OnHand table', 'error'); return }

    // Update the bin column (index 1)
    const existing = allRows[matchIdx].values[0].slice()
    existing[1] = trimBin
    await graph('PATCH', wbUrl() + '/tables/OnHand/rows/itemAt(index=' + matchIdx + ')', {
      values: [existing]
    })

    // Update local cache
    hit.bin = trimBin
    inventoryCache.set(sku, hit)

    // Update the bin field in the form
    $('bin').value = trimBin
    toast('Moved "' + sku + '" to bin ' + trimBin, 'success')
    validateAndCompute()
  } catch (e) {
    toast('Failed to move bin: ' + (e?.message || e), 'error')
  }
}

// === Photo Upload ===
const uploadPhotoToOneDrive = async (file, sku) => {
  const folder = 'CycleCountPhotos'
  const ts = nowIso().replaceAll(':', '-')
  const safeSku = (sku || 'SKU').replaceAll('/', '_').replaceAll('\\', '_').replaceAll(' ', '_')
  const name = safeSku + '_' + ts + '.jpg'
  const up = await graphRaw('PUT', '/me/drive/root:/' + folder + '/' + name + ':/content', { 'Content-Type': file.type || 'image/jpeg' }, file)
  const shared = await graph('POST', '/me/drive/items/' + up.id + '/createLink', { type: 'view', scope: 'anonymous' })
  return shared.link?.webUrl || ''
}

// =========================================================
// FEATURE 3: Existing photo lookup on SKU entry
// =========================================================
const lookupExistingPhoto = async (sku) => {
  const wrap = $('existingPhotoWrap')
  const img = $('existingPhoto')
  const label = $('existingPhotoLabel')
  if (!wrap || !img) return

  if (!sku) {
    wrap.classList.add('hidden')
    return
  }

  // Error handler — hide gracefully if image fails to load
  img.onerror = () => {
    label.textContent = 'Previous photo unavailable'
    img.classList.add('hidden')
    wrap.classList.remove('hidden')
  }
  img.onload = () => {
    img.classList.remove('hidden')
  }

  // 1. Check local records first (works offline — base64 data URLs never expire)
  const rows = db.get()
  const localMatch = [...rows].reverse().find(r =>
    r.SKU === sku && r._photoLocal
  )
  if (localMatch) {
    img.src = localMatch._photoLocal
    label.textContent = 'Previous photo \u00b7 ' + timeAgo(localMatch.Timestamp)
    wrap.classList.remove('hidden')
    return
  }

  // 2. Fallback: search shared CycleCounts table for any user's photo of this SKU
  if (!accessToken || !driveId || !itemId) {
    wrap.classList.add('hidden')
    return
  }
  try {
    // CycleCounts columns: [0:Timestamp, 1:SKU, 2:Bin, 3:ExpectedQty, 4:CountedQty, 5:Variance, 6:PhotoUrl, 7:Device, 8:Notes]
    const range = await graphSafe('GET', wbUrl() + '/tables/CycleCounts/dataBodyRange')
    if (!range) { wrap.classList.add('hidden'); return }
    const formulas = range.formulas || []
    const values = range.values || []
    const len = formulas.length || values.length
    // Search from bottom (most recent) for matching SKU with a photo
    for (let i = len - 1; i >= 0; i--) {
      const rowSku = ((values[i] || [])[1] ?? '').toString().trim()
      if (rowSku !== sku) continue
      // Extract photo URL from HYPERLINK formula or plain value
      const formula = ((formulas[i] || [])[6] ?? '').toString().trim()
      const value = ((values[i] || [])[6] ?? '').toString().trim()
      let photoShareUrl = ''
      const m = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/)
      if (m) photoShareUrl = m[1]
      if (!photoShareUrl && value.startsWith('http')) photoShareUrl = value
      if (!photoShareUrl) continue
      // Resolve anonymous share link to a direct download URL
      const sid = shareIdFromLink(photoShareUrl)
      const item = await graphSafe('GET', '/shares/' + sid + '/driveItem')
      const dlUrl = item?.['@microsoft.graph.downloadUrl']
      if (dlUrl) {
        img.src = dlUrl
        const ts = ((values[i] || [])[0] ?? '').toString().trim()
        label.textContent = 'Previous photo' + (ts ? ' \u00b7 ' + timeAgo(ts) : '')
        wrap.classList.remove('hidden')
        return
      }
    }
  } catch (e) {
    console.warn('Photo lookup error:', e.message)
  }
  wrap.classList.add('hidden')
}

const debouncedPhotoLookup = (sku) => {
  if (photoLookupTimer) clearTimeout(photoLookupTimer)
  photoLookupTimer = setTimeout(() => lookupExistingPhoto(sku), 500)
}

// === Excel Write ===
const addRowToExcel = async row => {
  const photoCell = row.PhotoUrl
    ? '=HYPERLINK("' + row.PhotoUrl.replaceAll('"', '""') + '","\uD83D\uDCF7 View Photo")'
    : ''
  // Use formulas array when we have a HYPERLINK formula, values array otherwise
  if (row.PhotoUrl) {
    const formulas = [[row.Timestamp, row.SKU, row.Bin, row.ExpectedQty, row.CountedQty, row.Variance, photoCell, row.Device, row.Notes]]
    await graph('POST', wbUrl() + '/tables/CycleCounts/rows/add', { values: formulas })
  } else {
    await graph('POST', wbUrl() + '/tables/CycleCounts/rows/add', {
      values: [[row.Timestamp, row.SKU, row.Bin, row.ExpectedQty, row.CountedQty, row.Variance, '', row.Device, row.Notes]]
    })
  }
}

// === Validate & Compute ===
const validateAndCompute = () => {
  const sku = $('sku').value.trim()
  const bin = $('bin').value.trim()
  const counted = Number($('countedQty').value || 0)
  const hit = inventoryCache.get(sku)

  const lastCountEl = $('lastCountInfo')
  const inlineActionsEl = $('inlineActions')
  const addSkuBtn = $('btnAddSku')
  const moveBinBtn = $('btnMoveBin')

  if (hit) {
    $('expectedQty').value = hit.expectedQty
    $('salespadQty').value = hit.salespadQty !== null ? hit.salespadQty : '\u2014'

    // Auto-fill bin from cache when field is empty
    if (!bin && hit.bin) {
      $('bin').value = hit.bin
    }

    const currentBin = $('bin').value.trim()
    setBadge('lookupState', 'FOUND', 'good')
    if (currentBin && hit.bin && currentBin.toLowerCase() === hit.bin.toLowerCase()) {
      setBadge('binState', 'OK', 'good')
    } else if (currentBin && hit.bin) {
      setBadge('binState', 'MISMATCH (' + hit.bin + ')', 'bad')
    } else {
      setBadge('binState', '\u2014', '')
    }
    const variance = counted - Number(hit.expectedQty || 0)
    setBadge('varState', String(variance), variance === 0 ? 'good' : (Math.abs(variance) <= 2 ? 'warn' : 'bad'))

    // Show last count info
    if (lastCountEl) {
      if (hit.lastCountedQty !== null && hit.lastCountDate) {
        let info = 'Last counted: <strong>' + hit.lastCountedQty + ' units</strong> \u00b7 ' + timeAgo(hit.lastCountDate)
        if (hit.lastCountedBy) info += ' by ' + hit.lastCountedBy
        lastCountEl.innerHTML = info
        lastCountEl.classList.remove('hidden')
      } else {
        lastCountEl.classList.add('hidden')
      }
    }

    // Show Move Bin button, hide Add SKU button
    if (inlineActionsEl) {
      inlineActionsEl.classList.remove('hidden')
      if (addSkuBtn) addSkuBtn.classList.add('hidden')
      if (moveBinBtn) moveBinBtn.classList.remove('hidden')
    }
  } else {
    $('expectedQty').value = ''
    $('salespadQty').value = ''
    setBadge('lookupState', sku ? 'NOT FOUND' : '\u2014', sku ? 'bad' : '')
    setBadge('binState', '\u2014', '')
    setBadge('varState', '\u2014', '')

    // Hide last count info
    if (lastCountEl) lastCountEl.classList.add('hidden')

    // Show Add SKU button when SKU typed but not found, hide Move Bin
    if (inlineActionsEl) {
      if (sku) {
        inlineActionsEl.classList.remove('hidden')
        if (addSkuBtn) addSkuBtn.classList.remove('hidden')
        if (moveBinBtn) moveBinBtn.classList.add('hidden')
      } else {
        inlineActionsEl.classList.add('hidden')
      }
    }
  }

  // Feature 3: trigger photo lookup (debounced)
  debouncedPhotoLookup(sku)

  render()
}

// =========================================================
// FEATURE 2: Edit records before sync
// =========================================================
const startEdit = (index) => {
  const rows = db.get()
  if (index < 0 || index >= rows.length) return
  const r = rows[index]
  if (r._status === 'SYNCED') {
    toast('Cannot edit synced records', 'warning')
    return
  }
  editingIndex = index
  $('sku').value = r.SKU || ''
  $('bin').value = r.Bin || ''
  $('countedQty').value = r.CountedQty ?? ''
  $('notes').value = r.Notes || ''
  $('expectedQty').value = r.ExpectedQty ?? ''

  const hit = inventoryCache.get(r.SKU)
  $('salespadQty').value = (hit?.salespadQty !== null && hit?.salespadQty !== undefined) ? hit.salespadQty : '\u2014'

  if (r._photoLocal) {
    $('photoPreview').src = r._photoLocal
    $('photoPreview').classList.remove('hidden')
  } else {
    $('photoPreview').classList.add('hidden')
  }

  // Show edit UI
  $('editBanner').classList.remove('hidden')
  $('barSave').textContent = 'Update Count'
  $('barCancelEdit2').classList.remove('hidden')

  validateAndCompute()
  toast('Editing record for ' + r.SKU, 'info')
  document.querySelector('.count-card')?.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

const cancelEdit = () => {
  editingIndex = -1
  $('sku').value = ''
  $('countedQty').value = ''
  $('notes').value = ''
  $('photo').value = ''
  $('photoPreview').classList.add('hidden')
  $('expectedQty').value = ''
  $('salespadQty').value = ''
  $('existingPhotoWrap').classList.add('hidden')
  setBadge('lookupState', '\u2014', '')
  setBadge('binState', '\u2014', '')
  setBadge('varState', '\u2014', '')

  // Hide inline elements
  const lastCountEl = $('lastCountInfo')
  if (lastCountEl) lastCountEl.classList.add('hidden')
  const inlineActionsEl = $('inlineActions')
  if (inlineActionsEl) inlineActionsEl.classList.add('hidden')

  $('editBanner').classList.add('hidden')
  $('barSave').textContent = 'Save Count'
  $('barCancelEdit2').classList.add('hidden')

  render()
  toast('Edit cancelled', 'info')
}

// === Render ===
const render = () => {
  const rows = db.get()
  const pending = rows.filter(r => r._status !== 'SYNCED').length

  setBadge('chipQueue', String(pending), pending > 0 ? 'warn' : 'good')
  $('kpiTotal').textContent = rows.length
  $('kpiPending').textContent = pending
  $('kpiMatched').textContent = rows.filter(r => r._matched).length
  $('kpiVariance').textContent = rows.filter(r => Number(r.Variance) !== 0).length

  const syncCountEl = $('barSyncCount')
  if (syncCountEl) syncCountEl.textContent = '(' + pending + ')'

  const list = $('recordList')
  if (!list) return

  if (rows.length === 0) {
    list.innerHTML = '<div class="empty-state"><div class="empty-state-icon">\uD83D\uDCCB</div><div class="empty-state-text">No records yet. Start counting!</div></div>'
    return
  }

  // Show last 40 records, newest first, preserving original index
  const last = rows.map((r, i) => ({ ...r, _idx: i })).slice(-40).reverse()
  list.innerHTML = last.map(r => {
    const st = (r._status || 'LOCAL').toLowerCase()
    const variance = Number(r.Variance)
    const varClass = variance === 0 ? 'var-zero' : (variance > 0 ? 'var-positive' : 'var-negative')
    const varSign = variance > 0 ? '+' : ''
    const canEdit = r._status !== 'SYNCED'
    const isEditing = r._idx === editingIndex
    return '<div class="record-item' + (isEditing ? '" style="border-color:var(--warning)"' : '"') + '>' +
      '<div class="record-top">' +
        '<span class="record-sku">' + (r.SKU || '\u2014') + '</span>' +
        '<span class="record-status s-' + st + '">' + (r._status || 'LOCAL') + '</span>' +
      '</div>' +
      '<div class="record-details">' +
        '<span>Bin: ' + (r.Bin || '\u2014') + '</span>' +
        '<span>Exp: ' + r.ExpectedQty + '</span>' +
        '<span>Count: ' + r.CountedQty + '</span>' +
        '<span class="' + varClass + '">Var: ' + varSign + r.Variance + '</span>' +
      '</div>' +
      '<div class="record-bottom">' +
        '<span class="record-time">' + timeAgo(r.Timestamp) + '</span>' +
        (canEdit ? '<button class="btn-edit" data-index="' + r._idx + '">' + (isEditing ? 'Editing\u2026' : 'Edit') + '</button>' : '') +
      '</div>' +
    '</div>'
  }).join('')
}

// === File Helpers ===
const fileToDataUrl = file => new Promise((res, rej) => {
  const fr = new FileReader()
  fr.onload = () => res(fr.result)
  fr.onerror = () => rej(fr.error)
  fr.readAsDataURL(file)
})

const dataUrlToBlob = dataUrl => {
  const [h, b] = dataUrl.split(',')
  const m = /data:(.*?);base64/.exec(h)
  const mime = m ? m[1] : 'image/jpeg'
  const bin = atob(b)
  const u8 = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
  return new Blob([u8], { type: mime })
}

// === Scanner ===
const toggleScan = async () => {
  if (scanning) { await stopScan() } else { await startScan() }
}

const startScan = async () => {
  if (scanner) return
  scanner = new Html5Qrcode('reader')
  $('reader').classList.remove('hidden')
  $('btnToggleScan').textContent = 'Stop Scanner'
  scanning = true
  await scanner.start(
    { facingMode: 'environment' },
    { fps: 12, qrbox: { width: 260, height: 170 }, experimentalFeatures: { useBarCodeDetectorIfSupported: true } },
    text => {
      const target = document.querySelector('input[name="scanTarget"]:checked')?.value || 'sku'
      $(target).value = text.trim()
      toast('Scanned ' + target.toUpperCase() + ': ' + text.trim(), 'success')
      validateAndCompute()
      // Auto-close camera after successful scan
      stopScan()
    },
    () => {}
  )
}

const stopScan = async () => {
  if (!scanner) return
  try { await scanner.stop() } catch (e) { /* ignore */ }
  try { await scanner.clear() } catch (e) { /* ignore */ }
  scanner = null
  scanning = false
  $('reader').classList.add('hidden')
  $('btnToggleScan').textContent = 'Start Scanner'
}

// === Save Local (with edit support) ===
const saveLocal = async () => {
  const sku = $('sku').value.trim()
  if (!sku) {
    toast('Enter a SKU first', 'warning')
    return
  }

  const bin = $('bin').value.trim()
  const notes = $('notes').value.trim()
  const device = account?.name || account?.username || 'Unknown'
  const hit = inventoryCache.get(sku)
  const expected = hit ? Number(hit.expectedQty || 0) : 0
  const counted = Number($('countedQty').value || 0)
  const variance = counted - expected

  const file = $('photo').files?.[0] || null
  const photoLocal = file ? await fileToDataUrl(file) : ''

  const rows = db.get()

  // Feature 2: Update existing record when editing
  if (editingIndex >= 0 && editingIndex < rows.length) {
    const existing = rows[editingIndex]
    existing.SKU = sku
    existing.Bin = bin
    existing.ExpectedQty = expected
    existing.CountedQty = counted
    existing.Variance = variance
    existing.Notes = notes
    existing.Device = device
    existing._matched = !!hit
    if (photoLocal) existing._photoLocal = photoLocal
    if (existing._status === 'FAILED') existing._status = 'LOCAL'
    db.set(rows)
    toast('Record updated', 'success')
    cancelEdit()
    return
  }

  // Normal save — new record
  const row = {
    Timestamp: nowIso(),
    SKU: sku,
    Bin: bin,
    ExpectedQty: expected,
    CountedQty: counted,
    Variance: variance,
    PhotoUrl: '',
    Device: device,
    Notes: notes,
    _photoLocal: photoLocal,
    _status: 'LOCAL',
    _matched: !!hit
  }

  rows.push(row)
  db.set(rows)

  // Clear form but keep bin for batch counting same location
  $('sku').value = ''
  $('countedQty').value = ''
  $('notes').value = ''
  $('photo').value = ''
  $('photoPreview').classList.add('hidden')
  $('expectedQty').value = ''
  $('salespadQty').value = ''
  $('existingPhotoWrap').classList.add('hidden')
  const _lci = $('lastCountInfo'); if (_lci) _lci.classList.add('hidden')
  const _iae = $('inlineActions'); if (_iae) _iae.classList.add('hidden')
  setBadge('lookupState', '\u2014', '')
  setBadge('binState', '\u2014', '')
  setBadge('varState', '\u2014', '')

  toast('Count saved', 'success')
  render()
  $('sku').focus()
}

// === Sync (with OnHand update) ===
const syncNow = async () => {
  const rows = db.get()
  const pending = rows.filter(r => r._status !== 'SYNCED')
  if (!pending.length) {
    toast('Nothing to sync', 'info')
    return
  }
  if (!driveId || !itemId) await resolveWorkbook()

  let synced = 0
  for (const r of pending) {
    try {
      r._status = 'UPLOADING'
      db.set(rows); render()

      let photoUrl = ''
      if (r._photoLocal) {
        const blob = dataUrlToBlob(r._photoLocal)
        const file = new File([blob], 'photo.jpg', { type: blob.type || 'image/jpeg' })
        photoUrl = await uploadPhotoToOneDrive(file, r.SKU)
      }
      r.PhotoUrl = photoUrl
      r._status = 'WRITING'
      db.set(rows); render()

      await addRowToExcel(r)

      // Feature 4: Update OnHand row with last counted qty + timestamp
      try {
        await updateOnHandRow(r.SKU, r.CountedQty, r.Timestamp)
      } catch (e) {
        console.warn('OnHand update failed for ' + r.SKU + ':', e.message)
      }

      r._status = 'SYNCED'
      r._photoLocal = ''
      db.set(rows); render()
      synced++
    } catch (e) {
      r._status = 'FAILED'
      db.set(rows); render()
      toast('Sync failed: ' + r.SKU + ' \u2014 ' + (e?.message || e), 'error')
      throw e
    }
  }
  toast(synced + ' record(s) synced', 'success')

  // Auto-refresh KPI sheet in background (non-blocking)
  toast('KPI sheet updating in background...', 'info')
  refreshKPIs()
    .then(() => toast('KPI sheet updated', 'success'))
    .catch(e => {
      console.warn('KPI refresh failed:', e.message)
      toast('KPI update failed — try syncing again', 'warning')
    })
}

// =========================================================
// FEATURE 5: KPI sheet with 5 charts
// =========================================================
const refreshKPIs = async () => {
  if (!driveId || !itemId) await resolveWorkbook()

  // 1. Delete existing KPI sheet (if exists)
  try {
    await graph('DELETE', wbUrl() + '/worksheets/KPI')
  } catch (e) { /* sheet doesn't exist — fine */ }

  // 2. Create fresh KPI sheet
  await graph('POST', wbUrl() + '/worksheets', { name: 'KPI' })

  const kpiBase = wbUrl() + '/worksheets/KPI'

  // ---- SECTION A: Summary metrics (A1:B11) ----
  const summaryFormulas = [
    ['Metric', 'Value'],
    ['Total SKUs', '=ROWS(OnHand[SKU])'],
    ['SKUs Counted', '=COUNTA(OnHand[LastCountedQty])'],
    ['Coverage %', '=IF(ROWS(OnHand[SKU])=0,0,ROUND(COUNTA(OnHand[LastCountedQty])/ROWS(OnHand[SKU])*100,1))'],
    ['Total Counts', '=ROWS(CycleCounts[SKU])'],
    ['Perfect Counts', '=COUNTIF(CycleCounts[Variance],0)'],
    ['Accuracy %', '=IF(ROWS(CycleCounts[SKU])=0,0,ROUND(COUNTIF(CycleCounts[Variance],0)/ROWS(CycleCounts[SKU])*100,1))'],
    ['Positive Variance', '=COUNTIF(CycleCounts[Variance],">"&0)'],
    ['Negative Variance', '=COUNTIF(CycleCounts[Variance],"<"&0)'],
    ['Avg Abs Variance', '=IF(ROWS(CycleCounts[SKU])=0,0,ROUND(SUMPRODUCT(ABS(CycleCounts[Variance]))/ROWS(CycleCounts[SKU]),2))'],
    ['Devices Used', '=ROWS(UNIQUE(CycleCounts[Device]))']
  ]
  await graph('PATCH', kpiBase + "/range(address='A1:B11')", { formulas: summaryFormulas })

  // ---- SECTION B: Variance Distribution (D1:E6) → ColumnClustered ----
  const varianceData = [
    ['Variance Range', 'Count'],
    ['Zero (Perfect)', '=COUNTIF(CycleCounts[Variance],0)'],
    ['+1 to +5', '=COUNTIFS(CycleCounts[Variance],">"&0,CycleCounts[Variance],"<="&5)'],
    ['+6 or more', '=COUNTIF(CycleCounts[Variance],">"&5)'],
    ['-1 to -5', '=COUNTIFS(CycleCounts[Variance],"<"&0,CycleCounts[Variance],">="&-5)'],
    ['-6 or less', '=COUNTIF(CycleCounts[Variance],"<"&-5)']
  ]
  await graph('PATCH', kpiBase + "/range(address='D1:E6')", { formulas: varianceData })

  // ---- SECTION C: Accuracy Breakdown (D9:E10) → Pie ----
  const accuracyData = [
    ['Category', 'Count'],
    ['Perfect (0 Variance)', '=COUNTIF(CycleCounts[Variance],0)'],
    ['Imperfect (Non-zero)', '=COUNTIF(CycleCounts[Variance],"<>"&0)']
  ]
  await graph('PATCH', kpiBase + "/range(address='D9:E11')", { formulas: accuracyData })

  // ---- SECTION D: Coverage Breakdown (D14:E15) → Pie ----
  const coverageData = [
    ['Status', 'Count'],
    ['Counted', '=COUNTA(OnHand[LastCountedQty])'],
    ['Not Counted', '=ROWS(OnHand[SKU])-COUNTA(OnHand[LastCountedQty])']
  ]
  await graph('PATCH', kpiBase + "/range(address='D14:E16')", { formulas: coverageData })

  // ---- SECTION E: Counts by Device (computed values in G1:H?) ----
  // Use local synced data since UNIQUE() + COUNTIF by device is tricky cross-table
  const rows = db.get().filter(r => r._status === 'SYNCED')
  const deviceMap = new Map()
  for (const r of rows) {
    const d = r.Device || 'Unknown'
    deviceMap.set(d, (deviceMap.get(d) || 0) + 1)
  }
  const deviceEntries = [...deviceMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10)
  if (deviceEntries.length === 0) deviceEntries.push(['No Data', 0])
  const deviceGrid = [['Device', 'Counts'], ...deviceEntries]
  const deviceEndRow = 1 + deviceEntries.length
  await graph('PATCH', kpiBase + "/range(address='G1:H" + deviceEndRow + "')", {
    values: deviceGrid
  })

  // ---- SECTION F: Top 10 Variance SKUs (computed in G' + offset) ----
  const topVarOffset = deviceEndRow + 2
  const allCounts = db.get()
  const skuVarMap = new Map()
  for (const r of allCounts) {
    const absVar = Math.abs(Number(r.Variance || 0))
    if (absVar > 0) {
      const existing = skuVarMap.get(r.SKU) || 0
      if (absVar > existing) skuVarMap.set(r.SKU, absVar)
    }
  }
  const topVar = [...skuVarMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10)
  if (topVar.length === 0) topVar.push(['No Variance', 0])
  const topVarGrid = [['SKU', 'Abs Variance'], ...topVar]
  const topVarEndRow = topVarOffset + topVar.length
  await graph('PATCH', kpiBase + "/range(address='G" + topVarOffset + ":H" + topVarEndRow + "')", {
    values: topVarGrid
  })

  // ---- Format headers bold ----
  try {
    await graph('PATCH', kpiBase + "/range(address='A1:B1')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='D1:E1')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='D9:E9')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='D14:E14')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='G1:H1')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='G" + topVarOffset + ":H" + topVarOffset + "')/format/font", { bold: true })
  } catch (e) {
    console.warn('Header formatting failed:', e.message)
  }

  // ---- SECTION G: Chart Descriptions (A13:B35) ----
  const descRows = [
    ['', ''],
    ['CHART GUIDE', ''],
    ['', ''],
    ['📊 Variance Distribution', ''],
    ['', 'Shows how count variances are spread across ranges. "Zero" means the physical count matched'],
    ['', 'the expected quantity exactly. Positive means you found MORE than expected; negative means LESS.'],
    ['', 'Ideally the "Zero (Perfect)" bar should be the tallest. Large bars on +6/−6 indicate systemic issues.'],
    ['', ''],
    ['🎯 Count Accuracy', ''],
    ['', 'Pie chart showing the ratio of perfect counts (zero variance) vs imperfect counts.'],
    ['', 'Target: 95%+ accuracy. If the "Imperfect" slice is large, investigate top-variance SKUs below.'],
    ['', ''],
    ['📦 SKU Coverage', ''],
    ['', 'Pie chart showing how many of your inventory SKUs have been cycle-counted at least once.'],
    ['', 'Target: 100% coverage over the count cycle. A large "Not Counted" slice means more SKUs need attention.'],
    ['', ''],
    ['👥 Counts by Device', ''],
    ['', 'Bar chart showing how many counts each team member (device) has submitted.'],
    ['', 'Use this to balance workload across counters and identify top contributors.'],
    ['', ''],
    ['⚠️ Top Variance SKUs', ''],
    ['', 'Horizontal bar chart of the 10 SKUs with the largest absolute variance.'],
    ['', 'These are your problem SKUs — investigate for misplacement, theft, receiving errors, or bad bin locations.'],
    ['', 'Recount these SKUs first to confirm or correct the discrepancy.']
  ]
  const descEndRow = 12 + descRows.length
  await graph('PATCH', kpiBase + "/range(address='A13:B" + descEndRow + "')", { values: descRows })

  // Format chart guide header
  try {
    await graph('PATCH', kpiBase + "/range(address='A14:B14')/format/font", { bold: true, size: 14 })
    // Format each chart title row bold
    await graph('PATCH', kpiBase + "/range(address='A16:A16')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='A21:A21')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='A25:A25')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='A29:A29')/format/font", { bold: true })
    await graph('PATCH', kpiBase + "/range(address='A33:A33')/format/font", { bold: true })
    // Make description text italic and gray
    await graph('PATCH', kpiBase + "/range(address='B17:B19')/format/font", { italic: true, color: '#666666' })
    await graph('PATCH', kpiBase + "/range(address='B22:B23')/format/font", { italic: true, color: '#666666' })
    await graph('PATCH', kpiBase + "/range(address='B26:B27')/format/font", { italic: true, color: '#666666' })
    await graph('PATCH', kpiBase + "/range(address='B30:B30')/format/font", { italic: true, color: '#666666' })
    await graph('PATCH', kpiBase + "/range(address='B34:B36')/format/font", { italic: true, color: '#666666' })
  } catch (e) {
    console.warn('Description formatting failed:', e.message)
  }

  // ---- Auto-fit columns ----
  try {
    await graph('POST', kpiBase + "/range(address='A:H')/format/autofitColumns", {})
  } catch (e) {
    console.warn('AutofitColumns failed:', e.message)
  }

  // ---- Create Charts ----
  // Helper: create chart, set position, then set title (separate API calls)
  const chartsBase = kpiBase + '/charts'
  let chartCount = 0

  const createChart = async (type, sourceData, titleText, h, w, t, l) => {
    const c = await graph('POST', chartsBase + '/add', { type, sourceData, seriesBy: 'Columns' })
    if (c?.name) {
      const chartPath = chartsBase + "('" + c.name + "')"
      // Position & size (direct properties)
      await graphSafe('PATCH', chartPath, { height: h, width: w, top: t, left: l })
      // Title (navigation property — separate endpoint)
      await graphSafe('PATCH', chartPath + '/title', { text: titleText, visible: true })
    }
    chartCount++
  }

  // Chart 1: Variance Distribution
  try { await createChart('ColumnClustered', 'KPI!D1:E6', 'Variance Distribution', 280, 420, 0, 550) }
  catch (e) { console.warn('Variance chart failed:', e.message) }

  // Chart 2: Accuracy Pie
  try { await createChart('Pie', 'KPI!D9:E11', 'Count Accuracy', 280, 420, 300, 550) }
  catch (e) { console.warn('Accuracy chart failed:', e.message) }

  // Chart 3: Coverage Pie
  try { await createChart('Pie', 'KPI!D14:E16', 'SKU Coverage', 280, 420, 600, 550) }
  catch (e) { console.warn('Coverage chart failed:', e.message) }

  // Chart 4: Counts by Device
  try { await createChart('ColumnClustered', 'KPI!G1:H' + deviceEndRow, 'Counts by Device', 280, 420, 0, 1000) }
  catch (e) { console.warn('Device chart failed:', e.message) }

  // Chart 5: Top Variance SKUs (horizontal bar)
  try { await createChart('BarClustered', 'KPI!G' + topVarOffset + ':H' + topVarEndRow, 'Top Variance SKUs', 280, 420, 300, 1000) }
  catch (e) { console.warn('Top variance chart failed:', e.message) }

  toast('KPI sheet refreshed — ' + chartCount + ' chart(s) created', 'success')
}

// === Clear ===
const clearLocal = async () => {
  const rows = db.get()
  if (!rows.length) {
    toast('Already empty', 'info')
    return
  }
  const ok = await showConfirm('Clear all records?', 'This will delete ' + rows.length + ' local record(s). This cannot be undone.', 'Delete', 'var(--danger)')
  if (!ok) return
  if (editingIndex >= 0) cancelEdit()
  db.set([])
  render()
  toast('All local records cleared', 'info')
}

// === Setup Collapse ===
const toggleSetup = () => {
  setupCollapsed = !setupCollapsed
  const body = $('setupBody')
  const header = $('toggleSetup')
  if (setupCollapsed) {
    body.classList.add('collapsed')
    header.classList.add('collapsed')
  } else {
    body.classList.remove('collapsed')
    header.classList.remove('collapsed')
  }
}

// === Auto-clear synced records older than 24h ===
const autoCleanSynced = () => {
  const rows = db.get()
  const cutoff = Date.now() - 24 * 60 * 60 * 1000
  const before = rows.length
  const filtered = rows.filter(r => {
    if (r._status !== 'SYNCED') return true
    const ts = new Date(r.Timestamp).getTime()
    return ts > cutoff
  })
  if (filtered.length < before) {
    db.set(filtered)
    console.log('Auto-cleaned ' + (before - filtered.length) + ' old synced record(s)')
  }
}

// === Wire Everything ===
const wire = async () => {
  // Auto-clean old synced records on startup
  autoCleanSynced()

  // Service worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('service-worker.js')
      .then(() => console.log('SW registered'))
      .catch(e => console.warn('SW error', e))
  }

  // Theme
  const savedTheme = localStorage.getItem('theme') || 'dark'
  applyTheme(savedTheme)
  $('themeToggle').onclick = toggleTheme

  if (!requireMsal()) return
  await msalInit()

  setBadge('chipAuth', 'Signed out', 'bad')
  setBadge('chipCache', '0', '')
  render()

  await handleRedirect()

  // Auto-reconnect: if user is signed in and has a saved workbook link,
  // automatically connect and load inventory so they're ready to go
  if (account) {
    const savedLink = localStorage.getItem('shareLink')
    if (savedLink) {
      $('shareLink').value = savedLink
      try {
        toast('Reconnecting to workbook...', 'info')
        await resolveWorkbook()
        await loadInventoryCache()
      } catch (e) {
        console.warn('Auto-reconnect failed:', e.message)
        toast('Auto-connect failed \u2014 try Setup manually', 'warning')
      }
    }
  }

  // Setup toggle
  $('toggleSetup').onclick = toggleSetup

  // Step 1: Sign in
  $('btnSignIn').onclick = withLoading($('btnSignIn'), async () => {
    await msalApp.loginRedirect({ scopes: CFG.scopes })
  }, 'Signing in...')

  // Step 2: Connect workbook
  $('btnLoadWorkbook').onclick = withLoading($('btnLoadWorkbook'), async () => {
    await resolveWorkbook()
  }, 'Connecting...')

  // Step 3: Load inventory
  $('btnRefreshInventory').onclick = withLoading($('btnRefreshInventory'), async () => {
    await resolveWorkbook()
    await loadInventoryCache()
  }, 'Loading...')

  // Scanner
  $('btnToggleScan').onclick = () => toggleScan().catch(e => toast('Scanner: ' + (e?.message || e), 'error'))

  // Photo
  $('btnPhoto').onclick = () => $('photo').click()
  $('photo').addEventListener('change', async () => {
    const f = $('photo').files?.[0]
    if (!f) { $('photoPreview').classList.add('hidden'); return }
    const url = await fileToDataUrl(f)
    $('photoPreview').src = url
    $('photoPreview').classList.remove('hidden')
  })

  // Bottom bar
  $('barSave').onclick = () => saveLocal().catch(e => toast('Save error: ' + (e?.message || e), 'error'))
  const syncFn = withLoading($('barSync'), syncNow, 'Syncing...')
  $('barSync').onclick = async () => { await syncFn(); render() }

  // Feature 2: Edit cancel buttons
  $('btnCancelEdit').onclick = cancelEdit
  $('barCancelEdit2').onclick = cancelEdit

  // Feature 2: Event delegation for edit buttons in record list
  $('recordList').addEventListener('click', e => {
    const btn = e.target.closest('.btn-edit')
    if (!btn) return
    const idx = parseInt(btn.dataset.index, 10)
    if (!isNaN(idx)) startEdit(idx)
  })

  // Inline inventory management
  $('btnAddSku').onclick = () => addSkuToInventory().catch(e => toast('Add SKU error: ' + (e?.message || e), 'error'))
  $('btnMoveBin').onclick = () => moveSkuBin().catch(e => toast('Move bin error: ' + (e?.message || e), 'error'))

  // Records actions
  $('btnClearLocal').onclick = clearLocal

  // Form auto-validation
  $('sku').addEventListener('input', validateAndCompute)
  $('bin').addEventListener('input', validateAndCompute)
  $('countedQty').addEventListener('input', validateAndCompute)

  // Persist settings
  const prevLink = localStorage.getItem('shareLink') || ''
  if (prevLink) $('shareLink').value = prevLink
  $('shareLink').addEventListener('input', () => localStorage.setItem('shareLink', $('shareLink').value))

  // Scan target persistence
  const savedScanTarget = localStorage.getItem('scanTarget') || 'sku'
  const radio = document.querySelector('input[name="scanTarget"][value="' + savedScanTarget + '"]')
  if (radio) radio.checked = true
  document.querySelectorAll('input[name="scanTarget"]').forEach(r => {
    r.addEventListener('change', () => localStorage.setItem('scanTarget', r.value))
  })

  console.log('Cycle Count ready')
}

window.addEventListener('load', () => wire().catch(e => toast('Init error: ' + (e?.message || e), 'error')))
