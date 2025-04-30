// xlsx-to-json.js  –  ESM, header-sanitising exporter
import { existsSync, mkdirSync, writeFileSync } from 'fs'
import { resolve, join, relative } from 'path'
import xlsx from 'xlsx'

/* ▸ EDIT THESE THREE CONSTANTS IF NEEDED ◂ */
const INPUT_FILE = resolve(process.cwd(), 'New Microsoft Excel Worksheet.xlsx')
const OUTPUT_DIR = resolve(process.cwd(), 'exports')
const CHUNK_SIZE = 1_000

/* ───────────────── helpers ───────────────── */
const sanitizePath = s => s.replace(/[^a-z0-9-]/gi, '_').toLowerCase()
const ensureDir = d => {
  if (!existsSync(d)) mkdirSync(d, { recursive: true })
}

const BAD_KEY_CHARS = /[.#$/\[\]\s]/g // Firebase-illegal + spaces
function sanitizeKey(k) {
  return k
    .trim()
    .replace(BAD_KEY_CHARS, '_')
    .replace(/_+/g, '_') // collapse runs of "_"s
    .replace(/^_+|_+$/g, '') // strip leading/trailing "_"
    .toLowerCase()
}
/* ─────────────────────────────────────────── */

ensureDir(OUTPUT_DIR)
console.time('xlsx→json')

const wb = xlsx.readFile(INPUT_FILE, { cellDates: true, dense: true })
const { utils } = xlsx

wb.SheetNames.forEach((sheetName, idx) => {
  const ws = wb.Sheets[sheetName]
  if (!ws) return // sheet object missing: skip

  /* 1️⃣  read the sheet twice if necessary (skip blank row 1) */
  let rows = utils.sheet_to_json(ws, { defval: null, blankrows: false })
  if (rows.length === 0)
    rows = utils.sheet_to_json(ws, { defval: null, blankrows: false, range: 1 })

  const safeName = `${String(idx).padStart(2, '0')}_${sanitizePath(sheetName)}`
  const sheetDir = join(OUTPUT_DIR, safeName)
  ensureDir(sheetDir)

  if (rows.length === 0) {
    // genuinely empty
    writeFileSync(join(sheetDir, 'empty.json'), '[]')
    console.log(`⚠  ${sheetName} is empty`)
    return
  }

  /* 2️⃣  build a header-sanitising map */
  const keyMap = {}
  Object.keys(rows[0]).forEach(orig => {
    let safe = sanitizeKey(orig)
    if (safe === '') safe = 'col' // all bad chars?  call it "col"
    while (Object.values(keyMap).includes(safe)) {
      safe += '_dup' // dedupe on collision
    }
    keyMap[orig] = safe
  })

  /* 3️⃣  rewrite every row with safe keys */
  const safeRows = rows.map(r => {
    const obj = {}
    for (const [origKey, val] of Object.entries(r)) {
      obj[keyMap[origKey]] = val
    }
    return obj
  })

  /* 4️⃣  chunk & write */
  for (let part = 0; part * CHUNK_SIZE < safeRows.length; part++) {
    const slice = safeRows.slice(part * CHUNK_SIZE, (part + 1) * CHUNK_SIZE)
    const file = join(sheetDir, `${safeName}.${String(part).padStart(4, '0')}.json`)
    writeFileSync(file, JSON.stringify(slice, null, 2))
    console.log(
      `✔  ${sheetName} rows ${part * CHUNK_SIZE}-${
        part * CHUNK_SIZE + slice.length - 1
      } → ${relative(process.cwd(), file)}`
    )
  }
})

console.timeEnd('xlsx→json')
