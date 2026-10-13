import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'

/* ------------------------------------------------------------------ api */
const api = {
  categories: () => fetch('/api/categories').then(r => r.json()),
  jobs: () => fetch('/api/jobs').then(r => r.json()),
  submit: (form) => fetch('/api/jobs', { method: 'POST', body: form }).then(async r => {
    if (!r.ok) throw new Error((await r.json()).detail || 'Submit failed')
    return r.json()
  }),
  stop: (id) => fetch(`/api/jobs/${id}/stop`, { method: 'POST' }),
  updateCheck: () => fetch('/api/update').then(r => r.json()),
  updateStatus: () => fetch('/api/update/status').then(r => r.json()),
  updatePost: async (url) => {
    const r = await fetch(url, { method: 'POST' })
    const body = await r.json().catch(() => ({}))
    if (!r.ok) throw new Error(body.detail || 'Update failed')
    return body
  },
  remove: (id) => fetch(`/api/jobs/${id}`, { method: 'DELETE' }),
}

/* ------------------------------------------------------------------ bits */
function Cube({ size = 26, speed = 9 }) {
  const s = size
  const style = { '--s': `${s}px`, animationDuration: `${speed}s` }
  return (
    <div className="cube-scene" style={{ width: s * 1.7, height: s * 1.7 }}>
      <div className="cube" style={style}>
        <div className="face f-front" /><div className="face f-back" />
        <div className="face f-right" /><div className="face f-left" />
        <div className="face f-top" /><div className="face f-bottom" />
      </div>
    </div>
  )
}

const STATUS_META = {
  queued: { label: 'Queued', cls: 'chip-queued' },
  running: { label: 'Processing', cls: 'chip-running' },
  done: { label: 'Completed', cls: 'chip-done' },
  error: { label: 'Failed', cls: 'chip-error' },
  stopped: { label: 'Stopped', cls: 'chip-stopped' },
}

function fmtTime(ts) {
  if (!ts) return ''
  return new Date(ts * 1000).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
}

/* ------------------------------------------------------------------ app */
export default function App() {
  const [theme, setTheme] = useState(() => {
    const q = new URLSearchParams(window.location.search).get('theme')
    if (q === 'dark' || q === 'light') return q
    return localStorage.getItem('docpilot-theme') || 'light'
  })
  const [cats, setCats] = useState([])
  const [category, setCategory] = useState('')
  const [mode, setMode] = useState('upload')        // 'upload' | 'link'
  const [link, setLink] = useState('')
  const [link2, setLink2] = useState('')
  const [files, setFiles] = useState([])
  const [files2, setFiles2] = useState([])
  const [startPage, setStartPage] = useState('')
  const [endPage, setEndPage] = useState('')
  const [coverDate, setCoverDate] = useState('')
  const [coverIndex, setCoverIndex] = useState('')
  const [coverShorts, setCoverShorts] = useState('')
  const [coverQr, setCoverQr] = useState('')
  const [coverIndex2, setCoverIndex2] = useState('')
  // Test Series details: { series, type, testno, subject, code, marks, permark, time, year }
  const [details, setDetails] = useState({})
  const [jobs, setJobs] = useState([])
  const [busy, setBusy] = useState(false)
  const [toasts, setToasts] = useState([])
  const [dragOver, setDragOver] = useState(false)
  // in-app update: what /api/update said, and the download / install step
  const [update, setUpdate] = useState(null)
  const [upd, setUpd] = useState({ state: 'idle', progress: 0 })
  const fileInput = useRef(null)
  const fileInput2 = useRef(null)

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('docpilot-theme', theme)
  }, [theme])

  useEffect(() => {
    api.categories().then(cs => {
      setCats(cs)
      // ?cat=<id> deep-links a category (handy for bookmarks and screenshots)
      const wanted = new URLSearchParams(window.location.search).get('cat')
      const pick = cs.find(c => c.id === wanted && c.ready) || cs.find(c => c.ready)
      if (pick) setCategory(pick.id)
    }).catch(() => toast('Cannot reach the DocPilot backend', 'error'))
  }, [])

  useEffect(() => {
    let alive = true
    const tick = async () => {
      try {
        const j = await api.jobs()
        if (alive) setJobs(j)
      } catch { /* backend briefly unavailable */ }
    }
    tick()
    const iv = setInterval(tick, 1200)
    return () => { alive = false; clearInterval(iv) }
  }, [])

  const toast = useCallback((msg, kind = 'info') => {
    const id = Math.random().toString(36).slice(2)
    setToasts(t => [...t, { id, msg, kind }])
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 5200)
  }, [])

  const cat = useMemo(() => cats.find(c => c.id === category) || null, [cats, category])
  const needsTwo = cat?.files === 2
  const pageMode = cat?.pages || 'none'
  const hasCover = !!cat?.cover
  const dateRequired = hasCover && !!cat?.cover_date_required
  // which front-page blocks this category takes (older categories: the three)
  const coverFields = cat?.cover_fields || ['date', 'index', 'shorts']
  const hasField = (f) => hasCover && coverFields.includes(f)
  // the two file slots' names (the bilingual Test Series default)
  const slots = cat?.slots || ['English paper', 'हिंदी paper']
  const isPair = !!cat?.slots
  // which Test Series details this category prints (Prelims only)
  const testFields = cat?.test_fields || []
  const hasTest = (f) => testFields.includes(f)
  const setDetail = (key) => (e) => setDetails(d => ({ ...d, [key]: e.target.value }))
  // year / marks / time: plain text boxes that keep digits only - a number
  // box jumps to its min on a stray arrow key or mouse-wheel scroll
  const setDigits = (key) => (e) => setDetails(d => ({ ...d, [key]: e.target.value.replace(/\D/g, '') }))
  const setDecimal = (key) => (e) => setDetails(d => ({ ...d, [key]: e.target.value.replace(/[^\d.]/g, '') }))

  const addTo = useCallback((setter) => (list) => {
    const ok = [], bad = []
    for (const f of list) (f.name.toLowerCase().endsWith('.docx') ? ok : bad).push(f)
    if (bad.length) toast(`${bad.length} file(s) skipped — only .docx is supported`, 'error')
    if (ok.length) setter(prev => [...prev, ...ok.filter(f => !prev.some(p => p.name === f.name && p.size === f.size))])
  }, [toast])

  const addFiles = useCallback((list) => addTo(setFiles)(list), [addTo])
  const addFiles2 = useCallback((list) => addTo(setFiles2)(list), [addTo])

  // page fields are per-category, so clear them when the category changes
  useEffect(() => {
    setStartPage(''); setEndPage(''); setFiles2([]); setLink2('')
    setCoverDate(''); setCoverIndex(''); setCoverShorts(''); setCoverQr(''); setCoverIndex2('')
    setDetails({})
  }, [category])

  const addPageFields = (form) => {
    if (pageMode !== 'none' && startPage) form.append('start_page', startPage)
    if (pageMode === 'start_end' && endPage) form.append('end_page', endPage)
    // front-page blocks: an empty block is filled from the document itself
    if (hasCover) {
      if (coverDate.trim()) form.append('cover_date', coverDate.trim())
      if (coverIndex.trim()) form.append('cover_index', coverIndex)
      if (coverShorts.trim() && hasField('shorts')) form.append('cover_shorts', coverShorts)
      if (coverQr.trim() && hasField('qr')) form.append('cover_qr', coverQr.trim())
      if (coverIndex2.trim() && hasField('index2')) form.append('cover_index2', coverIndex2)
    }
    // Test Series details: an empty one keeps the file name / template value
    for (const f of testFields) {
      const v = (details[f] || '').trim()
      if (v) form.append(`ts_${f}`, v)
    }
  }

  const submit = async () => {
    if (!category) return toast('Pick a category first', 'error')
    if (pageMode !== 'none' && !startPage) {
      return toast('Enter the start page for this category', 'error')
    }
    if (dateRequired && !coverDate.trim()) {
      return toast('Enter the front-page date (e.g. 17 August 2026)', 'error')
    }
    if (hasTest('year') && details.year?.trim() && !/^20\d\d$/.test(details.year.trim())) {
      return toast('Target year must be a year like 2027', 'error')
    }
    setBusy(true)
    try {
      if (mode === 'link') {
        const url = link.trim()
        if (!/(drive|docs)\.google\.com/.test(url)) {
          toast('Paste a Google Drive / Google Docs share link', 'error'); setBusy(false); return
        }
        if (needsTwo && !link2.trim()) {
          toast(isPair ? `This category needs both files — add the ${slots[1]} link too`
                       : 'This category needs both papers — add the Hindi link too', 'error')
          setBusy(false); return
        }
        const form = new FormData()
        form.append('category_id', category)
        form.append('link', url)
        if (needsTwo) form.append('link2', link2.trim())
        addPageFields(form)
        await api.submit(form)
        setLink(''); setLink2('')
        toast('Link queued ✓', 'ok')
      } else {
        if (!files.length) { toast('Add at least one .docx file', 'error'); setBusy(false); return }
        if (needsTwo) {
          if (files.length !== 1 || files2.length !== 1) {
            toast(isPair ? `Add exactly one ${slots[0]} and one ${slots[1]}`
                         : 'Add exactly one English file and one Hindi file', 'error')
            setBusy(false); return
          }
          const form = new FormData()
          form.append('category_id', category)
          form.append('file', files[0])
          form.append('file2', files2[0])
          addPageFields(form)
          await api.submit(form)
          toast(isPair ? 'File pair queued ✓' : 'Bilingual pair queued ✓', 'ok')
        } else {
          for (const f of files) {
            const form = new FormData()
            form.append('category_id', category)
            form.append('file', f)
            addPageFields(form)
            await api.submit(form)
          }
          toast(`${files.length} file(s) queued ✓`, 'ok')
        }
        setFiles([]); setFiles2([])
      }
    } catch (e) {
      toast(e.message, 'error')
    } finally {
      setBusy(false)
    }
  }

  // ---- updates: checked once on open; the button appears only when the
  // Drive notice offers a newer version
  useEffect(() => {
    api.updateCheck().then(setUpdate).catch(() => {})
  }, [])

  const recheck = async () => {
    try {
      const u = await api.updateCheck()
      setUpdate(u)
      if (u.error) toast(u.error, 'error')
      else if (!u.configured) toast('Updates are not set up in this build', 'error')
      else if (!u.available) toast(`DocPilot ${u.current} is the latest version ✓`, 'ok')
    } catch { toast('Cannot reach the DocPilot backend', 'error') }
  }

  const installUpdate = async () => {
    try {
      await api.updatePost('/api/update/install')
      setUpd({ state: 'installing', progress: 100 })
    } catch (e) {
      setUpd(u => ({ ...u, state: 'ready' }))
      toast(e.message, 'error')
    }
  }

  const startUpdate = async () => {
    if (upd.state === 'ready') return installUpdate()
    const notes = update?.notes ? `\n\nWhat's new:\n${update.notes}` : ''
    if (!window.confirm(`Update DocPilot ${update.current} → ${update.latest}?${notes}\n\n` +
        'The new version is downloaded first. DocPilot then closes, installs it ' +
        'and opens again by itself (about a minute).')) return
    try {
      setUpd(await api.updatePost('/api/update/download'))
    } catch (e) { toast(e.message, 'error') }
  }

  // follow the download; install as soon as it is complete
  useEffect(() => {
    if (upd.state !== 'downloading') return
    const iv = setInterval(async () => {
      try {
        const s = await api.updateStatus()
        setUpd(s)
        if (s.state === 'error') toast(s.error || 'Update download failed', 'error')
        if (s.state === 'ready') installUpdate()
      } catch { /* backend busy */ }
    }, 1000)
    return () => clearInterval(iv)
  }, [upd.state])

  const active = useMemo(() => jobs.filter(j => j.status === 'running' || j.status === 'queued').length, [jobs])

  return (
    <div className="shell">
      {/* decorative background cubes */}
      <div className="bg-deco" aria-hidden="true">
        <div className="deco-cube c1"><Cube size={90} speed={26} /></div>
        <div className="deco-cube c2"><Cube size={54} speed={19} /></div>
      </div>

      <header className="topbar">
        <div className="brand">
          <Cube size={22} speed={8} />
          <div className="brand-text">
            <h1>DocPilot <span className="beta-tag">BETA</span></h1>
            <span>PW OnlyIAS · Document Studio</span>
          </div>
        </div>
        <div className="topbar-right">
          {active > 0 && <span className="active-pill">{active} in queue</span>}
          {update?.available && (
            <button className="update-btn" onClick={startUpdate}
                    disabled={upd.state === 'downloading' || upd.state === 'installing'}
                    title={update.notes || ''}>
              {upd.state === 'downloading' ? `Downloading ${upd.progress}%`
                : upd.state === 'ready' ? 'Install update'
                : `⬆ Update to ${update.latest}`}
            </button>
          )}
          {update && (
            <button className="version-tag" onClick={recheck} title="Check for updates">
              v{update.current}
            </button>
          )}
          <button className="theme-btn" onClick={() => setTheme(t => t === 'light' ? 'dark' : 'light')}
                  title="Toggle theme" aria-label="Toggle color theme">
            {theme === 'light' ? '🌙' : '☀️'}
          </button>
        </div>
      </header>

      {upd.state === 'installing' && (
        <div className="update-overlay" role="alert">
          <div className="update-card">
            <Cube size={34} speed={4} />
            <h2>Installing DocPilot {update?.latest}</h2>
            <p>DocPilot has closed to install the update and will open again by
               itself in about a minute. You can close this tab.</p>
          </div>
        </div>
      )}

      <main className="content">
        {/* ------------------------------------------------ new task card */}
        <section className="card new-task">
          <div className="card-head">
            <h2>New Task</h2>
            <p>Drop a raw typed document, pick its category, get the branded PDF.</p>
          </div>

          <label className="field-label" htmlFor="cat">Category</label>
          <div className="select-wrap">
            <select id="cat" value={category} onChange={e => setCategory(e.target.value)}>
              {cats.map(c => (
                <option key={c.id} value={c.id} disabled={!c.ready}>
                  {c.label}{c.ready ? '' : '  — template pending'}
                </option>
              ))}
            </select>
          </div>

          {cat?.hint && <p className="cat-hint">{cat.hint}</p>}

          {pageMode !== 'none' && (
            <div className="pagerow">
              <div className="pagefield">
                <label htmlFor="sp">Start page</label>
                <input id="sp" type="number" min="1" placeholder="e.g. 23"
                       value={startPage} onChange={e => setStartPage(e.target.value)} />
              </div>
              {pageMode === 'start_end' && (
                <div className="pagefield">
                  <label htmlFor="ep">End page <span className="opt">(optional)</span></label>
                  <input id="ep" type="number" min="1" placeholder="e.g. 22"
                         value={endPage} onChange={e => setEndPage(e.target.value)} />
                </div>
              )}
            </div>
          )}

          {hasCover && (
            <div className="coverbox">
              <div className="coverbox-head">
                <span>Front page</span>
                <small>{dateRequired
                  ? 'The date is required; leave the other blocks empty to take them from the document'
                  : 'Leave a block empty to take it from the document'}</small>
              </div>
              <div className="pagefield">
                <label htmlFor="cd">Date {dateRequired
                  ? <span className="req" title="required">*</span>
                  : <span className="opt">(optional)</span>}</label>
                <input id="cd" type="text"
                       placeholder={dateRequired
                         ? 'e.g. 17 August 2026  or  17 अगस्त 2026'
                         : 'e.g. 17th August 2026 - read from the file if empty'}
                       aria-required={dateRequired}
                       value={coverDate} onChange={e => setCoverDate(e.target.value)} />
              </div>
              {hasField('index') && (
                <div className="pagefield">
                  <label htmlFor="ci">{hasField('index2') ? 'Daily CA index'
                    : hasField('qr') ? 'Topics' : 'Index'} <span className="opt">(optional - one topic per line)</span></label>
                  <textarea id="ci" rows={4}
                            placeholder={(hasField('qr')
                              ? (cat?.variant === 'Hindi'
                                ? ["भारत-पश्चिम एशिया संबंध समस्या से संभावनाओं तक", "... (editorial topics from the file if empty; page numbers are added)"]
                                : ["India is losing West Asia, but it is not too Late", "... (editorial topics from the file if empty)"])
                              : (cat?.variant === 'Hindi'
                                ? ["राष्ट्रीय शहरीकरण परिषद (NUP)", "सूर्य ग्रहण, 2026", "... (topic headings from the file if empty; page numbers are added)"]
                                : ["Forest fires in India", "Mines and Minerals Amendment Bill, 2026", "... (topic headings from the file if empty)"])).join("\n")}
                            value={coverIndex} onChange={e => setCoverIndex(e.target.value)} />
                </div>
              )}
              {hasField('index2') && (
                <div className="pagefield">
                  <label htmlFor="ci2">Editorial index <span className="opt">(optional - one topic per line)</span></label>
                  <textarea id="ci2" rows={3}
                            placeholder={(cat?.variant === 'Hindi'
                              ? ["संसदीय अक्षमता", "भारत का E20 अभियान", "... (editorial topics from the file if empty; page numbers are added)"]
                              : ["India’s Democracy needs different Electoral Rhythms", "... (editorial topics from the file if empty)"]).join("\n")}
                            value={coverIndex2} onChange={e => setCoverIndex2(e.target.value)} />
                </div>
              )}
              {hasField('shorts') && (
                <div className="pagefield">
                  <label htmlFor="cs">News in Shorts <span className="opt">(optional - one item per line)</span></label>
                  <textarea id="cs" rows={4}
                            placeholder={(cat?.variant === 'Hindi'
                              ? ["AI द्वारा डिजाइन किया गया वायरस", "भारत-केन्या संबंध", "... (news-table topics from the file if empty)"]
                              : ["SURYAPATH TIRANGA", "Cyber Kushti 2026", "... (news-table topics from the file if empty)"]).join("\n")}
                            value={coverShorts} onChange={e => setCoverShorts(e.target.value)} />
                </div>
              )}
              {hasField('qr') && (
                <div className="pagefield">
                  <label htmlFor="cq">Video link <span className="opt">(optional - replaces the QR code in the file)</span></label>
                  <input id="cq" type="url"
                         placeholder="https://youtu.be/…  - the file's own QR is used if empty"
                         value={coverQr} onChange={e => setCoverQr(e.target.value)} />
                </div>
              )}
            </div>
          )}

          {testFields.length > 0 && (
            <div className="coverbox">
              <div className="coverbox-head">
                <span>Test details</span>
                <small>Leave a field empty to keep the value from the file name or the template</small>
              </div>
              {hasTest('series') && (
                <div className="pagefield">
                  <label htmlFor="ts-series">Test series name</label>
                  <input id="ts-series" type="text" placeholder="e.g. Prelims Test Series 1.0"
                         value={details.series || ''} onChange={setDetail('series')} />
                </div>
              )}
              {hasTest('type') && (
                <div className="pagefield">
                  <label htmlFor="ts-type">Test series type</label>
                  <div className="select-wrap ts-select">
                    <select id="ts-type" value={details.type || ''} onChange={setDetail('type')}>
                      <option value="">— as in the template (Full Length) —</option>
                      <option value="Sectional">Sectional</option>
                      <option value="Full Length">Full Length</option>
                    </select>
                  </div>
                </div>
              )}
              {hasTest('subject') && (
                <div className="pagefield">
                  <label htmlFor="ts-subject">Subject name</label>
                  <input id="ts-subject" type="text" placeholder="e.g. Indian Polity (Part II)"
                         value={details.subject || ''} onChange={setDetail('subject')} />
                </div>
              )}
              <div className="pagerow">
                {hasTest('testno') && (
                  <div className="pagefield">
                    <label htmlFor="ts-testno">Test number</label>
                    <input id="ts-testno" type="text" inputMode="numeric" maxLength={3} placeholder="e.g. 08"
                           value={details.testno || ''} onChange={setDigits('testno')} />
                  </div>
                )}
                {hasTest('code') && (
                  <div className="pagefield">
                    <label htmlFor="ts-code">Test code</label>
                    <input id="ts-code" type="text" placeholder="e.g. PTS 2708"
                           value={details.code || ''} onChange={setDetail('code')} />
                  </div>
                )}
                {hasTest('year') && (
                  <div className="pagefield">
                    <label htmlFor="ts-year">Target year</label>
                    <input id="ts-year" type="text" inputMode="numeric" maxLength={4} placeholder="e.g. 2027"
                           value={details.year || ''} onChange={setDigits('year')} />
                  </div>
                )}
              </div>
              {(hasTest('marks') || hasTest('permark') || hasTest('time')) && (
                <div className="pagerow">
                  {hasTest('marks') && (
                    <div className="pagefield">
                      <label htmlFor="ts-marks">Total marks</label>
                      <input id="ts-marks" type="text" inputMode="numeric" maxLength={4} placeholder="e.g. 200"
                             value={details.marks || ''} onChange={setDigits('marks')} />
                    </div>
                  )}
                  {hasTest('permark') && (
                    <div className="pagefield">
                      <label htmlFor="ts-permark">Marks per question</label>
                      <input id="ts-permark" type="text" inputMode="decimal" maxLength={4} placeholder="e.g. 2"
                             value={details.permark || ''} onChange={setDecimal('permark')} />
                    </div>
                  )}
                  {hasTest('time') && (
                    <div className="pagefield">
                      <label htmlFor="ts-time">Time <span className="opt">(minutes)</span></label>
                      <input id="ts-time" type="text" inputMode="numeric" maxLength={4} placeholder="e.g. 120"
                             value={details.time || ''} onChange={setDigits('time')} />
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          <div className="seg" role="tablist" aria-label="Input source">
            <button role="tab" aria-selected={mode === 'upload'}
                    className={mode === 'upload' ? 'seg-on' : ''}
                    onClick={() => setMode('upload')}>⬆ Local Upload</button>
            <button role="tab" aria-selected={mode === 'link'}
                    className={mode === 'link' ? 'seg-on' : ''}
                    onClick={() => setMode('link')}>🔗 Drive Link</button>
            <div className={`seg-glider ${mode === 'link' ? 'right' : ''}`} />
          </div>

          {mode === 'upload' ? (
            <>
              {needsTwo && <p className="slot-label">{slots[0]}</p>}
              <div className={`dropzone ${dragOver ? 'over' : ''}`}
                   onDragOver={e => { e.preventDefault(); setDragOver(true) }}
                   onDragLeave={() => setDragOver(false)}
                   onDrop={e => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files) }}
                   onClick={() => fileInput.current?.click()}>
                <div className="dz-icon">📄</div>
                <strong>{needsTwo ? (isPair ? `Drop the ${slots[0]} (.docx)` : 'Drop the English .docx') : 'Drag & drop .docx files'}</strong>
                <span>{needsTwo ? 'one file' : 'or click to browse — multiple files become separate queue items'}</span>
                <input ref={fileInput} type="file" accept=".docx" multiple={!needsTwo} hidden
                       onChange={e => { addFiles(e.target.files); e.target.value = '' }} />
              </div>
              {files.length > 0 && (
                <ul className="filelist">
                  {files.map((f, i) => (
                    <li key={i}>
                      <span className="fname" title={f.name}>{f.name}</span>
                      <span className="fsize">{(f.size / 1048576).toFixed(1)} MB</span>
                      <button className="x" aria-label={`Remove ${f.name}`}
                              onClick={() => setFiles(fs => fs.filter((_, k) => k !== i))}>✕</button>
                    </li>
                  ))}
                </ul>
              )}

              {needsTwo && (
                <>
                  <p className="slot-label">{slots[1]}</p>
                  <div className="dropzone"
                       onDragOver={e => e.preventDefault()}
                       onDrop={e => { e.preventDefault(); addFiles2(e.dataTransfer.files) }}
                       onClick={() => fileInput2.current?.click()}>
                    <div className="dz-icon">📄</div>
                    <strong>{isPair ? `Drop the ${slots[1]} (.docx)` : 'Drop the Hindi .docx'}</strong>
                    <span>one file</span>
                    <input ref={fileInput2} type="file" accept=".docx" hidden
                           onChange={e => { addFiles2(e.target.files); e.target.value = '' }} />
                  </div>
                  {files2.length > 0 && (
                    <ul className="filelist">
                      {files2.map((f, i) => (
                        <li key={i}>
                          <span className="fname" title={f.name}>{f.name}</span>
                          <span className="fsize">{(f.size / 1048576).toFixed(1)} MB</span>
                          <button className="x" aria-label={`Remove ${f.name}`}
                                  onClick={() => setFiles2(fs => fs.filter((_, k) => k !== i))}>✕</button>
                        </li>
                      ))}
                    </ul>
                  )}
                </>
              )}
            </>
          ) : (
            <div className="linkbox">
              {needsTwo && <p className="slot-label">{slots[0]}</p>}
              <input type="url" placeholder="https://docs.google.com/document/d/…  or  https://drive.google.com/file/d/…"
                     value={link} onChange={e => setLink(e.target.value)}
                     onKeyDown={e => e.key === 'Enter' && !needsTwo && submit()} />
              {needsTwo && (
                <>
                  <p className="slot-label">{slots[1]}</p>
                  <input type="url" placeholder={isPair ? `Google Drive / Docs link for the ${slots[1]}` : "Google Drive / Docs link for the Hindi paper"}
                         value={link2} onChange={e => setLink2(e.target.value)}
                         onKeyDown={e => e.key === 'Enter' && submit()} />
                </>
              )}
              <p className="hint">Share setting must be “Anyone with the link can view”. Native Google Docs are exported as .docx automatically.</p>
            </div>
          )}

          <button className="primary" disabled={busy} onClick={submit}>
            {busy ? 'Queuing…'
              : needsTwo ? (isPair ? 'Queue File Pair' : 'Queue Bilingual Pair')
              : mode === 'link' ? 'Queue Link'
              : files.length > 1 ? `Queue ${files.length} Files` : 'Queue File'}
          </button>
        </section>

        {/* ------------------------------------------------ queue */}
        <section className="queue">
          <div className="queue-head">
            <h2>Queue</h2>
            {jobs.some(j => ['done', 'error', 'stopped'].includes(j.status)) && (
              <button className="ghost" onClick={() => jobs.filter(j => ['done', 'error', 'stopped'].includes(j.status)).forEach(j => api.remove(j.id))}>
                Clear finished
              </button>
            )}
          </div>

          {jobs.length === 0 && (
            <div className="empty">
              <Cube size={64} speed={14} />
              <p>No tasks yet — queue your first document above.</p>
            </div>
          )}

          {jobs.map(j => {
            const meta = STATUS_META[j.status] || STATUS_META.queued
            const jcat = cats.find(c => c.id === j.category_id)
            return (
              <article className={`job ${j.status}`} key={j.id}>
                <div className="job-main">
                  <div className="job-title-row">
                    <span className={`chip ${meta.cls}`}>{meta.label}</span>
                    <h3 title={j.display_name}>{j.display_name}</h3>
                  </div>
                  <div className="job-sub">
                    <span>{jcat ? jcat.label : j.category_id}</span>
                    <span className="dot">·</span>
                    <span>{j.source_kind === 'link' ? 'Drive link' : 'Upload'}</span>
                    {j.display_name2 && <><span className="dot">·</span><span>+ {j.display_name2}</span></>}
                    {j.start_page && <><span className="dot">·</span>
                      <span>p{j.start_page}{j.end_page ? `–${j.end_page}` : '+'}</span></>}
                    {j.started && <><span className="dot">·</span><span>{fmtTime(j.started)}</span></>}
                  </div>
                  {(j.status === 'running' || j.status === 'queued') && (
                    <div className="prog">
                      <div className="prog-bar" style={{ width: `${j.progress}%` }} />
                    </div>
                  )}
                  <div className="job-stage">
                    {j.status === 'running' && <span className="mini-cube"><Cube size={11} speed={4} /></span>}
                    {j.status === 'error' ? <span className="err">{j.error}</span> : <span>{j.stage}</span>}
                  </div>
                  {j.status === 'done' && j.notes?.length > 0 && (
                    <ul className="job-notes">
                      {j.notes.map((n, k) => <li key={k}>{n}</li>)}
                    </ul>
                  )}
                </div>
                <div className="job-actions">
                  {(j.status === 'running' || j.status === 'queued') && (
                    <button className="danger" onClick={() => api.stop(j.id)}>⏹ Stop</button>
                  )}
                  {j.status === 'done' && (
                    <a className="btn-dl" href={`/api/jobs/${j.id}/download`} download>
                      ⬇ {(jcat?.output || 'pdf').toUpperCase()}
                    </a>
                  )}
                  {j.status === 'done' && j.docx_name && (
                    <a className="btn-dl secondary"
                       href={`/api/jobs/${j.id}/download?fmt=docx`} download
                       title="Editable Word file (page numbers and the promo page are added in the PDF)">
                      ⬇ WORD
                    </a>
                  )}
                  {j.status === 'done' && j.pdf_name && (
                    <a className="btn-dl secondary"
                       href={`/api/jobs/${j.id}/download?fmt=pdf`} download
                       title="Print-ready PDF rendered from the Word file">
                      ⬇ PDF
                    </a>
                  )}
                  {['done', 'error', 'stopped'].includes(j.status) && (
                    <button className="ghost" onClick={() => api.remove(j.id)}>Remove</button>
                  )}
                </div>
              </article>
            )
          })}
        </section>
      </main>

      <footer className="foot">
        <span>DocPilot · runs fully on this machine · Word-fidelity engine</span>
      </footer>

      <div className="toasts" aria-live="polite">
        {toasts.map(t => <div key={t.id} className={`toast ${t.kind}`}>{t.msg}</div>)}
      </div>
    </div>
  )
}
