import { useEffect, useMemo, useState } from 'react'

type FormData = {
  license_type: string; name: string; birth_date: string; phone: string; email: string;
  identity_number: string; exam_date: string; region: string; station: string; keep_browser: boolean;
}
type Options = { license_types: string[]; stations: Record<string, string[]> }
type Job = { state: string; message: string }

const today = new Date().toLocaleDateString('en-CA')
const defaultForm: FormData = {
  license_type: '', name: '', birth_date: '', phone: '', email: '', identity_number: '',
  exam_date: today, region: '', station: '', keep_browser: true,
}

async function api<T>(url: string, init?: RequestInit): Promise<T> {
  const response = await fetch(url, { ...init, headers: { 'Content-Type': 'application/json', ...init?.headers } })
  if (!response.ok) {
    const body = await response.json().catch(() => ({}))
    const detail = Array.isArray(body.detail)
      ? body.detail.map((item: { msg?: string }) => item.msg ?? '資料格式不正確').join('、')
      : body.detail
    throw new Error(detail || `操作失敗（${response.status}）`)
  }
  if (response.status === 204) return undefined as T
  return response.json()
}

export default function App() {
  const [options, setOptions] = useState<Options>({ license_types: [], stations: {} })
  const [form, setForm] = useState<FormData>(defaultForm)
  const [job, setJob] = useState<Job>({ state: 'idle', message: '正在連接本機服務…' })
  const [notice, setNotice] = useState('')
  const [showIdentity, setShowIdentity] = useState(false)
  const running = job.state === 'running' || job.state === 'stopping'
  const stationOptions = useMemo(() => options.stations[form.region] ?? [], [options, form.region])

  useEffect(() => {
    Promise.all([api<Options>('/api/options'), api<FormData | null>('/api/profile'), api<Job>('/api/registration/status')])
      .then(([nextOptions, profile, status]) => {
        setOptions(nextOptions)
        setForm(profile ?? {
          ...defaultForm,
          license_type: nextOptions.license_types[0] ?? '',
          region: Object.keys(nextOptions.stations)[0] ?? '',
          station: Object.values(nextOptions.stations)[0]?.[0] ?? '',
        })
        setJob(status)
      })
      .catch((error) => setJob({ state: 'error', message: `無法連接後端：${error.message}` }))
  }, [])

  useEffect(() => {
    if (!running) return
    const timer = window.setInterval(() => {
      api<Job>('/api/registration/status').then(setJob).catch((error) => setJob({ state: 'error', message: error.message }))
    }, 1000)
    return () => window.clearInterval(timer)
  }, [running])

  const update = (key: keyof FormData, value: string | boolean) => {
    setForm((current) => ({ ...current, [key]: value }))
    setNotice('')
  }

  const changeRegion = (region: string) => {
    setForm((current) => ({ ...current, region, station: options.stations[region]?.[0] ?? '' }))
  }

  const submit = async (event: { preventDefault(): void }, mode: 'save' | 'start') => {
    event.preventDefault()
    setNotice('')
    try {
      if (mode === 'save') {
        await api<void>('/api/profile', { method: 'PUT', body: JSON.stringify(form) })
        setNotice('資料已加密儲存在這台電腦。')
      } else {
        setJob(await api<Job>('/api/registration/start', { method: 'POST', body: JSON.stringify(form) }))
        setNotice('')
      }
    } catch (error) {
      setNotice(error instanceof Error ? error.message : '操作失敗')
    }
  }

  const stop = async () => {
    try { setJob(await api<Job>('/api/registration/stop', { method: 'POST' })) }
    catch (error) { setNotice(error instanceof Error ? error.message : '停止失敗') }
  }

  return (
    <main>
      <header className="hero">
        <div className="eyebrow">LOCAL REGISTRATION ASSISTANT</div>
        <h1>駕照報名，<span>少一點手忙腳亂。</span></h1>
        <p>資料只留在這台電腦。填好一次，讓瀏覽器替你完成報名流程。</p>
        <div className="privacy-note"><span className="lock">●</span> 本機加密保存 · 不建立帳號 · 不上傳雲端</div>
      </header>

      <form className="card" onSubmit={(event) => submit(event, 'start')}>
        <section>
          <div className="section-title"><span>01</span><div><h2>報考資訊</h2><p>選擇考照類型、日期與地點</p></div></div>
          <div className="grid two">
            <label>駕照類型<select value={form.license_type} onChange={(e) => update('license_type', e.target.value)} required>{options.license_types.map((item) => <option key={item}>{item}</option>)}</select></label>
            <label>考試日期<input type="date" min={today} value={form.exam_date} onChange={(e) => update('exam_date', e.target.value)} required /></label>
            <label>監理所區域<select value={form.region} onChange={(e) => changeRegion(e.target.value)} required>{Object.keys(options.stations).map((item) => <option key={item}>{item}</option>)}</select></label>
            <label>監理所<select value={form.station} onChange={(e) => update('station', e.target.value)} required>{stationOptions.map((item) => <option key={item}>{item}</option>)}</select></label>
          </div>
        </section>

        <section>
          <div className="section-title"><span>02</span><div><h2>個人資料</h2><p>用於監理服務網報名表單</p></div></div>
          <div className="grid two">
            <label>姓名<input value={form.name} onChange={(e) => update('name', e.target.value)} autoComplete="name" required /></label>
            <label>出生日期<input type="date" max={today} value={form.birth_date} onChange={(e) => update('birth_date', e.target.value)} required /></label>
            <label>電話<input inputMode="tel" value={form.phone} onChange={(e) => update('phone', e.target.value)} placeholder="0912345678" autoComplete="tel" required /></label>
            <label>電子郵件<input type="email" value={form.email} onChange={(e) => update('email', e.target.value)} placeholder="name@example.com" autoComplete="email" required /></label>
            <label className="wide">身分證字號<div className="password"><input type={showIdentity ? 'text' : 'password'} value={form.identity_number} onChange={(e) => update('identity_number', e.target.value.toUpperCase())} maxLength={10} autoComplete="off" required /><button type="button" onClick={() => setShowIdentity((value) => !value)}>{showIdentity ? '隱藏' : '顯示'}</button></div></label>
          </div>
        </section>

        <section className="finish">
          <label className="check"><input type="checkbox" checked={form.keep_browser} onChange={(e) => update('keep_browser', e.target.checked)} /><span>完成後保留瀏覽器，方便確認報名結果</span></label>
          <div className={`status ${job.state}`}><span className="status-dot" /><div><small>目前狀態</small><strong>{job.message}</strong></div></div>
          {notice && <div className="notice" role="alert">{notice}</div>}
          <div className="actions">
            <button className="secondary" type="button" disabled={running} onClick={(event) => submit(event, 'save')}>儲存資料</button>
            {running ? <button className="danger" type="button" onClick={stop}>停止作業</button> : <button className="primary" type="submit">開始報名 <span>→</span></button>}
          </div>
        </section>
      </form>
      <footer>此工具只在本機運作。送出前請再次確認報考資訊。</footer>
    </main>
  )
}
