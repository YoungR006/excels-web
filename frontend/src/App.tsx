import { useEffect, useMemo, useRef, useState, useLayoutEffect } from 'react'
import {
  Bot,
  User,
  Send,
  Paperclip,
  FileSpreadsheet,
  Play,
  RotateCcw,
  Download,
  FilePlus,
  Layers,
  Code2,
  Table as TableIcon,
  PanelRightClose,
  PanelRightOpen,
  Sparkles,
} from 'lucide-react'

// Styles (inline to match your current structure)
const styles = `
:root {
  --bg-app: #f9fafb;
  --bg-sidebar: #f3f4f6;
  --bg-card: #ffffff;
  --bg-hover: #f3f4f6;
  --border: #e5e7eb;
  --text-primary: #111827;
  --text-secondary: #6b7280;
  --text-accent: #2563eb;
  --accent: #111827;
  --accent-hover: #000000;
  --danger: #ef4444;
  --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
  --shadow-md: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
  --shadow-lg: 0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1);
  --radius-md: 0.5rem;
  --radius-lg: 0.75rem;
  --radius-xl: 1rem;
}

body {
  margin: 0;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  color: var(--text-primary);
  background: var(--bg-app);
  height: 100vh;
  overflow: hidden;
}

.app-container {
  display: flex;
  height: 100vh;
  width: 100vw;
  overflow: hidden;
}

.sidebar {
  width: 260px;
  background: var(--bg-sidebar);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  overflow: hidden;
}

.sidebar-header {
  padding: 1.25rem;
  display: flex;
  align-items: center;
  gap: 0.75rem;
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
}

.sidebar-header h1 {
  font-size: 1rem;
  font-weight: 600;
  margin: 0;
  color: var(--text-primary);
}

.sidebar-content {
  flex: 1;
  overflow-y: auto;
  padding: 1rem;
  display: flex;
  flex-direction: column;
  gap: 1.5rem;
  min-width: 260px;
}

.section-title {
  font-size: 0.75rem;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--text-secondary);
  font-weight: 600;
  margin-bottom: 0.75rem;
}

.file-actions {
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
}

.history-list {
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
}

.history-item {
  padding: 0.75rem;
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  font-size: 0.85rem;
  display: flex;
  flex-direction: column;
  gap: 0.4rem;
  transition: all 0.2s;
  position: relative;
}

.history-item:hover {
  border-color: var(--text-secondary);
}

.history-meta {
  display: flex;
  justify-content: space-between;
  align-items: center;
  color: var(--text-secondary);
  font-size: 0.75rem;
}

.main-chat {
  flex: 1;
  display: flex;
  flex-direction: column;
  position: relative;
  background: var(--bg-app);
  min-width: 0;
}

.chat-scroll-area {
  flex: 1;
  overflow-y: auto;
  padding: 2rem;
  display: flex;
  flex-direction: column;
  gap: 1.5rem;
  scroll-behavior: smooth;
  padding-bottom: 2rem;
}

.welcome-hero {
  text-align: center;
  margin-top: 10vh;
  color: var(--text-primary);
  animation: fadeIn 0.5s ease-out;
}

.welcome-hero h2 {
  font-size: 2rem;
  font-weight: 600;
  margin-bottom: 1rem;
}

.message-bubble {
  display: flex;
  gap: 1rem;
  max-width: 800px;
  margin: 0 auto;
  width: 100%;
  animation: slideIn 0.3s ease-out;
}

@keyframes fadeIn {
  from { opacity: 0; }
  to { opacity: 1; }
}

@keyframes slideIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}

.avatar {
  width: 2rem;
  height: 2rem;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  font-size: 0.875rem;
}

.avatar.user {
  background: var(--text-primary);
  color: white;
}

.avatar.assistant {
  background: var(--accent);
  color: white;
}

.avatar.system {
  background: var(--text-secondary);
  color: white;
}

.message-content {
  flex: 1;
  padding-top: 0.25rem;
  line-height: 1.6;
  font-size: 0.95rem;
  white-space: pre-wrap;
}

.message-content.user {
  font-weight: 500;
  text-align: right;
}

.input-container {
  padding: 1.5rem 2rem 2rem;
  background: linear-gradient(to top, var(--bg-app) 80%, transparent);
  max-width: 800px;
  margin: 0 auto;
  width: 100%;
  box-sizing: border-box;
  z-index: 10;
}

.input-box {
  background: var(--bg-card);
  border: 1px solid var(--border);
  box-shadow: var(--shadow-lg);
  border-radius: var(--radius-xl);
  padding: 0.75rem;
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
  transition: border-color 0.2s;
}

.input-box:focus-within {
  border-color: var(--text-secondary);
}

textarea.chat-textarea {
  border: none;
  resize: none;
  outline: none;
  width: 100%;
  font-family: inherit;
  padding: 0.5rem;
  font-size: 1rem;
  max-height: 200px;
  background: transparent;
  color: var(--text-primary);
}

.input-actions {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 0 0.5rem;
}

.helper-text {
  font-size: 0.75rem;
  color: var(--text-secondary);
}

.artifact-panel {
  width: 45%;
  background: var(--bg-card);
  border-left: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  transition: width 0.3s ease;
  min-width: 300px;
}

.artifact-header {
  padding: 0.75rem 1.25rem;
  border-bottom: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: center;
  background: #fcfcfd;
}

.artifact-tabs {
  display: flex;
  gap: 0.5rem;
  background: var(--bg-sidebar);
  padding: 0.25rem;
  border-radius: var(--radius-md);
}

.tab-btn {
  padding: 0.4rem 0.8rem;
  font-size: 0.8rem;
  border-radius: 0.35rem;
  border: none;
  cursor: pointer;
  background: transparent;
  color: var(--text-secondary);
  font-weight: 500;
}

.tab-btn.active {
  background: var(--bg-card);
  color: var(--text-primary);
  box-shadow: var(--shadow-sm);
}

.artifact-content {
  flex: 1;
  overflow: hidden;
  position: relative;
  display: flex;
  flex-direction: column;
}

.table-container {
  overflow: auto;
  flex: 1;
  width: 100%;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.85rem;
}

th {
  background: var(--bg-sidebar);
  position: sticky;
  top: 0;
  z-index: 10;
  font-weight: 600;
  text-align: left;
  color: var(--text-secondary);
  white-space: nowrap;
}

th, td {
  padding: 0.6rem 1rem;
  border-bottom: 1px solid var(--border);
  border-right: 1px solid var(--border);
  white-space: nowrap;
}

td {
  color: var(--text-primary);
}

.script-editor {
  flex: 1;
  padding: 1rem;
  font-family: 'Menlo', 'Monaco', monospace;
  font-size: 0.9rem;
  border: none;
  resize: none;
  outline: none;
  background: #fafafa;
  color: var(--text-primary);
  line-height: 1.5;
}

.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 0.5rem;
  padding: 0.5rem 0.75rem;
  border-radius: var(--radius-md);
  font-size: 0.875rem;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
  border: 1px solid transparent;
}

.btn-primary {
  background: var(--accent);
  color: white;
}

.btn-primary:hover {
  background: var(--accent-hover);
}

.btn-secondary {
  background: white;
  border: 1px solid var(--border);
  color: var(--text-primary);
}

.btn-secondary:hover {
  background: var(--bg-hover);
  border-color: var(--text-secondary);
}

.btn-icon {
  padding: 0.4rem;
  color: var(--text-secondary);
  border-radius: var(--radius-md);
  background: transparent;
  border: none;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
}

.btn-icon:hover:not(:disabled) {
  background: var(--bg-sidebar);
  color: var(--text-primary);
}

.btn-icon:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.status-badge {
  font-size: 0.75rem;
  padding: 0.2rem 0.5rem;
  border-radius: 99px;
  background: #fff1f2;
  color: var(--danger);
}

.cell-warn {
  background: #fee2e2;
  color: #991b1b;
  font-weight: 600;
}

.empty-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100%;
  color: var(--text-secondary);
  gap: 1rem;
}

@media (max-width: 1024px) {
  .app-container {
    flex-direction: column;
  }

  .sidebar {
    display: none;
  }

  .artifact-panel {
    width: 100%;
    height: 40vh;
    border-left: none;
    border-top: 1px solid var(--border);
  }
}
`

type PreviewResponse = {
  sheet: string
  columns: string[]
  rows: Record<string, unknown>[]
  rules?: { type: string; sheet: string; column: string; threshold: number; color: string }[]
  row_count: number
  col_count: number
}

type AnalysisResult = {
  type: 't_test'
  sheet: string
  column_a: string
  column_b: string
  n_a: number
  n_b: number
  mean_a: number
  mean_b: number
  t_stat: number
  p_value: number
  df: number
}

type CommitSummary = {
  id: string
  message: string
  timestamp: string
  changed_sheets: string[]
}

const API_BASE = 'http://localhost:8000'

function App() {
  useLayoutEffect(() => {
    const styleTag = document.createElement('style')
    styleTag.textContent = styles
    document.head.appendChild(styleTag)
    return () => {
      document.head.removeChild(styleTag)
    }
  }, [])

  const [sessionId, setSessionId] = useState<string | null>(null)
  const [workbook, setWorkbook] = useState<string | null>(null)
  const [sheet, setSheet] = useState<string>('')

  const [preview, setPreview] = useState<PreviewResponse | null>(null)
  const [history, setHistory] = useState<CommitSummary[]>([])

  const [opsText, setOpsText] = useState<string>('')
  const [status, setStatus] = useState<string>('')
  const [nlpText, setNlpText] = useState('')
  const [messages, setMessages] = useState<
    { role: 'user' | 'assistant' | 'system'; content: string; isError?: boolean }[]
  >([{ role: 'assistant', content: '\u4f60\u597d\uff01\u6211\u662f\u4f60\u7684 Excel \u52a9\u624b\u3002\u4e0a\u4f20\u6587\u4ef6\u540e\uff0c\u544a\u8bc9\u6211\u4f60\u60f3\u5982\u4f55\u4fee\u6539\u3002' }])

  const [isParsing, setIsParsing] = useState(false)
  const [hasParsedScript, setHasParsedScript] = useState(false)
  const [activeTab, setActiveTab] = useState<'preview' | 'code'>('preview')
  const [isSidebarOpen, setIsSidebarOpen] = useState(true)

  const [emptyName] = useState('new_table.xlsx')
  const [batchFiles, setBatchFiles] = useState<FileList | null>(null)
  const [analysisOutput, setAnalysisOutput] = useState<'sheet' | 'column'>('sheet')
  const [analysisPrefix, setAnalysisPrefix] = useState('t_test')

  const fileInputRef = useRef<HTMLInputElement>(null)
  const messagesEndRef = useRef<HTMLDivElement>(null)

  const canOperate = useMemo(() => !!sessionId && !!workbook, [sessionId, workbook])

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  useEffect(() => {
    const createSession = async () => {
      try {
        const res = await fetch(`${API_BASE}/sessions`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: 'default' }),
        })
        const data = await res.json()
        setSessionId(data.id)
      } catch {
        setStatus('\u65e0\u6cd5\u8fde\u63a5\u540e\u7aef\u670d\u52a1')
      }
    }
    createSession()
  }, [])

  const refreshPreview = async (targetSheet?: string) => {
    if (!sessionId || !workbook) return
    const encodedWorkbook = encodeURIComponent(workbook)
    try {
      const res = await fetch(`${API_BASE}/sessions/${sessionId}/workbooks/${encodedWorkbook}/preview`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheet: targetSheet || sheet || undefined, limit: 100 }),
      })
      const data: PreviewResponse = await res.json()
      setPreview(data)
      setSheet(data.sheet)
    } catch (e) {
      console.error(e)
    }
  }

  const shouldHighlight = (col: string, value: unknown) => {
    if (!preview?.rules || !preview.rules.length) return false
    const rule = preview.rules.find((r) => r.column === col && r.type === 'lt')
    if (!rule) return false
    const num = Number(value)
    if (Number.isNaN(num)) return false
    return num < rule.threshold
  }

  const refreshHistory = async () => {
    if (!sessionId || !workbook) return
    const encodedWorkbook = encodeURIComponent(workbook)
    const res = await fetch(`${API_BASE}/sessions/${sessionId}/workbooks/${encodedWorkbook}/history`, {
      method: 'POST',
    })
    const data = await res.json()
    setHistory(data.commits ?? [])
  }

  const handleUpload = async (file: File) => {
    if (!sessionId) return
    setStatus('')
    const form = new FormData()
    form.append('file', file)

    setMessages((prev) => [...prev, { role: 'user', content: `\u4e0a\u4f20\u6587\u4ef6\uff1a${file.name}` }])

    try {
      const res = await fetch(`${API_BASE}/sessions/${sessionId}/workbooks/upload`, {
        method: 'POST',
        body: form,
      })
      const data = await res.json()

      setWorkbook(data.filename)
      setSheet(data.sheets?.[0] ?? '')
      await refreshPreview(data.sheets?.[0])
      await refreshHistory()

      setMessages((prev) => [
        ...prev,
        { role: 'system', content: `\u2705 \u6587\u4ef6\u5df2\u52a0\u8f7d\uff1a${data.filename}` },
      ])
      setHasParsedScript(false)
    } catch {
      setMessages((prev) => [
        ...prev,
        { role: 'assistant', content: '\u4e0a\u4f20\u5931\u8d25', isError: true },
      ])
    }
  }

  const handleCreateEmpty = async () => {
    if (!sessionId) return
    try {
      const res = await fetch(`${API_BASE}/sessions/${sessionId}/workbooks/empty`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: emptyName, sheet_name: 'Sheet1' }),
      })
      const data = await res.json()
      setWorkbook(data.filename)
      setSheet(data.sheets?.[0] ?? '')
      await refreshPreview(data.sheets?.[0])
      await refreshHistory()
      setMessages((prev) => [
        ...prev,
        { role: 'system', content: `\u2705 \u5df2\u521b\u5efa\u7a7a\u767d\u8868\uff1a${data.filename}` },
      ])
    } catch {
      setStatus('\u521b\u5efa\u5931\u8d25')
    }
  }

  const handleParse = async () => {
    if (!canOperate) {
      setStatus('\u8bf7\u5148\u4e0a\u4f20\u6216\u521b\u5efa\u8868\u683c')
      return
    }
    if (!nlpText.trim()) return
    setStatus('')
    setIsParsing(true)

    const userMsg = nlpText
    setNlpText('')
    setMessages((prev) => [...prev, { role: "user", content: userMsg }])

    try {
      const res = await fetch(`${API_BASE}/nlp/parse`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: userMsg, sheet: sheet || undefined }),
      })

      if (!res.ok) {
        setStatus('\u89e3\u6790\u5931\u8d25\uff0c\u8bf7\u68c0\u67e5 API Key \u6216\u8f93\u5165\u5185\u5bb9')
        return
      }

      const data = await res.json()
      const patched = {
        ...data,
        operations: (data.operations || []).map((op: any) =>
          op.type === 't_test'
            ? { ...op, output: analysisOutput, output_prefix: analysisPrefix }
            : op
        ),
      }
      setOpsText(JSON.stringify(patched, null, 2))
      setHasParsedScript(true)
      setActiveTab('code')
      setMessages((prev) => [...prev, { role: "assistant", content: "\u6211\u5df2\u751f\u6210\u4fee\u6539\u811a\u672c\uff0c\u8bf7\u5728\u53f3\u4fa7\u786e\u8ba4\u5e76\u6267\u884c\u3002" }])
    } catch {
      setMessages((prev) => [...prev, { role: "assistant", content: "\u89e3\u6790\u5931\u8d25\uff0c\u8bf7\u91cd\u8bd5\u6216\u68c0\u67e5\u540e\u7aef\u8fde\u63a5\u3002", isError: true }])
    } finally {
      setIsParsing(false)
    }
  }

  const handleApplyOps = async () => {
    if (!sessionId || !workbook) return
    setStatus('')
    let payload
    try {
      payload = JSON.parse(opsText)
    } catch {
      setStatus('JSON \u683c\u5f0f\u9519\u8bef')
      return
    }

    const encodedWorkbook = encodeURIComponent(workbook)
    const res = await fetch(`${API_BASE}/sessions/${sessionId}/workbooks/${encodedWorkbook}/operations`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })

    if (!res.ok) {
      setStatus('\u6267\u884c\u5931\u8d25')
      return
    }

    const data = await res.json()
    await refreshPreview()
    await refreshHistory()
    setMessages((prev) => {
      const next = [...prev, { role: 'system', content: '\u2705 \u811a\u672c\u6267\u884c\u6210\u529f\uff0c\u9884\u89c8\u5df2\u5237\u65b0\u3002' }]
      if (data?.analysis?.length) {
        const results = data.analysis as AnalysisResult[]
        results.forEach((item) => {
          next.push({
            role: 'assistant',
            content:
              `T\u68c0\u9a8c\u7ed3\u679c\uff08${item.column_a} vs ${item.column_b}\uff09` +
              `\n\u6837\u672c\u91cf: ${item.n_a} / ${item.n_b}` +
              `\n\u5747\u503c: ${item.mean_a.toFixed(4)} / ${item.mean_b.toFixed(4)}` +
              `\nt = ${item.t_stat.toFixed(4)}, df = ${item.df.toFixed(2)}, p = ${item.p_value.toFixed(6)}`
          })
        })
      }
      return next
    })
    setHasParsedScript(false)
    setActiveTab('preview')
    setOpsText('')
  }

  const handleRollback = async (commitId: string) => {
    if (!sessionId || !workbook) return
    const encodedWorkbook = encodeURIComponent(workbook)
    await fetch(`${API_BASE}/sessions/${sessionId}/workbooks/${encodedWorkbook}/rollback`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ commit_id: commitId }),
    })
    await refreshPreview()
    await refreshHistory()
    setMessages((prev) => [
      ...prev,
      { role: 'system', content: `\u5df2\u56de\u6eda\u5230\u7248\u672c ${commitId.slice(0, 6)}` },
    ])
  }

  const handleExport = (format: 'xlsx' | 'csv') => {
    if (!sessionId || !workbook) return
    const encodedWorkbook = encodeURIComponent(workbook)
    window.open(`${API_BASE}/sessions/${sessionId}/workbooks/${encodedWorkbook}/export?format=${format}`, '_blank')
  }

  const handleBatch = async () => {
    if (!sessionId || !batchFiles) return
    const form = new FormData()
    for (const file of Array.from(batchFiles)) {
      form.append('files', file)
    }
    form.append('payload', opsText)

    setMessages((prev) => [
      ...prev,
      { role: 'user', content: `\u5df2\u63d0\u4ea4 ${batchFiles.length} \u4e2a\u6587\u4ef6\u8fdb\u884c\u6279\u91cf\u5904\u7406...` },
    ])

    const res = await fetch(`${API_BASE}/sessions/${sessionId}/batch`, {
      method: 'POST',
      body: form,
    })

    if (res.ok) {
      const data = await res.json()
      setMessages((prev) => [
        ...prev,
        { role: 'assistant', content: `\u6279\u91cf\u5904\u7406\u5b8c\u6210\uff1a${data.results?.length ?? 0} \u4e2a\u6587\u4ef6` },
      ])
    } else {
      setMessages((prev) => [
        ...prev,
        { role: 'assistant', content: '\u6279\u91cf\u5904\u7406\u5931\u8d25\u3002', isError: true },
      ])
    }
  }

  return (
    <div className="app-container">
      <aside
        className="sidebar"
        style={{
          width: isSidebarOpen ? '260px' : '0',
          opacity: isSidebarOpen ? 1 : 0,
          pointerEvents: isSidebarOpen ? 'auto' : 'none',
        }}
      >
        <div className="sidebar-header">
          <Sparkles size={18} color="#2563eb" />
          <h1 style={{ marginLeft: 8 }}>Excels Canvas</h1>
        </div>

        <div className="sidebar-content">
          <div>
            <div className="section-title">{'\u5f53\u524d\u6587\u4ef6'}</div>
            <div className="file-actions">
              {workbook ? (
                <div className="history-item" style={{ background: '#eff6ff', borderColor: '#bfdbfe' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <FileSpreadsheet size={16} color="#3b82f6" />
                    <span
                      style={{
                        fontWeight: 500,
                        color: '#1e40af',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                      }}
                    >
                      {workbook}
                    </span>
                  </div>
                </div>
              ) : (
                <div style={{ padding: '0 0.5rem', color: '#9ca3af' }}>{'\u5c1a\u672a\u9009\u62e9\u6587\u4ef6'}</div>
              )}
            </div>
          </div>

          <div>
            <div className="section-title">{'\u6587\u4ef6\u64cd\u4f5c'}</div>
            <div className="file-actions">
              <button className="btn btn-secondary" onClick={() => fileInputRef.current?.click()}>
                <Paperclip size={14} /> {'\u4e0a\u4f20 Excel'}
              </button>
              <input
                type="file"
                hidden
                ref={fileInputRef}
                accept=".xlsx,.xls,.csv"
                onChange={(e) => e.target.files?.[0] && handleUpload(e.target.files[0])}
              />

              <button className="btn btn-secondary" onClick={handleCreateEmpty}>
                <FilePlus size={14} /> {'\u521b\u5efa\u7a7a\u767d'}
              </button>

              <div style={{ display: 'flex', gap: 8 }}>
                <button className="btn btn-secondary" style={{ flex: 1 }} onClick={() => handleExport('xlsx')} disabled={!canOperate}>
                  <Download size={14} /> .xlsx
                </button>
                <button className="btn btn-secondary" style={{ flex: 1 }} onClick={() => handleExport('csv')} disabled={!canOperate}>
                  <Download size={14} /> .csv
                </button>
              </div>
            </div>
          </div>

          <div>
            <div className="section-title">{'\u6279\u91cf\u5904\u7406'}</div>
            <div className="file-actions">
              <input
                type="file"
                multiple
                style={{ fontSize: '0.75rem', color: '#6b7280' }}
                onChange={(e) => setBatchFiles(e.target.files)}
              />
              <button className="btn btn-secondary" style={{ fontSize: '0.75rem' }} onClick={handleBatch} disabled={!batchFiles || !opsText}>
                <Layers size={14} /> {'\u6267\u884c\u6279\u91cf'}
              </button>
            </div>
          </div>

          <div>
            <div className="section-title">{'\u7edf\u8ba1\u8f93\u51fa'}</div>
            <div className="file-actions">
              <label style={{ fontSize: '0.75rem', color: '#6b7280' }}>{'T \u68c0\u9a8c\u7ed3\u679c\u8f93\u51fa\u5230'}</label>
              <select
                value={analysisOutput}
                onChange={(e) => setAnalysisOutput(e.target.value as 'sheet' | 'column')}
                className="btn btn-secondary"
                style={{ justifyContent: 'space-between' }}
              >
                <option value="sheet">{'\u65b0\u5efa Sheet'}</option>
                <option value="column">{'\u65b0\u589e\u5217'}</option>
              </select>
              <input
                type="text"
                placeholder="\u7ed3\u679c\u5217\u524d\u7f00\uff08\u5982 t_test\uff09"
                value={analysisPrefix}
                onChange={(e) => setAnalysisPrefix(e.target.value)}
                className="btn btn-secondary"
                disabled={analysisOutput !== 'column'}
                style={{ textAlign: 'left', opacity: analysisOutput === 'column' ? 1 : 0.6 }}
              />
            </div>
          </div>

          {history.length > 0 && (
            <div>
              <div className="section-title">{'\u7248\u672c\u5386\u53f2'}</div>
              <div className="history-list">
                {history.map((commit) => (
                  <div key={commit.id} className="history-item">
                    <div className="history-meta">
                      <strong style={{ color: '#374151' }}>{commit.message}</strong>
                      <button
                        className="btn-icon"
                        style={{ padding: 0 }}
                        title="\u56de\u6eda\u5230\u6b64\u7248\u672c"
                        onClick={() => handleRollback(commit.id)}
                      >
                        <RotateCcw size={12} color="#6b7280" />
                      </button>
                    </div>
                    <div className="history-meta" style={{ fontSize: '0.7rem', marginTop: 4 }}>
                      <span>{commit.timestamp.split('T')[1]?.split('.')[0]}</span>
                      <code style={{ background: '#f3f4f6', padding: '0 4px', borderRadius: 4 }}>
                        {commit.id.slice(0, 6)}
                      </code>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </aside>

      <main className="main-chat">
        {!isSidebarOpen && (
          <button
            className="btn-icon"
            style={{
              position: 'absolute',
              top: 16,
              left: 16,
              zIndex: 20,
              background: 'white',
              boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
              border: '1px solid #e5e7eb',
            }}
            onClick={() => setIsSidebarOpen(true)}
          >
            <PanelRightOpen size={20} />
          </button>
        )}
        {isSidebarOpen && (
          <button
            className="btn-icon"
            style={{ position: 'absolute', top: 16, left: 16, zIndex: 20, color: '#9ca3af' }}
            onClick={() => setIsSidebarOpen(false)}
          >
            <PanelRightClose size={20} />
          </button>
        )}

        <div className="chat-scroll-area">
          {messages.length === 0 && (
            <div className="welcome-hero">
              <h2>Excel AI Canvas</h2>
              <p style={{ color: '#6b7280' }}>\u4e0a\u4f20\u8868\u683c\uff0c\u50cf\u804a\u5929\u4e00\u6837\u5904\u7406\u6570\u636e\u3002</p>
            </div>
          )}

          {messages.map((msg, idx) => (
            <div
              key={idx}
              className={`message-bubble ${msg.role === 'user' ? 'justify-end' : ''}`}
              style={msg.role === 'user' ? { justifyContent: 'flex-end' } : {}}
            >
              {msg.role !== 'user' && (
                <div className={`avatar ${msg.role}`}>
                  {msg.role === 'assistant' ? <Bot size={18} /> : <Sparkles size={16} />}
                </div>
              )}

              <div className={`message-content ${msg.role}`} style={msg.isError ? { color: '#ef4444' } : {}}>
                {msg.content}
              </div>

              {msg.role === 'user' && (
                <div className="avatar user">
                  <User size={18} />
                </div>
              )}
            </div>
          ))}
          <div ref={messagesEndRef} />
        </div>

        <div className="input-container">
          {status && (
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 8 }}>
              <span className="status-badge">{status}</span>
            </div>
          )}
          <div className="input-box">
            <textarea
              className="chat-textarea"
              placeholder={canOperate ? '\u63cf\u8ff0\u4f60\u60f3\u5982\u4f55\u4fee\u6539\u8868\u683c...' : '\u8bf7\u5148\u4e0a\u4f20 Excel \u6587\u4ef6...'}
              rows={1}
              value={nlpText}
              onChange={(e) => setNlpText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault()
                  handleParse()
                }
              }}
              disabled={!canOperate || isParsing}
            />
            <div className="input-actions">
              <button className="btn-icon" onClick={() => fileInputRef.current?.click()}>
                <Paperclip size={18} />
              </button>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span className="helper-text">{isParsing ? '\u601d\u8003\u4e2d...' : 'Enter \u53d1\u9001'}</span>
                <button
                  className="btn-icon"
                  style={nlpText.trim() ? { background: '#111827', color: 'white' } : {}}
                  onClick={handleParse}
                  disabled={!nlpText.trim() || isParsing}
                >
                  <Send size={16} />
                </button>
              </div>
            </div>
          </div>
        </div>
      </main>

      <aside className="artifact-panel">
        <header className="artifact-header">
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontWeight: 500, fontSize: '0.875rem', color: '#374151' }}>
            {activeTab === 'preview' ? <TableIcon size={16} /> : <Code2 size={16} />}
            {activeTab === 'preview' ? '\u6570\u636e\u9884\u89c8' : '\u751f\u6210\u811a\u672c'}
          </div>
          <div className="artifact-tabs">
            <button
              className="btn-icon"
              title="\u5237\u65b0\u9884\u89c8"
              onClick={() => refreshPreview()}
              disabled={!canOperate}
              style={{ marginRight: 6 }}
            >
              <RotateCcw size={14} />
            </button>
            <button className={`tab-btn ${activeTab === 'preview' ? 'active' : ''}`} onClick={() => setActiveTab('preview')}>
              \u9884\u89c8
            </button>
            <button className={`tab-btn ${activeTab === 'code' ? 'active' : ''}`} onClick={() => setActiveTab('code')}>
              \u811a\u672c
            </button>
          </div>
        </header>

        <div className="artifact-content">
          {activeTab === 'preview' ? (
            <div className="table-container">
              {!preview || preview.rows.length === 0 ? (
                <div className="empty-state">
                  <FileSpreadsheet size={48} strokeWidth={1} color="#e5e7eb" />
                  <p>{'\u6682\u65e0\u6570\u636e\u9884\u89c8'}</p>
                </div>
              ) : (
                <table>
                  <thead>
                    <tr>
                      {preview.columns.map((col) => (
                        <th key={col}>{col}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.rows.map((row, idx) => (
                      <tr key={idx}>
                        {preview.columns.map((col) => (
                          <td key={col} className={shouldHighlight(col, row[col]) ? 'cell-warn' : ''}>
                            {String(row[col] ?? '')}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
              <textarea className="script-editor" value={opsText} onChange={(e) => setOpsText(e.target.value)} spellCheck={false} />
              {hasParsedScript && (
                <div
                  style={{
                    padding: '1rem',
                    borderTop: '1px solid #e5e7eb',
                    background: '#f9fafb',
                    display: 'flex',
                    justifyContent: 'flex-end',
                  }}
                >
                  <button className="btn btn-primary" onClick={handleApplyOps}>
                    <Play size={16} /> \u6267\u884c\u811a\u672c
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </aside>
    </div>
  )
}

export default App
