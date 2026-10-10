import { useEffect, useRef, useState } from 'react'
import type { ChangeEvent, DragEvent, ReactNode } from 'react'
import {
  ArrowDownToLine,
  ArrowRight,
  Check,
  ChevronDown,
  CircleAlert,
  Code2,
  FileCheck2,
  FileText,
  FolderOpen,
  Image,
  LayoutTemplate,
  LockKeyhole,
  Mail,
  MousePointer2,
  Palette,
  RotateCcw,
  Settings2,
  ShieldCheck,
  Sparkles,
  UploadCloud,
  WandSparkles,
  X,
} from 'lucide-react'
import { extractDocx } from './lib/document'
import type { ContentBlock, ExtractedDocument, InlineNode } from './lib/document'
import { getComponentDefinition, renderCustomComponent, renderCustomComponentDraft, resolveComponentMarkers } from './lib/custom-components'
import type { ComponentBlock, ComponentOverrides } from './lib/custom-components'
import { resolveConditionMarkers } from './lib/conditions'
import { renderEmail } from './lib/email'
import type { EmailSettings, RenderedEmail } from './lib/email'
import { createProjectData, loadProjectFromStorage, parseProject, saveProjectToStorage, serializeProject } from './lib/project'
import { DEFAULT_SETTINGS, FONT_OPTIONS, loadSettings, saveSettings } from './lib/settings'
import { HEADER_HEIGHT_MAX, HEADER_HEIGHT_MIN, HEADER_MARGIN_MAX, HEADER_MARGIN_MIN } from './lib/header'
import type { HeaderSettings } from './lib/header'

type View = 'content' | 'preview' | 'html'
type PreviewDevice = 'desktop' | 'mobile'

const MAX_FILE_SIZE = 20 * 1024 * 1024
const restoredProject = loadProjectFromStorage()

function upgradeDocumentMarkers(value: ExtractedDocument): ExtractedDocument {
  const resolved = resolveComponentMarkers(value.blocks)
  const conditions = resolveConditionMarkers(resolved.blocks)
  const blocks = conditions.blocks.map((block) => block.type === 'component' && block.componentId === 'word-button' && block.initial?.url === ''
    ? { ...block, initial: { ...block.initial, url: '#' } }
    : block)
  const remainingCopy = blocks
    .filter((block) => block.type !== 'component')
    .map((block) => block.runs.map((run) => run.text).join(''))
    .join('\n')
  const priorWarnings = value.warnings.filter((warning) => {
    const outdated = /^O marcador (%[^%\r\n]+%) precisa ficar sozinho em um parágrafo/.exec(warning)
    return !outdated || remainingCopy.includes(outdated[1])
  }).map((warning) => warning.replace(/; informe (?:outra )?(?:a )?URL do botão nas configurações\.$/, '; o botão usará # até você alterar a URL.'))
  return {
    ...value, blocks, warnings: [...new Set([...priorWarnings, ...resolved.warnings])],
    errors: conditions.errors,
  }
}

function App() {
  const [settings, setSettings] = useState<EmailSettings>(() => restoredProject?.settings ?? loadSettings())
  const [document, setDocument] = useState<ExtractedDocument | null>(() => restoredProject ? upgradeDocumentMarkers(restoredProject.document) : null)
  const [componentOverrides, setComponentOverrides] = useState<Record<string, ComponentOverrides>>(restoredProject?.componentOverrides ?? {})
  const [selectedComponentId, setSelectedComponentId] = useState<string | null>(null)
  const [fileSize, setFileSize] = useState<number | null>(null)
  const [view, setView] = useState<View>('content')
  const [previewDevice, setPreviewDevice] = useState<PreviewDevice>('desktop')
  const [belongsToAudience, setBelongsToAudience] = useState(true)
  const [busy, setBusy] = useState(false)
  const [dragging, setDragging] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState(true)
  const [projectSaved, setProjectSaved] = useState(true)
  const [result, setResult] = useState<RenderedEmail | null>(null)
  const [simulatedResult, setSimulatedResult] = useState<RenderedEmail | null>(null)
  const [rendering, setRendering] = useState(false)
  const [simulating, setSimulating] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)
  const [projectDownloadUrl, setProjectDownloadUrl] = useState<string | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const projectInputRef = useRef<HTMLInputElement>(null)
  const componentBlocks = document?.blocks.filter((block): block is ComponentBlock => block.type === 'component') ?? []
  const hasConditions = document?.blocks.some((block) => block.type === 'condition-start' || block.type === 'condition-end') ?? false
  const selectedComponent = componentBlocks.find((block) => block.instanceId === selectedComponentId) ?? componentBlocks[0] ?? null
  const selectedDefinition = selectedComponent ? getComponentDefinition(selectedComponent.componentId) : undefined
  const selectedValues = selectedComponent && selectedDefinition
    ? { ...selectedDefinition.defaults, ...selectedComponent.initial, ...componentOverrides[selectedComponent.instanceId] }
    : null

  useEffect(() => {
    setSaved(saveSettings(settings))
  }, [settings])

  useEffect(() => {
    if (!document) {
      setProjectSaved(true)
      setProjectDownloadUrl(null)
      return
    }
    let url: string | null = null
    try {
      const project = createProjectData(document, componentOverrides, settings)
      setProjectSaved(saveProjectToStorage(project))
      url = URL.createObjectURL(new Blob([serializeProject(project)], { type: 'application/json;charset=utf-8' }))
      setProjectDownloadUrl(url)
    } catch {
      setProjectSaved(false)
      setProjectDownloadUrl(null)
    }
    return () => { if (url) URL.revokeObjectURL(url) }
  }, [document, componentOverrides, settings])

  useEffect(() => {
    if (!document) {
      setResult(null)
      return
    }
    let canceled = false
    setResult(null)
    setRendering(true)
    const timer = window.setTimeout(() => {
      void renderEmail(document, settings, componentOverrides)
        .then((generated) => { if (!canceled) setResult(generated) })
        .catch((cause: unknown) => {
          if (!canceled) setResult({ mjml: '', html: '', errors: [cause instanceof Error ? cause.message : 'Falha ao gerar o email.'] })
        })
        .finally(() => { if (!canceled) setRendering(false) })
    }, 120)
    return () => { canceled = true; window.clearTimeout(timer) }
  }, [document, settings, componentOverrides])

  useEffect(() => {
    if (!document || !hasConditions || document.errors?.length) {
      setSimulatedResult(null)
      setSimulating(false)
      return
    }
    let canceled = false
    setSimulatedResult(null)
    setSimulating(true)
    const timer = window.setTimeout(() => {
      void renderEmail(document, settings, componentOverrides, {
        audienceSimulation: belongsToAudience ? 'member' : 'non-member',
      })
        .then((generated) => { if (!canceled) setSimulatedResult(generated) })
        .catch((cause: unknown) => {
          if (!canceled) setSimulatedResult({ mjml: '', html: '', errors: [cause instanceof Error ? cause.message : 'Falha na simulação.'] })
        })
        .finally(() => { if (!canceled) setSimulating(false) })
    }, 120)
    return () => { canceled = true; window.clearTimeout(timer) }
  }, [document, hasConditions, belongsToAudience, settings, componentOverrides])

  useEffect(() => {
    if (!result?.html || result.errors.length > 0) {
      setDownloadUrl(null)
      return
    }
    const url = URL.createObjectURL(new Blob([result.html], { type: 'text/html;charset=utf-8' }))
    setDownloadUrl(url)
    return () => URL.revokeObjectURL(url)
  }, [result])

  const updateSetting = <Key extends keyof EmailSettings>(key: Key, value: EmailSettings[Key]) => {
    setSettings((current) => ({ ...current, [key]: value }))
  }

  const updateHeader = <Key extends keyof HeaderSettings>(key: Key, value: HeaderSettings[Key]) => {
    setSettings((current) => ({ ...current, header: { ...current.header, [key]: value } }))
  }

  const updateComponent = <Key extends keyof ComponentOverrides>(key: Key, value: ComponentOverrides[Key]) => {
    if (!selectedComponent || !selectedDefinition) return
    const instanceId = selectedComponent.instanceId
    setComponentOverrides((current) => ({
      ...current,
      [instanceId]: { ...selectedDefinition.defaults, ...selectedComponent.initial, ...current[instanceId], [key]: value },
    }))
  }

  function restoreComponent() {
    if (!selectedComponent) return
    setComponentOverrides((current) => {
      const next = { ...current }
      delete next[selectedComponent.instanceId]
      return next
    })
  }

  async function readFile(file: File | undefined) {
    if (!file) return
    setError(null)
    if (!file.name.toLowerCase().endsWith('.docx')) {
      setError('Selecione um documento .docx. Arquivos .doc antigos não são compatíveis.')
      return
    }
    if (file.size > MAX_FILE_SIZE) {
      setError('O arquivo deve ter até 20 MB para processamento local.')
      return
    }
    setBusy(true)
    try {
      const extracted = await extractDocx(file)
      if (extracted.blocks.length === 0) {
        setError('Não encontramos texto neste DOCX. Confira o arquivo e tente novamente.')
        return
      }
      setDocument(extracted)
      setComponentOverrides({})
      setSelectedComponentId(null)
      setFileSize(file.size)
      setView('content')
      setBelongsToAudience(true)
    } catch {
      setError('Não foi possível ler o DOCX. Confirme que o arquivo não está corrompido ou protegido por senha.')
    } finally {
      setBusy(false)
      if (inputRef.current) inputRef.current.value = ''
    }
  }

  function onFileChange(event: ChangeEvent<HTMLInputElement>) {
    void readFile(event.target.files?.[0])
  }

  async function readProjectFile(file: File | undefined) {
    if (!file) return
    setError(null)
    if (!file.name.toLowerCase().endsWith('.json') || file.size > 10_000_000) {
      setError('Selecione um projeto JSON exportado pelo Ateliê Email (até 10 MB).')
      return
    }
    try {
      const project = parseProject(await file.text())
      setDocument(upgradeDocumentMarkers(project.document))
      setComponentOverrides(project.componentOverrides)
      setSettings(project.settings)
      setSelectedComponentId(null)
      setFileSize(null)
      setView('content')
      setBelongsToAudience(true)
    } catch {
      setError('Não foi possível abrir o projeto. Verifique se o arquivo JSON foi gerado por esta versão.')
    } finally {
      if (projectInputRef.current) projectInputRef.current.value = ''
    }
  }

  function onDrop(event: DragEvent<HTMLElement>) {
    event.preventDefault()
    setDragging(false)
    void readFile(event.dataTransfer.files?.[0])
  }

  return (
    <div className="app-shell min-h-screen text-slate-800">
      <div className="mx-auto flex min-h-screen max-w-[1800px] flex-col lg:flex-row">
        <aside className="sidebar flex w-full shrink-0 flex-col border-b border-slate-200/80 bg-white lg:min-h-screen lg:w-[248px] lg:border-b-0 lg:border-r">
          <div className="flex items-center gap-3 px-5 py-5 lg:px-6 lg:py-7">
            <div className="brand-mark flex h-10 w-10 items-center justify-center rounded-2xl text-white shadow-lg shadow-indigo-500/20">
              <Mail size={21} strokeWidth={2.2} />
            </div>
            <div>
              <div className="text-[15px] font-bold tracking-tight text-slate-900">Ateliê Email</div>
              <div className="text-[11px] font-medium tracking-wide text-slate-400">GERADOR LOCAL</div>
            </div>
          </div>

          <div className="hidden flex-1 lg:block">
            <div className="px-6 pt-8 text-[11px] font-bold uppercase tracking-[0.16em] text-slate-400">Espaço de trabalho</div>
            <nav className="mt-4 space-y-1 px-3" aria-label="Etapas do projeto">
              <div className="nav-active flex items-center gap-3 rounded-xl px-4 py-3 text-sm font-semibold text-indigo-700">
                <LayoutTemplate size={18} />
                <span>Editor de email</span>
                <span className="ml-auto rounded-md bg-indigo-100 px-1.5 py-0.5 text-[10px]">01</span>
              </div>
              <RoadmapItem icon={<Image size={18} />} label="Layout por imagem" />
              <RoadmapItem icon={<MousePointer2 size={18} />} label="Editor visual" />
              <RoadmapItem icon={<WandSparkles size={18} />} label="Personalização avançada" />
            </nav>
            <div className="mx-6 mt-8 border-t border-slate-100 pt-6">
              <div className="mb-3 text-[11px] font-bold uppercase tracking-[0.16em] text-slate-400">Nesta etapa</div>
              <div className="space-y-3 text-[12px] text-slate-500">
                <div className="flex items-center gap-2.5"><Check className="text-emerald-500" size={15} /> Leitura fiel do Word</div>
                <div className="flex items-center gap-2.5"><Check className="text-emerald-500" size={15} /> Templates MJML</div>
                <div className="flex items-center gap-2.5"><Check className="text-emerald-500" size={15} /> Exportação HTML</div>
                <div className="flex items-center gap-2.5"><Check className="text-emerald-500" size={15} /> Condições AJO</div>
              </div>
            </div>
          </div>

          <div className="hidden border-t border-slate-100 p-5 lg:block">
            <div className="rounded-2xl border border-indigo-100 bg-indigo-50/70 p-4">
              <div className="mb-2 flex h-8 w-8 items-center justify-center rounded-lg bg-white text-indigo-600"><ShieldCheck size={18} /></div>
              <p className="text-xs font-semibold text-slate-800">Seus arquivos ficam aqui</p>
              <p className="mt-1 text-[11px] leading-relaxed text-slate-500">O documento é processado no seu navegador. Nenhum serviço externo é chamado.</p>
            </div>
          </div>
        </aside>

        <main className="min-w-0 flex-1 px-4 pb-10 pt-6 sm:px-7 lg:px-8 lg:pt-8 xl:px-10">
          <header className="mb-8 flex flex-wrap items-start justify-between gap-4">
            <div>
              <div className="mb-3 flex items-center gap-2 text-xs font-semibold text-indigo-600"><span className="h-1.5 w-1.5 rounded-full bg-indigo-500" /> ETAPA 01 / FUNDAÇÃO</div>
              <h1 className="max-w-xl text-2xl font-bold tracking-[-0.035em] text-slate-900 sm:text-[30px]">Do Word para um email pronto para revisar.</h1>
              <p className="mt-2 max-w-xl text-sm leading-relaxed text-slate-500">Importe seu copy, confira a formatação e gere HTML compatível com clientes de email.</p>
            </div>
            <div className="flex items-center gap-2 rounded-full border border-slate-200 bg-white px-3 py-1.5 text-[11px] font-medium text-slate-500 shadow-sm">
              <span className="h-2 w-2 rounded-full bg-emerald-400" /> Funciona offline após instalar
            </div>
          </header>

          <div className="mb-6 grid grid-cols-3 gap-2 rounded-2xl border border-slate-200/80 bg-white p-2 shadow-sm sm:gap-4 sm:p-3">
            <Step number="01" label="Importar" active />
            <Step number="02" label="Revisar" active={Boolean(document)} />
            <Step number="03" label="Exportar" active={Boolean(result?.html && !result.errors.length)} />
          </div>

          <section aria-labelledby="upload-title" className={`upload-card relative overflow-hidden rounded-[22px] border-2 border-dashed bg-white p-5 transition sm:p-7 ${dragging ? 'border-indigo-500 bg-indigo-50' : 'border-slate-200'}`}
            onDragOver={(event) => { event.preventDefault(); setDragging(true) }}
            onDragLeave={(event) => { event.preventDefault(); setDragging(false) }}
            onDrop={onDrop}>
            <input ref={inputRef} type="file" className="sr-only" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={onFileChange} aria-label="Selecionar documento Word DOCX" />
            <div className="flex flex-col items-start gap-5 sm:flex-row sm:items-center">
              <div className="upload-icon flex h-14 w-14 shrink-0 items-center justify-center rounded-2xl text-indigo-600"><UploadCloud size={27} /></div>
              <div className="min-w-0 flex-1">
                <div className="mb-1 flex flex-wrap items-center gap-2"><h2 id="upload-title" className="text-base font-semibold text-slate-900">Importar documento Word</h2><span className="rounded-full bg-slate-100 px-2 py-0.5 text-[10px] font-bold text-slate-500">.DOCX</span></div>
                <p className="text-xs leading-relaxed text-slate-500">Arraste o arquivo aqui ou selecione no computador. Textos, negritos, itálicos e links serão preservados.</p>
              </div>
              <button type="button" disabled={busy} onClick={() => inputRef.current?.click()} className="primary-button flex shrink-0 items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-xs font-semibold text-white disabled:cursor-wait disabled:opacity-60">
                <FolderOpen size={16} />{busy ? 'Lendo arquivo...' : document ? 'Trocar DOCX' : 'Escolher arquivo'}
              </button>
            </div>
          </section>

          <div className="mt-3 flex flex-wrap items-center gap-2 text-xs">
            <input ref={projectInputRef} type="file" className="sr-only" accept=".json,application/json" aria-label="Selecionar projeto JSON" onChange={(event) => { void readProjectFile(event.target.files?.[0]) }} />
            <button type="button" onClick={() => projectInputRef.current?.click()} className="secondary-button rounded-xl px-3 py-2 font-semibold"><FolderOpen size={14} className="mr-1 inline" /> Abrir projeto</button>
            {document && <a href={projectDownloadUrl ?? undefined} download={`${document.sourceName.replace(/\.docx$/i, '') || 'email'}-projeto.json`} aria-disabled={!projectDownloadUrl} onClick={(event) => { if (!projectDownloadUrl) event.preventDefault() }} className={`secondary-button rounded-xl px-3 py-2 font-semibold ${!projectDownloadUrl ? 'pointer-events-none opacity-40' : ''}`}><ArrowDownToLine size={14} className="mr-1 inline" /> Salvar projeto JSON</a>}
            {document && <span className={`ml-auto ${projectSaved ? 'text-emerald-700' : 'text-amber-700'}`}>{projectSaved ? 'Projeto salvo neste navegador' : 'Não foi possível salvar o projeto neste navegador'}</span>}
          </div>

          {error && <div role="alert" className="mt-4 flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-xs text-rose-700"><CircleAlert size={16} className="shrink-0" />{error}<button className="ml-auto" onClick={() => setError(null)} aria-label="Fechar erro"><X size={15} /></button></div>}

          {document && <div className="mt-4 flex flex-wrap items-center gap-x-3 gap-y-2 rounded-xl border border-emerald-100 bg-emerald-50/80 px-4 py-3 text-xs text-slate-700">
            <FileCheck2 size={17} className="text-emerald-600" />
            <span className="max-w-[280px] truncate font-semibold" title={document.sourceName}>{document.sourceName}</span>
            <span className="text-slate-400">{fileSize === null ? '' : formatBytes(fileSize)} · {document.blocks.length} blocos</span>
            <span className="ml-auto flex items-center gap-1 font-medium text-emerald-700"><Check size={13} /> Extração concluída</span>
          </div>}

          {document?.warnings.length ? <div role="status" className="mt-4 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-xs leading-relaxed text-amber-900">
            <div className="mb-1 flex items-center gap-2 font-semibold"><CircleAlert size={15} /> Revise estes detalhes do documento</div>
            <ul className="list-disc space-y-1 pl-5">{document.warnings.map((warning, index) => <li key={`${warning}-${index}`}>{warning}</li>)}</ul>
          </div> : null}

          {document?.errors?.length ? <div role="alert" className="mt-4 rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-xs leading-relaxed text-rose-800">
            <div className="mb-1 flex items-center gap-2 font-semibold"><CircleAlert size={15} /> Corrija os marcadores de audiência no Word antes de exportar</div>
            <ul className="list-disc space-y-1 pl-5">{document.errors.map((issue, index) => <li key={`${issue}-${index}`}>{issue}</li>)}</ul>
          </div> : null}

          {hasConditions && <section aria-label="Simulação de audiência" className="mt-4 rounded-xl border border-indigo-100 bg-indigo-50/60 px-4 py-3 text-xs">
            <div className="mb-2 font-semibold text-indigo-900">Simular participação nas audiências</div>
            <div className="flex flex-wrap gap-2">
              <button type="button" aria-pressed={belongsToAudience} onClick={() => setBelongsToAudience(true)} className={`rounded-lg border px-3 py-1.5 font-semibold ${belongsToAudience ? 'border-indigo-400 bg-white text-indigo-700' : 'border-slate-200 bg-white text-slate-500'}`}>Perfil pertence</button>
              <button type="button" aria-pressed={!belongsToAudience} onClick={() => setBelongsToAudience(false)} className={`rounded-lg border px-3 py-1.5 font-semibold ${!belongsToAudience ? 'border-indigo-400 bg-white text-indigo-700' : 'border-slate-200 bg-white text-slate-500'}`}>Perfil não pertence</button>
            </div>
            <p className="mt-2 text-[11px] text-indigo-700">A simulação afeta apenas as prévias Conteúdo e Email. O HTML exportado mantém as expressões AJO.</p>
          </section>}

          <section className="mt-6 overflow-hidden rounded-[22px] border border-slate-200/80 bg-white shadow-sm" aria-label="Área de revisão">
            <div className="flex flex-col gap-4 border-b border-slate-100 px-5 py-5 sm:flex-row sm:items-center sm:justify-between sm:px-6">
              <div><h2 className="text-base font-semibold text-slate-900">Área de revisão</h2><p className="mt-0.5 text-xs text-slate-400">Confira o copy antes de usar o HTML.</p></div>
              <div className="flex items-center gap-1 rounded-xl bg-slate-100 p-1" role="tablist" aria-label="Visualização">
                <ViewTab label="Conteúdo" icon={<FileText size={14} />} selected={view === 'content'} onClick={() => setView('content')} />
                <ViewTab label="Email" icon={<Mail size={14} />} selected={view === 'preview'} onClick={() => setView('preview')} />
                <ViewTab label="HTML" icon={<Code2 size={14} />} selected={view === 'html'} onClick={() => setView('html')} />
              </div>
            </div>

            {document && view === 'preview' && <div className="flex items-center justify-end gap-1 border-b border-slate-100 px-5 py-2 text-[11px]" aria-label="Largura da prévia">
              <span className="mr-2 text-slate-500">Visualização</span>
              <button type="button" aria-pressed={previewDevice === 'desktop'} onClick={() => setPreviewDevice('desktop')} className={`rounded-lg px-3 py-1.5 font-semibold ${previewDevice === 'desktop' ? 'bg-indigo-50 text-indigo-700' : 'text-slate-500 hover:bg-slate-50'}`}>Desktop</button>
              <button type="button" aria-pressed={previewDevice === 'mobile'} onClick={() => setPreviewDevice('mobile')} className={`rounded-lg px-3 py-1.5 font-semibold ${previewDevice === 'mobile' ? 'bg-indigo-50 text-indigo-700' : 'text-slate-500 hover:bg-slate-50'}`}>Mobile</button>
            </div>}

            {!document ? <EmptyState onUpload={() => inputRef.current?.click()} /> : view === 'content' ? <div className="min-h-[360px] p-5 sm:p-7"><ContentPreview blocks={document.blocks} overrides={componentOverrides} selectedId={selectedComponent?.instanceId ?? null} onSelect={setSelectedComponentId} belongsToAudience={belongsToAudience} /></div> : rendering || (view === 'preview' && hasConditions && simulating) ? <div className="flex min-h-[360px] items-center justify-center text-sm text-slate-500">Gerando prévia local...</div> : result?.errors.length ? <div role="alert" className="min-h-[360px] p-7 text-sm text-rose-700">Não foi possível gerar o email: {result.errors.join(' · ')}</div> : view === 'preview' && simulatedResult?.errors.length ? <div role="alert" className="min-h-[360px] p-7 text-sm text-rose-700">Não foi possível simular a audiência: {simulatedResult.errors.join(' · ')}</div> : view === 'preview' ? <div className="preview-stage min-h-[450px] p-3 sm:p-6"><iframe className="mx-auto h-[650px] w-full rounded-xl border border-slate-200 bg-white shadow-sm" style={{ maxWidth: previewDevice === 'desktop' ? `${settings.width}px` : '375px' }} title="Prévia do HTML do email" sandbox="" srcDoc={(hasConditions ? simulatedResult : result)?.html ?? ''} /></div> : <div className="min-h-[450px] p-5 sm:p-6"><div className="mb-3 flex items-center justify-between text-xs text-slate-500"><span>HTML gerado por MJML e componentes preservados</span><span>{result?.html.length.toLocaleString('pt-BR')} caracteres</span></div><pre className="code-preview max-h-[650px] overflow-auto rounded-xl p-5 text-[11px] leading-relaxed text-slate-200"><code>{result?.html}</code></pre></div>}

            <div className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 px-5 py-4 sm:px-6">
              <div className="flex items-center gap-2 text-[11px] text-slate-500"><ShieldCheck size={15} className="text-emerald-500" /> Copy preservado; botões usam a URL do template ou a que você definir</div>
              <div className="flex gap-2">
                <button type="button" disabled={!document || rendering || Boolean(result?.errors.length)} onClick={() => setView('preview')} className="secondary-button rounded-xl px-4 py-2.5 text-xs font-semibold disabled:cursor-not-allowed disabled:opacity-40">Visualizar email</button>
                <a href={downloadUrl ?? undefined} download={`${document?.sourceName.replace(/\.docx$/i, '') || 'email'}-email.html`} role="button" aria-disabled={!downloadUrl || rendering} tabIndex={downloadUrl && !rendering ? 0 : -1} onClick={(event) => { if (!downloadUrl || rendering) event.preventDefault() }} className={`primary-button flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs font-semibold text-white ${!downloadUrl || rendering ? 'cursor-not-allowed opacity-40' : ''}`}><ArrowDownToLine size={15} /> Exportar HTML</a>
              </div>
            </div>
          </section>
        </main>

        <aside className="settings-panel w-full shrink-0 border-t border-slate-200 bg-white lg:min-h-screen lg:w-[304px] lg:border-l lg:border-t-0 xl:w-[330px]">
          <div className="sticky top-0 p-5 lg:max-h-screen lg:overflow-y-auto lg:p-6">
            <div className="mb-7 flex items-center justify-between"><div className="flex items-center gap-2"><Settings2 size={17} className="text-slate-500" /><h2 className="text-sm font-bold text-slate-900">Configurações</h2></div><span className="rounded-full bg-slate-100 px-2.5 py-1 text-[10px] font-semibold text-slate-500">LOCAL</span></div>

            <section className="mb-7 rounded-2xl border border-slate-200 bg-slate-50/70 p-4" aria-label="Cabeçalho">
              <div className="mb-3 flex items-center justify-between gap-2">
                <div className="flex items-center gap-2 text-xs font-semibold text-slate-800"><LayoutTemplate size={15} className="text-indigo-600" /> Cabeçalho</div>
                <label className="flex cursor-pointer items-center gap-2 text-[11px] font-medium text-slate-600"><input type="checkbox" checked={settings.header.enabled} onChange={(event) => updateHeader('enabled', event.target.checked)} className="accent-indigo-600" /> Exibir cabeçalho</label>
              </div>
              <div className="space-y-3">
                <HeaderColorField value={settings.header.backgroundColor} onChange={(value) => updateHeader('backgroundColor', value)} />
                <PixelField label="Altura do cabeçalho" value={settings.header.height} min={HEADER_HEIGHT_MIN} max={HEADER_HEIGHT_MAX} onChange={(value) => updateHeader('height', value)} />
                <PixelField label="Espaçamento acima" value={settings.header.marginTop} min={HEADER_MARGIN_MIN} max={HEADER_MARGIN_MAX} onChange={(value) => updateHeader('marginTop', value)} />
                <PixelField label="Espaçamento abaixo" value={settings.header.marginBottom} min={HEADER_MARGIN_MIN} max={HEADER_MARGIN_MAX} onChange={(value) => updateHeader('marginBottom', value)} />
              </div>
              <p className="mt-3 text-[10px] leading-relaxed text-slate-500">Largura herdada do email: {settings.width} px. O cabeçalho não usa imagem, link ou texto.</p>
            </section>

            {componentBlocks.length > 0 && <section className="mb-7 rounded-2xl border border-indigo-100 bg-indigo-50/40 p-4" aria-label="Componentes personalizados">
              <div className="mb-1 flex items-center gap-2 text-xs font-semibold text-slate-800"><MousePointer2 size={15} className="text-indigo-600" /> Botões do Word</div>
              <p className="mb-3 text-[11px] leading-relaxed text-slate-500">Selecione um marcador %texto% para editar o botão e seu destino.</p>
              <div className="mb-4 flex flex-wrap gap-1.5">{componentBlocks.map((block, index) => <button key={block.instanceId} type="button" onClick={() => { setSelectedComponentId(block.instanceId); setView('content') }} aria-pressed={selectedComponent?.instanceId === block.instanceId} className={`rounded-lg border px-2.5 py-1.5 text-[11px] font-semibold ${selectedComponent?.instanceId === block.instanceId ? 'border-indigo-400 bg-white text-indigo-700' : 'border-slate-200 bg-white/60 text-slate-600 hover:border-indigo-300'}`}>{getComponentDefinition(block.componentId)?.name ?? block.componentId} {index + 1}</button>)}</div>
              {selectedComponent && selectedDefinition && selectedValues && <div className="space-y-3">
                <label className="block text-[11px] font-medium text-slate-600" htmlFor="component-text">Texto exibido</label>
                <input id="component-text" type="text" maxLength={1000} value={selectedValues.text} onChange={(event) => updateComponent('text', event.target.value)} className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-800 outline-none focus:border-indigo-400" />
                <label className="block text-[11px] font-medium text-slate-600" htmlFor="component-url">URL de destino</label>
                <input id="component-url" type="text" inputMode="url" maxLength={4096} value={selectedValues.url} onChange={(event) => updateComponent('url', event.target.value)} placeholder="https://... ou #" className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-800 outline-none focus:border-indigo-400" />
                {(selectedValues.url === '#' || selectedValues.url === '') && <p className="text-[11px] text-slate-500">Sem redirecionamento: o HTML será exportado com href="#".</p>}
                {!canExportComponent(selectedComponent, componentOverrides) && <p role="alert" className="text-[11px] text-rose-700">Use uma URL HTTP ou HTTPS válida, ou #.</p>}
                <div className="overflow-hidden rounded-xl border border-slate-200 bg-white"><ComponentSnippetPreview block={selectedComponent} overrides={componentOverrides} title="Prévia do botão selecionado" /></div>
                <button type="button" onClick={restoreComponent} className="flex items-center gap-1 text-[11px] font-semibold text-indigo-700 hover:text-indigo-900"><RotateCcw size={13} /> Restaurar texto e URL iniciais</button>
              </div>}
            </section>}

            <div className="mb-7">
              <div className="mb-3 flex items-center gap-2 text-xs font-semibold text-slate-700"><LayoutTemplate size={15} className="text-indigo-500" /> Template</div>
              <div className="grid grid-cols-2 gap-2">
                <TemplateCard label="Clássico" description="Cartão central" selected={settings.template === 'classic'} onClick={() => updateSetting('template', 'classic')} compact={false} />
                <TemplateCard label="Compacto" description="Mais direto" selected={settings.template === 'compact'} onClick={() => updateSetting('template', 'compact')} compact />
              </div>
            </div>

            <div className="mb-7 border-t border-slate-100 pt-6">
              <div className="mb-4 flex items-center gap-2 text-xs font-semibold text-slate-700"><Mail size={15} className="text-indigo-500" /> Dimensões e tipografia</div>
              <label htmlFor="email-width" className="mb-2 flex justify-between text-xs text-slate-500"><span>Largura do email</span><span className="font-semibold text-slate-700">{settings.width} px</span></label>
              <input id="email-width" type="range" min="480" max="760" step="10" value={settings.width} onChange={(event) => updateSetting('width', Number(event.target.value))} className="w-full accent-indigo-600" />
              <div className="mt-1 flex justify-between text-[10px] text-slate-400"><span>480 px</span><span>760 px</span></div>
              <label htmlFor="font-family" className="mt-5 block text-xs text-slate-500">Fonte</label>
              <div className="relative mt-2"><select id="font-family" value={settings.fontFamily} onChange={(event) => updateSetting('fontFamily', event.target.value)} className="w-full appearance-none rounded-xl border border-slate-200 bg-white px-3 py-2.5 pr-9 text-xs font-medium text-slate-700 outline-none focus:border-indigo-400">{FONT_OPTIONS.map((font) => <option key={font.label} value={font.value}>{font.label}</option>)}</select><ChevronDown size={15} className="pointer-events-none absolute right-3 top-3 text-slate-400" /></div>
            </div>

            <div className="border-t border-slate-100 pt-6">
              <div className="mb-4 flex items-center gap-2 text-xs font-semibold text-slate-700"><Palette size={15} className="text-indigo-500" /> Cores</div>
              <div className="space-y-3">
                <ColorField label="Fundo" value={settings.backgroundColor} onChange={(value) => updateSetting('backgroundColor', value)} />
                <ColorField label="Superfície" value={settings.surfaceColor} onChange={(value) => updateSetting('surfaceColor', value)} />
                <ColorField label="Texto" value={settings.textColor} onChange={(value) => updateSetting('textColor', value)} />
                <ColorField label="Destaque" value={settings.accentColor} onChange={(value) => updateSetting('accentColor', value)} />
              </div>
            </div>

            <div className="mt-7 rounded-xl bg-slate-50 p-3 text-[11px] leading-relaxed text-slate-500"><div className="mb-1 flex items-center gap-1.5 font-semibold text-slate-700"><span className={`h-1.5 w-1.5 rounded-full ${saved && projectSaved ? 'bg-emerald-400' : 'bg-amber-400'}`} /> {saved && projectSaved ? 'Preferências e projeto salvos' : 'Armazenamento indisponível'}</div>{saved && projectSaved ? 'Suas alterações serão lembradas neste navegador.' : 'As alterações funcionam nesta sessão, mas não puderam ser salvas.'}</div>
            <button type="button" onClick={() => setSettings(DEFAULT_SETTINGS)} className="mt-4 w-full text-center text-[11px] font-medium text-slate-400 hover:text-indigo-600">Restaurar padrões</button>
          </div>
        </aside>
      </div>
    </div>
  )
}

function RoadmapItem({ icon, label }: { icon: ReactNode; label: string }) {
  return <div className="flex items-center gap-3 rounded-xl px-4 py-3 text-sm text-slate-400" title="Previsto para etapas futuras">{icon}<span>{label}</span><LockKeyhole className="ml-auto opacity-45" size={13} /></div>
}

function Step({ number, label, active }: { number: string; label: string; active: boolean }) {
  return <div className={`flex items-center gap-2 rounded-xl px-2 py-2 sm:px-3 ${active ? 'bg-indigo-50 text-indigo-700' : 'text-slate-400'}`}><span className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg text-[10px] font-bold ${active ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-400'}`}>{number}</span><span className="truncate text-[11px] font-semibold sm:text-xs">{label}</span><ArrowRight className="ml-auto hidden opacity-40 sm:block" size={13} /></div>
}

function ViewTab({ label, icon, selected, onClick }: { label: string; icon: ReactNode; selected: boolean; onClick: () => void }) {
  return <button type="button" role="tab" aria-selected={selected} onClick={onClick} className={`flex items-center gap-1.5 rounded-lg px-2.5 py-2 text-[11px] font-semibold transition sm:px-3 ${selected ? 'bg-white text-indigo-700 shadow-sm' : 'text-slate-500 hover:text-slate-800'}`}>{icon}<span>{label}</span></button>
}

function EmptyState({ onUpload }: { onUpload: () => void }) {
  return <div className="flex min-h-[360px] flex-col items-center justify-center px-5 text-center"><div className="mb-4 flex h-16 w-16 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-500"><FileText size={30} strokeWidth={1.6} /></div><div className="mb-2 flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.14em] text-indigo-500"><Sparkles size={13} /> PRONTO PARA COMEÇAR</div><h3 className="text-base font-semibold text-slate-800">Seu conteúdo aparecerá aqui</h3><p className="mt-1 max-w-xs text-xs leading-relaxed text-slate-500">Importe um DOCX para revisar a ordem dos parágrafos, estilos e links antes de exportar.</p><button type="button" onClick={onUpload} className="mt-5 flex items-center gap-1 text-xs font-semibold text-indigo-600 hover:text-indigo-800">Selecionar documento <ArrowRight size={14} /></button></div>
}

function canExportComponent(block: ComponentBlock, overrides: Record<string, ComponentOverrides>): boolean {
  try {
    renderCustomComponent(block.componentId, { ...block.initial, ...overrides[block.instanceId] })
    return true
  } catch {
    return false
  }
}

function safeComponentPreview(block: ComponentBlock, overrides: Record<string, ComponentOverrides>): string | null {
  try {
    return renderCustomComponentDraft(block.componentId, { ...block.initial, ...overrides[block.instanceId] })
  } catch {
    return null
  }
}

function ComponentSnippetPreview({ block, overrides, title }: { block: ComponentBlock; overrides: Record<string, ComponentOverrides>; title: string }) {
  const html = safeComponentPreview(block, overrides)
  if (!html) return <div className="px-3 py-4 text-center text-[11px] text-rose-700">Corrija a URL para visualizar este botão.</div>
  return <iframe className="pointer-events-none h-[100px] w-full bg-white" title={title} sandbox="" srcDoc={html} />
}

function ContentPreview({ blocks, overrides, selectedId, onSelect, belongsToAudience }: {
  blocks: ContentBlock[]
  overrides: Record<string, ComponentOverrides>
  selectedId: string | null
  onSelect: (instanceId: string) => void
  belongsToAudience: boolean
}) {
  let orderedIndex = 0
  let insideCondition = false
  return <div className="copy-preview mx-auto max-w-[660px] text-sm leading-[1.8] text-slate-700">{blocks.map((block, index) => {
    if (block.type === 'condition-start') {
      insideCondition = true
      orderedIndex = 0
      return <div key={`condition-start-${index}`} className="mb-3 rounded-lg border-l-4 border-indigo-500 bg-indigo-50 px-3 py-1.5 text-xs font-bold text-indigo-800">Condição: {block.audience}{!belongsToAudience && <span className="ml-2 font-normal">· conteúdo oculto nesta simulação</span>}</div>
    }
    if (block.type === 'condition-end') {
      insideCondition = false
      orderedIndex = 0
      return <div key={`condition-end-${index}`} className="mb-3 rounded-lg border-l-4 border-indigo-300 bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-500">Fim da condição: {block.audience}</div>
    }
    if (insideCondition && !belongsToAudience) return null
    if (block.type === 'list-item' && block.ordered) orderedIndex += 1
    else if (block.type !== 'list-item') orderedIndex = 0
    if (block.type === 'component') return <div key={block.instanceId} className={`mb-3 overflow-hidden rounded-xl border ${selectedId === block.instanceId ? 'border-indigo-400 ring-2 ring-indigo-100' : 'border-slate-200'}`}>
      <div className="flex items-center justify-between bg-slate-50 px-3 py-1.5 text-[11px]"><span className="font-semibold text-slate-600">{getComponentDefinition(block.componentId)?.name ?? block.componentId}</span><button type="button" onClick={() => onSelect(block.instanceId)} className="font-semibold text-indigo-700 hover:text-indigo-900">Editar botão</button></div>
      <ComponentSnippetPreview block={block} overrides={overrides} title={`Prévia do botão ${block.instanceId}`} />
    </div>
    const content = <RichText runs={block.runs} />
    if (block.type === 'heading') return <h3 key={index} className="mb-3 mt-5 text-lg font-bold leading-snug text-slate-900 first:mt-0">{content}</h3>
    if (block.type === 'list-item') return <div key={index} className="mb-1 flex gap-3 pl-3"><span className="shrink-0 text-indigo-500">{block.ordered ? `${orderedIndex}.` : '•'}</span><span>{content}</span></div>
    return <p key={index} className="mb-3 min-h-[1em] last:mb-0">{content}</p>
  })}</div>
}

function RichText({ runs }: { runs: InlineNode[] }) {
  return <>{runs.map((run, index) => {
    let content: ReactNode = run.text
    if (run.bold) content = <strong>{content}</strong>
    if (run.italic) content = <em>{content}</em>
    if (run.href) content = <a href={run.href} target="_blank" rel="noopener noreferrer" className="text-indigo-600 underline underline-offset-2">{content}</a>
    return <span key={index}>{content}</span>
  })}</>
}

function TemplateCard({ label, description, selected, onClick, compact }: { label: string; description: string; selected: boolean; onClick: () => void; compact: boolean }) {
  return <button type="button" aria-pressed={selected} onClick={onClick} className={`rounded-xl border p-2 text-left transition ${selected ? 'border-indigo-400 bg-indigo-50/40 ring-2 ring-indigo-100' : 'border-slate-200 hover:border-slate-300'}`}><div className="mb-2.5 flex h-12 items-center justify-center rounded-lg bg-slate-100 px-3"><div className={`h-9 w-12 rounded-[2px] border border-slate-200 bg-white p-1 shadow-sm ${compact ? '' : 'outline outline-2 outline-offset-2 outline-slate-200'}`}><div className="mb-1 h-1 w-5 rounded bg-indigo-300" /><div className="mb-0.5 h-0.5 w-full rounded bg-slate-200" /><div className="h-0.5 w-8 rounded bg-slate-200" /></div></div><div className="flex items-start justify-between gap-1"><div><div className="text-[11px] font-semibold text-slate-800">{label}</div><div className="text-[10px] text-slate-400">{description}</div></div>{selected && <Check size={14} className="text-indigo-600" />}</div></button>
}

function ColorField({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  return <label className="flex cursor-pointer items-center justify-between gap-3 text-xs text-slate-500"><span>{label}</span><span className="flex items-center gap-2 rounded-lg border border-slate-200 px-2 py-1"><input type="color" value={value} onChange={(event) => onChange(event.target.value)} aria-label={`Cor de ${label.toLowerCase()}`} className="h-6 w-6 cursor-pointer rounded border-0 bg-transparent p-0" /><span className="w-[62px] font-mono text-[10px] font-semibold uppercase text-slate-600">{value}</span></span></label>
}

function HeaderColorField({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const [draft, setDraft] = useState(value)
  useEffect(() => setDraft(value), [value])
  const valid = /^#[0-9a-fA-F]{6}$/.test(draft)

  return <div>
    <label htmlFor="header-color" className="mb-1 block text-[11px] font-medium text-slate-600">Cor de fundo</label>
    <div className="flex items-center gap-2">
      <input type="color" value={value} onChange={(event) => onChange(event.target.value)} aria-label="Selecionar cor do cabeçalho" className="h-9 w-10 shrink-0 cursor-pointer rounded-lg border border-slate-200 bg-white p-1" />
      <input id="header-color" type="text" maxLength={7} value={draft} onChange={(event) => {
        const next = event.target.value
        setDraft(next)
        if (/^#[0-9a-fA-F]{6}$/.test(next)) onChange(next)
      }} onBlur={() => { if (!valid) setDraft(value) }} aria-invalid={!valid} className={`min-w-0 flex-1 rounded-lg border bg-white px-2.5 py-2 font-mono text-[11px] uppercase outline-none ${valid ? 'border-slate-200 focus:border-indigo-400' : 'border-rose-300 focus:border-rose-400'}`} />
    </div>
    {!valid && <p className="mt-1 text-[10px] text-rose-700">Use uma cor hexadecimal como #000000.</p>}
  </div>
}

function PixelField({ label, value, min, max, onChange }: {
  label: string; value: number; min: number; max: number; onChange: (value: number) => void
}) {
  const [draft, setDraft] = useState(String(value))
  useEffect(() => setDraft(String(value)), [value])
  return <label className="flex items-center justify-between gap-2 text-[11px] font-medium text-slate-600">
    <span>{label}</span>
    <span className="flex items-center gap-1"><input type="number" min={min} max={max} step="1" value={draft} onChange={(event) => {
      const next = event.target.value
      setDraft(next)
      const numeric = Number(next)
      if (next !== '' && Number.isInteger(numeric) && numeric >= min && numeric <= max) onChange(numeric)
    }} onBlur={() => {
      const numeric = Number(draft)
      const normalized = draft.trim() && Number.isFinite(numeric) ? Math.min(max, Math.max(min, Math.round(numeric))) : value
      onChange(normalized)
      setDraft(String(normalized))
    }} className="w-20 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-right text-[11px] text-slate-800 outline-none focus:border-indigo-400" /><span className="text-slate-400">px</span></span>
  </label>
}

function formatBytes(bytes: number) {
  return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

export default App
