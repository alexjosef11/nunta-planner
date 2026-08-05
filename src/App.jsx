import { useState, useEffect, useCallback, useRef } from "react"
import * as XLSX from "xlsx"
import { dbGet, dbSet, dbSubscribe } from "./supabase.js"

// ─── ICONS ───────────────────────────────────────────────────────────────────
const Icon = ({ d, size = 18, stroke = "currentColor", fill = "none", strokeWidth = 1.6 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={fill} stroke={stroke} strokeWidth={strokeWidth} strokeLinecap="round" strokeLinejoin="round">
    <path d={d} />
  </svg>
)
const HeartIcon    = () => <Icon d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z" fill="currentColor" />
const UsersIcon    = () => <Icon d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" />
const CheckIcon    = () => <Icon d="M9 11l3 3L22 4M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
const TableIcon    = () => <Icon d="M3 3h18v4H3zM3 10h18M3 17h18M3 10v7M21 10v7M8 10v7M16 10v7" />
const WalletIcon   = () => <Icon d="M21 12V7a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-5M21 12h-4a2 2 0 0 0 0 4h4" />
const PlusIcon     = () => <Icon d="M12 5v14M5 12h14" strokeWidth={2} />
const TrashIcon    = () => <Icon d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
const EditIcon     = () => <Icon d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
const XIcon        = () => <Icon d="M18 6L6 18M6 6l12 12" strokeWidth={2} />
const SearchIcon   = () => <Icon d="M21 21l-4.35-4.35M17 11A6 6 0 1 1 5 11a6 6 0 0 1 12 0z" />
const SyncIcon     = () => <Icon d="M4 4v5h.582M20 20v-5h-.581M4.582 9a8 8 0 0 1 15.356 2M19.419 15A8 8 0 0 1 4.064 13" strokeWidth={1.8} />
const UserPlusIcon = () => <Icon d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M8.5 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM20 8v6M23 11h-6" />
const DownloadIcon = () => <Icon d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3" />
const CopyIcon     = () => <Icon d="M20 9H11a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h9a2 2 0 0 0 2-2v-9a2 2 0 0 0-2-2zM5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
const BriefcaseIcon = ({ size = 18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="7" width="20" height="14" rx="2" />
    <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16M2 13h20" />
  </svg>
)

// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function exportToExcel(sheets, filename) {
  const wb = XLSX.utils.book_new()
  sheets.forEach(({ name, data }) => {
    const ws = XLSX.utils.json_to_sheet(data)
    const cols = Object.keys(data[0] || {}).map(key => ({
      wch: Math.max(key.length, ...data.map(r => String(r[key] ?? "").length)) + 2
    }))
    ws["!cols"] = cols
    XLSX.utils.book_append_sheet(wb, ws, name)
  })
  XLSX.writeFile(wb, filename)
}

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const C = {
  bg: "#0F0E17", surface: "#1A1825", surfaceHi: "#232136",
  border: "#2E2B3E", borderHi: "#3D3A52",
  accent: "#C084A8", accentAlt: "#9B6FD4",
  accentGrad: "linear-gradient(135deg, #C084A8 0%, #9B6FD4 100%)",
  accentGlow: "0 0 24px rgba(192,132,168,0.35)",
  gold: "#E2B96F", goldDim: "#3D3018",
  sage: "#6EC9A0", sageDim: "#163328",
  textPri: "#F0EDF8", textSec: "#9D99B8", textDim: "#5E5A7A",
  danger: "#F07070", dangerDim: "#3D1515",
  purple: "#9B6FD4", purpleDim: "#251A3D",
}

// ─── GLOBAL STYLES ────────────────────────────────────────────────────────────
const globalStyles = `
  *, *::before, *::after { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  html, body, #root { margin: 0; padding: 0; min-height: 100vh; background: ${C.bg}; }
  body { font-family: 'Inter', 'Segoe UI', system-ui, sans-serif; color: ${C.textPri}; }
  ::-webkit-scrollbar { width: 4px; height: 4px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 4px; }
  input, select, textarea { color-scheme: dark; }
  input:focus, select:focus, textarea:focus { outline: none; border-color: ${C.accent} !important; }
  @keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
  .fadeup { animation: fadeUp 0.2s ease; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .spin { animation: spin 1s linear infinite; }
`

// ─── STYLES ───────────────────────────────────────────────────────────────────
const S = {
  app: { minHeight: "100vh", background: C.bg, color: C.textPri },
  topbar: { background: C.surface + "EE", borderBottom: `1px solid ${C.border}`, padding: "0 20px", position: "sticky", top: 0, zIndex: 200, backdropFilter: "blur(16px)", WebkitBackdropFilter: "blur(16px)" },
  topbarInner: { maxWidth: 1160, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 56 },
  logo: { display: "flex", alignItems: "center", gap: 8, color: C.accent, fontSize: 16, fontWeight: 700, letterSpacing: "-0.01em" },
  bottomNav: { position: "fixed", bottom: 0, left: 0, right: 0, zIndex: 200, background: C.surface + "F5", borderTop: `1px solid ${C.border}`, display: "flex", backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", paddingBottom: "env(safe-area-inset-bottom, 6px)" },
  bottomNavBtn: (a) => ({ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 3, padding: "10px 4px 8px", border: "none", background: "none", cursor: "pointer", color: a ? C.accent : C.textDim, transition: "color 0.15s" }),
  bottomNavLabel: (a) => ({ fontSize: 9, fontWeight: a ? 700 : 400, letterSpacing: "0.06em", textTransform: "uppercase", fontFamily: "inherit" }),
  main: (isMobile) => ({ maxWidth: 1160, margin: "0 auto", padding: isMobile ? "20px 16px 96px" : "28px 24px 48px" }),
  card: { background: C.surface, borderRadius: 18, border: `1px solid ${C.border}`, overflow: "hidden", marginBottom: 14, boxShadow: "0 2px 20px rgba(0,0,0,0.35)" },
  cardHeader: { padding: "16px 20px 12px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" },
  cardTitle: { fontSize: 14, fontWeight: 700, color: C.textPri, letterSpacing: "-0.01em" },
  cardBody: { padding: "14px 20px" },
  statsRow: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(110px, 1fr))", gap: 10, marginBottom: 16 },
  statCard: { background: C.surfaceHi, border: `1px solid ${C.border}`, borderRadius: 14, padding: "14px 12px", textAlign: "center" },
  statNum: { fontSize: 26, fontWeight: 800, background: C.accentGrad, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text", lineHeight: 1.1, letterSpacing: "-0.03em" },
  statLabel: { fontSize: 9, color: C.textDim, textTransform: "uppercase", letterSpacing: "0.08em", marginTop: 4, fontWeight: 600 },
  btn: (v = "primary") => {
    const variants = {
      primary: { background: C.accentGrad, color: "#fff", border: "none", boxShadow: C.accentGlow },
      sage: { background: `linear-gradient(135deg,${C.sage},#4BAF84)`, color: "#0A1F16", border: "none" },
      danger: { background: C.dangerDim, color: C.danger, border: `1px solid ${C.danger}44` },
      ghost: { background: "transparent", color: C.textSec, border: "none" },
      outline: { background: "transparent", color: C.textSec, border: `1.5px solid ${C.border}` },
      purple: { background: `linear-gradient(135deg,${C.purple},#7B52C4)`, color: "#fff", border: "none" },
      gold: { background: C.goldDim, color: C.gold, border: `1px solid ${C.gold}44` },
    }
    const m = variants[v] || variants.primary
    return { display: "inline-flex", alignItems: "center", gap: 6, padding: "9px 14px", borderRadius: 10, border: m.border || "none", cursor: "pointer", fontSize: 13, fontWeight: 600, background: m.background, color: m.color, boxShadow: m.boxShadow || "none", transition: "opacity 0.15s", letterSpacing: "-0.01em", whiteSpace: "nowrap", fontFamily: "inherit" }
  },
  input: { width: "100%", padding: "11px 13px", borderRadius: 10, border: `1.5px solid ${C.border}`, fontSize: 14, fontFamily: "inherit", color: C.textPri, background: C.surfaceHi, outline: "none", boxSizing: "border-box", transition: "border-color 0.15s" },
  select: { padding: "11px 13px", borderRadius: 10, border: `1.5px solid ${C.border}`, fontSize: 14, fontFamily: "inherit", color: C.textPri, background: C.surfaceHi, outline: "none", cursor: "pointer", width: "100%" },
  label: { fontSize: 11, color: C.textDim, textTransform: "uppercase", letterSpacing: "0.08em", display: "block", marginBottom: 6, fontWeight: 600 },
  badge: (color = "rose") => {
    const map = { rose: { bg: "rgba(192,132,168,0.18)", text: C.accent }, gold: { bg: C.goldDim, text: C.gold }, sage: { bg: C.sageDim, text: C.sage }, danger: { bg: C.dangerDim, text: C.danger }, neutral: { bg: C.surfaceHi, text: C.textSec }, purple: { bg: C.purpleDim, text: C.purple } }
    const { bg, text } = map[color] || map.neutral
    return { display: "inline-flex", alignItems: "center", padding: "3px 9px", borderRadius: 20, background: bg, color: text, fontSize: 11, fontWeight: 700, letterSpacing: "0.04em", whiteSpace: "nowrap" }
  },
  modal: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 999, padding: 0 },
  modalBox: { background: C.surface, borderRadius: "22px 22px 0 0", padding: "20px 22px 40px", maxWidth: 600, width: "100%", boxShadow: "0 -8px 60px rgba(0,0,0,0.7)", maxHeight: "92vh", overflowY: "auto", border: `1px solid ${C.border}`, borderBottom: "none" },
  modalHandle: { width: 36, height: 4, background: C.border, borderRadius: 4, margin: "0 auto 18px", display: "block" },
  modalTitle: { fontSize: 17, marginBottom: 16, color: C.textPri, fontWeight: 700, letterSpacing: "-0.02em" },
  formGroup: { marginBottom: 14 },
  row: { display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" },
  col: (f = 1) => ({ flex: f, minWidth: 0 }),
  checkCircle: (done) => ({ width: 22, height: 22, borderRadius: 7, border: `2px solid ${done ? C.sage : C.border}`, background: done ? C.sage : "transparent", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "all 0.2s" }),
  progressBar: { height: 8, borderRadius: 4, background: C.surfaceHi, overflow: "hidden", border: `1px solid ${C.border}` },
  progressFill: (pct, color = C.sage) => ({ height: "100%", width: `${Math.min(pct, 100)}%`, background: color, borderRadius: 4, transition: "width 0.5s cubic-bezier(.4,0,.2,1)" }),
  row2: (i) => ({ display: "flex", alignItems: "center", padding: "11px 14px", background: i % 2 === 0 ? C.surface : C.surfaceHi, borderRadius: 10, marginBottom: 3, gap: 10 }),
  emptyState: { textAlign: "center", padding: "44px 20px", color: C.textDim },
  syncDot: (ok) => ({ width: 7, height: 7, borderRadius: "50%", background: ok ? C.sage : C.gold, flexShrink: 0 }),
}

// ─── RESPONSIVE HOOK ──────────────────────────────────────────────────────────
function useIsMobile() {
  const [mobile, setMobile] = useState(window.innerWidth < 768)
  useEffect(() => {
    const fn = () => setMobile(window.innerWidth < 768)
    window.addEventListener("resize", fn)
    return () => window.removeEventListener("resize", fn)
  }, [])
  return mobile
}

// ─── BODY SCROLL LOCK (for modals) ────────────────────────────────────────────
function useLockBodyScroll(active) {
  useEffect(() => {
    if (!active) return
    const prev = document.body.style.overflow
    document.body.style.overflow = "hidden"
    return () => { document.body.style.overflow = prev }
  }, [active])
}

// ─── SUPABASE SHARED STATE ────────────────────────────────────────────────────
function useSharedState(key, def) {
  const [val, setVal] = useState(def)
  const [synced, setSynced] = useState(false)
  const [syncing, setSyncing] = useState(false)
  const latestVal = useRef(val)
  latestVal.current = val

  // Initial load
  useEffect(() => {
    dbGet(key).then(data => {
      if (data !== null) setVal(data)
      setSynced(true)
    }).catch(() => setSynced(true))
  }, [key])

  // Realtime subscription
  useEffect(() => {
    const unsub = dbSubscribe(key, (newVal) => {
      setVal(newVal)
    })
    return unsub
  }, [key])

  const update = useCallback((nv) => {
    const resolved = typeof nv === "function" ? nv(latestVal.current) : nv
    setVal(resolved)
    setSyncing(true)
    dbSet(key, resolved).finally(() => setSyncing(false))
  }, [key])

  return [val, update, synced, syncing]
}

// ─── GUESTS MODULE ────────────────────────────────────────────────────────────
const GUESTS_PAGE_SIZE = 15

function GuestsModule() {
  const [guests, setGuests, synced] = useSharedState("wedding_guests", [])
  const [groups, setGroups] = useSharedState("wedding_groups", [])
  const [, setTables] = useSharedState("wedding_tables", [])
  const [search, setSearch] = useState("")
  const [filter, setFilter] = useState("all")
  const [page, setPage] = useState(1)
  const [modal, setModal] = useState(null)
  const [groupModal, setGroupModal] = useState(null)
  const [groupForm, setGroupForm] = useState({ name: "" })
  const [groupViewId, setGroupViewId] = useState(null)
  const [genMsg, setGenMsg] = useState("")

  const emptyGuest = { name: "", side: "Mireasă", rsvp: "în așteptare", dietary: "", phone: "", groupId: "", companions: [] }
  const [form, setForm] = useState(emptyGuest)
  const [compForm, setCompForm] = useState({ name: "", type: "adult" })
  const [showCompForm, setShowCompForm] = useState(false)

  useLockBodyScroll(Boolean(modal) || Boolean(groupModal) || Boolean(groupViewId))
  useEffect(() => { setPage(1) }, [search, filter])

  const openAdd = () => { setForm(emptyGuest); setShowCompForm(false); setModal("add") }
  const openEdit = (g) => { setForm({ ...emptyGuest, ...g, companions: g.companions || [] }); setShowCompForm(false); setModal(g) }
  const openDuplicate = (g) => {
    setForm({ ...emptyGuest, ...g, name: "", companions: (g.companions || []).map(c => ({ ...c, id: Date.now() + Math.random() })) })
    setShowCompForm(false)
    setModal("add")
  }

  const addCompanion = () => {
    if (!compForm.name.trim()) return
    setForm(p => ({ ...p, companions: [...(p.companions || []), { id: Date.now(), ...compForm }] }))
    setCompForm({ name: "", type: "adult" })
    setShowCompForm(false)
  }
  const removeCompanion = (cid) => setForm(p => ({ ...p, companions: p.companions.filter(c => c.id !== cid) }))

  const save = () => {
    if (!form.name.trim()) return
    if (modal === "add") setGuests(prev => [...prev, { ...form, id: Date.now() }])
    else setGuests(prev => prev.map(g => g.id === modal.id ? { ...modal, ...form } : g))
    setModal(null)
  }
  const remove = (id) => { if (window.confirm("Ștergi invitatul?")) setGuests(prev => prev.filter(g => g.id !== id)) }

  const openAddGroup = () => { setGroupForm({ name: "" }); setGroupModal("add") }
  const openEditGroup = (g) => { setGroupForm({ name: g.name }); setGroupModal(g) }
  const openGroupView = (id) => { setGroupViewId(id); setGenMsg("") }
  const closeGroupView = () => { setGroupViewId(null); setGenMsg("") }
  const saveGroup = () => {
    if (!groupForm.name.trim()) return
    if (groupModal === "add") setGroups(prev => [...prev, { id: Date.now(), name: groupForm.name }])
    else setGroups(prev => prev.map(g => g.id === groupModal.id ? { ...g, name: groupForm.name } : g))
    setGroupForm({ name: "" }); setGroupModal(null)
  }
  const removeGroup = (id) => {
    setGroups(prev => prev.filter(g => g.id !== id))
    setGuests(prev => prev.map(g => g.groupId === id ? { ...g, groupId: "" } : g))
    if (groupViewId === id) closeGroupView()
  }

  const groupMembers = (id) => guests.filter(g => String(g.groupId) === String(id))
  const groupPeopleCount = (id) => groupMembers(id).reduce((a, g) => a + 1 + (g.companions?.length || 0), 0)

  const generateTableFromGroup = (group) => {
    const members = groupMembers(group.id)
    const seats = members.flatMap(g => [g.id, ...(g.companions || []).map(c => `${g.id}_${c.id}`)])
    if (seats.length === 0) return
    setTables(prev => [...prev, { id: Date.now(), name: group.name, capacity: seats.length, seats }])
    setGenMsg(`Masă „${group.name}” creată cu ${seats.length} locuri — o găsești în Plan mese.`)
  }

  const exportGuests = () => {
    const gName = (id) => groups.find(g => String(g.id) === String(id))?.name || ""
    const rows = guests.map(g => ({ "Nume": g.name, "Partea": g.side, "RSVP": g.rsvp, "Grup": gName(g.groupId), "Telefon": g.phone || "", "Restricții": g.dietary || "", "Însoțitori": (g.companions || []).length, "Total persoane": 1 + (g.companions || []).length }))
    const comps = guests.flatMap(g => (g.companions || []).map(c => ({ "Invitat principal": g.name, "Însoțitor": c.name, "Tip": c.type })))
    exportToExcel([{ name: "Invitați", data: rows.length ? rows : [{ Info: "Niciun invitat" }] }, { name: "Însoțitori", data: comps.length ? comps : [{ Info: "Niciun însoțitor" }] }], "invitatii-nunta.xlsx")
  }

  const totalPeople = guests.reduce((a, g) => a + 1 + (g.companions?.length || 0), 0)
  const confirmed = guests.filter(g => g.rsvp === "confirmat")
  const totalConfirmed = confirmed.reduce((a, g) => a + 1 + (g.companions?.length || 0), 0)

  const filtered = guests.filter(g => {
    const ms = g.name.toLowerCase().includes(search.toLowerCase())
    const mf = filter === "all" || g.rsvp === filter || g.side === filter || String(g.groupId) === String(filter)
    return ms && mf
  })
  const pageCount = Math.max(1, Math.ceil(filtered.length / GUESTS_PAGE_SIZE))
  const pageSafe = Math.min(page, pageCount)
  const paged = filtered.slice((pageSafe - 1) * GUESTS_PAGE_SIZE, pageSafe * GUESTS_PAGE_SIZE)
  const rsvpColor = (r) => r === "confirmat" ? "sage" : r === "refuzat" ? "danger" : "neutral"
  const groupName = (id) => groups.find(g => String(g.id) === String(id))?.name
  const viewedGroup = groupViewId ? groups.find(g => g.id === groupViewId) : null

  return (
    <div className="fadeup">
      <div style={S.statsRow}>
        {[{ num: guests.length, label: "Invitații" }, { num: totalPeople, label: "Total pers." }, { num: totalConfirmed, label: "Confirmați" }, { num: guests.filter(g => g.rsvp === "în așteptare").length, label: "În așteptare" }].map(s => (
          <div key={s.label} style={S.statCard}><div style={S.statNum}>{s.num}</div><div style={S.statLabel}>{s.label}</div></div>
        ))}
      </div>

      {/* Groups */}
      <div style={S.card}>
        <div style={S.cardHeader}>
          <span style={S.cardTitle}>Grupuri</span>
          <button style={S.btn("purple")} onClick={openAddGroup}><PlusIcon /> Grup nou</button>
        </div>
        <div style={{ padding: "12px 20px", display: "flex", flexWrap: "wrap", gap: 8 }}>
          {groups.length === 0 && <span style={{ fontSize: 13, color: C.textDim }}>Niciun grup.</span>}
          {groups.map(g => (
            <div key={g.id} style={{ display: "inline-flex", alignItems: "center", gap: 4, background: C.purpleDim, borderRadius: 20, padding: "4px 6px 4px 12px" }}>
              <button style={{ background: "none", border: "none", cursor: "pointer", padding: 0, display: "flex", alignItems: "center", gap: 6 }} onClick={() => openGroupView(g.id)}>
                <span style={{ fontSize: 12, color: C.purple, fontWeight: 700 }}>{g.name}</span>
                <span style={{ fontSize: 11, color: C.textDim }}>({guests.filter(gu => String(gu.groupId) === String(g.id)).length})</span>
              </button>
              <button style={{ background: "none", border: "none", cursor: "pointer", color: C.textDim, padding: "0 2px", lineHeight: 1, display: "flex" }} onClick={() => openEditGroup(g)}><EditIcon size={12} /></button>
              <button style={{ background: "none", border: "none", cursor: "pointer", color: C.textDim, padding: "0 2px", lineHeight: 1, display: "flex" }} onClick={() => removeGroup(g.id)}><XIcon size={12} /></button>
            </div>
          ))}
        </div>
      </div>

      {/* Guest list */}
      <div style={S.card}>
        <div style={S.cardHeader}>
          <span style={S.cardTitle}>Lista invitați</span>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={S.btn("gold")} onClick={exportGuests}><DownloadIcon /> Excel</button>
            <button style={S.btn("primary")} onClick={openAdd}><PlusIcon /> Adaugă</button>
          </div>
        </div>
        <div style={{ padding: "12px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 10, flexWrap: "wrap" }}>
          <div style={{ flex: 2, minWidth: 160, position: "relative" }}>
            <div style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: C.textDim }}><SearchIcon /></div>
            <input style={{ ...S.input, paddingLeft: 36 }} placeholder="Caută..." value={search} onChange={e => setSearch(e.target.value)} />
          </div>
          <select style={{ ...S.select, flex: 1, minWidth: 140 }} value={filter} onChange={e => setFilter(e.target.value)}>
            <option value="all">Toți</option>
            <option value="confirmat">Confirmați</option>
            <option value="în așteptare">În așteptare</option>
            <option value="refuzat">Au refuzat</option>
            <option value="Mireasă">Partea miresei</option>
            <option value="Mire">Partea mirelui</option>
            {groups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}
          </select>
        </div>
        <div style={{ padding: "10px 20px" }}>
          {!synced && <div style={S.emptyState}><div className="spin" style={{ width: 24, height: 24, border: `2px solid ${C.border}`, borderTopColor: C.accent, borderRadius: "50%", margin: "0 auto 12px" }} /></div>}
          {synced && filtered.length === 0 && <div style={S.emptyState}><div style={{ fontSize: 28, marginBottom: 8, opacity: 0.3 }}>💌</div><div style={{ fontSize: 13 }}>{guests.length === 0 ? "Adaugă primul invitat." : "Niciun rezultat."}</div></div>}
          {synced && paged.map((g, i) => (
            <div key={g.id} style={{ ...S.row2(i), flexDirection: "column", alignItems: "stretch", gap: 6 }}>
              <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{g.name}</div>
                  {(g.companions || []).length > 0 && <div style={{ fontSize: 11, color: C.textDim, marginTop: 2 }}>+ {g.companions.map(c => `${c.name} (${c.type})`).join(", ")}</div>}
                  {g.dietary && <div style={{ fontSize: 11, color: C.textDim }}>{g.dietary}</div>}
                </div>
                <div style={{ display: "flex", gap: 2, flexShrink: 0 }}>
                  <button style={{ ...S.btn("ghost"), padding: "4px 6px" }} title="Duplică" onClick={() => openDuplicate(g)}><CopyIcon /></button>
                  <button style={{ ...S.btn("ghost"), padding: "4px 6px" }} onClick={() => openEdit(g)}><EditIcon /></button>
                  <button style={{ ...S.btn("ghost"), padding: "4px 6px", color: C.danger }} onClick={() => remove(g.id)}><TrashIcon /></button>
                </div>
              </div>
              <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center" }}>
                {groupName(g.groupId) && <span style={S.badge("purple")}>{groupName(g.groupId)}</span>}
                <span style={S.badge(g.side === "Mireasă" ? "rose" : "gold")}>{g.side}</span>
                <span style={S.badge(rsvpColor(g.rsvp))}>{g.rsvp}</span>
                {(g.companions || []).length > 0 && <span style={S.badge("neutral")}>{1 + g.companions.length} pers.</span>}
              </div>
            </div>
          ))}
          {synced && filtered.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, marginTop: 14, paddingTop: 12, borderTop: `1px solid ${C.border}` }}>
              <button style={S.btn("outline")} disabled={pageSafe <= 1} onClick={() => setPage(p => Math.max(1, p - 1))}>‹ Anterior</button>
              <span style={{ fontSize: 12, color: C.textDim }}>Pagina {pageSafe} / {pageCount}</span>
              <button style={S.btn("outline")} disabled={pageSafe >= pageCount} onClick={() => setPage(p => Math.min(pageCount, p + 1))}>Următor ›</button>
            </div>
          )}
        </div>
      </div>

      {/* Guest Modal */}
      {modal && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setModal(null) }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div style={S.modalTitle}>{modal === "add" ? "Invitat nou" : "Editează invitat"}</div>
              <button style={{ ...S.btn("ghost"), padding: 4 }} onClick={() => setModal(null)}><XIcon /></button>
            </div>
            <div style={S.formGroup}><label style={S.label}>Nume complet *</label><input style={S.input} placeholder="ex: Ion și Maria Popescu" value={form.name} onChange={e => setForm(p => ({ ...p, name: e.target.value }))} /></div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Partea</label><select style={S.select} value={form.side} onChange={e => setForm(p => ({ ...p, side: e.target.value }))}><option>Mireasă</option><option>Mire</option></select></div>
              <div style={S.col(1)}><label style={S.label}>RSVP</label><select style={S.select} value={form.rsvp} onChange={e => setForm(p => ({ ...p, rsvp: e.target.value }))}><option>în așteptare</option><option>confirmat</option><option>refuzat</option></select></div>
            </div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Grup</label><select style={S.select} value={form.groupId} onChange={e => setForm(p => ({ ...p, groupId: e.target.value }))}><option value="">— Fără grup —</option>{groups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}</select></div>
              <div style={S.col(1)}><label style={S.label}>Telefon</label><input style={S.input} placeholder="07xx..." value={form.phone} onChange={e => setForm(p => ({ ...p, phone: e.target.value }))} /></div>
            </div>
            <div style={S.formGroup}><label style={S.label}>Restricții alimentare</label><input style={S.input} placeholder="vegetarian, fără gluten..." value={form.dietary} onChange={e => setForm(p => ({ ...p, dietary: e.target.value }))} /></div>
            <div style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <label style={S.label}>Însoțitori ({(form.companions || []).length})</label>
                <button style={S.btn("ghost")} onClick={() => setShowCompForm(p => !p)}><UserPlusIcon /> Adaugă</button>
              </div>
              {(form.companions || []).map(c => (
                <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", background: C.surfaceHi, borderRadius: 8, marginBottom: 5 }}>
                  <span style={{ flex: 1, fontSize: 13 }}>{c.name}</span>
                  <span style={S.badge(c.type === "adult" ? "neutral" : "gold")}>{c.type}</span>
                  <button style={{ ...S.btn("ghost"), padding: "2px 4px", color: C.danger }} onClick={() => removeCompanion(c.id)}><XIcon size={14} /></button>
                </div>
              ))}
              {showCompForm && (
                <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "flex-end" }}>
                  <div style={S.col(2)}><input style={S.input} placeholder="Numele" value={compForm.name} onChange={e => setCompForm(p => ({ ...p, name: e.target.value }))} /></div>
                  <select style={{ ...S.select, flex: 1 }} value={compForm.type} onChange={e => setCompForm(p => ({ ...p, type: e.target.value }))}><option value="adult">Adult</option><option value="copil">Copil</option></select>
                  <button style={S.btn("sage")} onClick={addCompanion}>OK</button>
                </div>
              )}
            </div>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={() => setModal(null)}>Anulează</button>
              <button style={S.btn("primary")} onClick={save}>Salvează</button>
            </div>
          </div>
        </div>
      )}

      {/* Group Modal (add / edit) */}
      {groupModal && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setGroupModal(null) }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={S.modalTitle}>{groupModal === "add" ? "Grup nou" : "Editează grup"}</div>
            <div style={S.formGroup}><label style={S.label}>Numele grupului</label><input style={S.input} placeholder="ex: Prieteni facultate" value={groupForm.name} onChange={e => setGroupForm({ name: e.target.value })} /></div>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={() => setGroupModal(null)}>Anulează</button>
              <button style={S.btn("purple")} onClick={saveGroup}>{groupModal === "add" ? "Creează" : "Salvează"}</button>
            </div>
          </div>
        </div>
      )}

      {/* Group members view */}
      {viewedGroup && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) closeGroupView() }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <div style={S.modalTitle}>{viewedGroup.name}</div>
              <button style={{ ...S.btn("ghost"), padding: 4 }} onClick={closeGroupView}><XIcon /></button>
            </div>
            <div style={{ fontSize: 12, color: C.textDim, marginBottom: 14 }}>{groupPeopleCount(viewedGroup.id)} persoane în total</div>
            <div style={{ maxHeight: 320, overflowY: "auto", marginBottom: 16 }}>
              {groupMembers(viewedGroup.id).length === 0 && <div style={{ fontSize: 13, color: C.textDim }}>Niciun invitat în acest grup încă.</div>}
              {groupMembers(viewedGroup.id).map(g => (
                <div key={g.id} style={{ padding: "8px 4px", borderBottom: `1px solid ${C.border}` }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                    <span style={{ fontSize: 13, fontWeight: 600 }}>{g.name}</span>
                    <span style={S.badge(rsvpColor(g.rsvp))}>{g.rsvp}</span>
                  </div>
                  {(g.companions || []).length > 0 && <div style={{ fontSize: 11, color: C.textDim, marginTop: 2 }}>+ {g.companions.map(c => c.name).join(", ")}</div>}
                </div>
              ))}
            </div>
            {genMsg && <div style={{ fontSize: 12, color: C.sage, marginBottom: 12 }}>{genMsg}</div>}
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={closeGroupView}>Închide</button>
              <button style={S.btn("sage")} disabled={groupMembers(viewedGroup.id).length === 0} onClick={() => generateTableFromGroup(viewedGroup)}><TableIcon /> Generează masă</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── TODO MODULE ──────────────────────────────────────────────────────────────
const TODO_CATS = ["Venue", "Catering", "Foto/Video", "Muzică", "Flori & Decor", "Haine", "Lună de miere", "Invitații", "Altele"]
const TIMEFRAMES = ["12+ luni", "9-12 luni", "6-9 luni", "3-6 luni", "1-3 luni", "< 1 lună", "Ziua nunții"]

function TodoModule() {
  const [tasks, setTasks, synced] = useSharedState("wedding_todos", [])
  const [modal, setModal] = useState(null)
  const [filterCat, setFilterCat] = useState("all")
  const [form, setForm] = useState({ title: "", category: "Venue", timeframe: "6-9 luni", priority: "normal", notes: "" })

  useLockBodyScroll(Boolean(modal))

  const openAdd = () => { setForm({ title: "", category: "Venue", timeframe: "6-9 luni", priority: "normal", notes: "" }); setModal("add") }
  const openEdit = (t) => { setForm({ ...t }); setModal(t) }
  const save = () => {
    if (!form.title.trim()) return
    if (modal === "add") setTasks(prev => [...prev, { ...form, id: Date.now(), done: false }])
    else setTasks(prev => prev.map(t => t.id === modal.id ? { ...modal, ...form } : t))
    setModal(null)
  }
  const toggle = (id) => setTasks(prev => prev.map(t => t.id === id ? { ...t, done: !t.done } : t))
  const remove = (id) => setTasks(prev => prev.filter(t => t.id !== id))

  const exportTodos = () => {
    const rows = tasks.map(t => ({ "Task": t.title, "Categorie": t.category, "Perioadă": t.timeframe, "Prioritate": t.priority, "Status": t.done ? "✓ Completat" : "În așteptare", "Note": t.notes || "" }))
    exportToExcel([{ name: "To-do", data: rows.length ? rows : [{ Info: "Niciun task" }] }], "todo-nunta.xlsx")
  }

  const filtered = filterCat === "all" ? tasks : tasks.filter(t => t.category === filterCat)
  const grouped = TIMEFRAMES.reduce((acc, tf) => { const its = filtered.filter(t => t.timeframe === tf); if (its.length) acc[tf] = its; return acc }, {})
  const done = tasks.filter(t => t.done).length
  const pct = tasks.length ? Math.round(done / tasks.length * 100) : 0
  const pColor = (p) => p === "urgent" ? "danger" : p === "important" ? "gold" : "neutral"

  return (
    <div className="fadeup">
      <div style={S.card}>
        <div style={S.cardHeader}>
          <span style={S.cardTitle}>Progres</span>
          <span style={{ fontSize: 12, color: C.textSec }}>{done}/{tasks.length} completate</span>
        </div>
        <div style={{ padding: "12px 20px" }}>
          <div style={S.progressBar}><div style={S.progressFill(pct, C.accent)} /></div>
          <div style={{ marginTop: 6, textAlign: "right", fontSize: 12, color: C.accent, fontWeight: 700 }}>{pct}%</div>
        </div>
      </div>
      <div style={S.card}>
        <div style={S.cardHeader}>
          <span style={S.cardTitle}>Taskuri</span>
          <div style={{ display: "flex", gap: 8 }}>
            <select style={{ ...S.select, width: "auto", fontSize: 12 }} value={filterCat} onChange={e => setFilterCat(e.target.value)}>
              <option value="all">Toate</option>
              {TODO_CATS.map(c => <option key={c}>{c}</option>)}
            </select>
            <button style={S.btn("gold")} onClick={exportTodos}><DownloadIcon /></button>
            <button style={S.btn("primary")} onClick={openAdd}><PlusIcon /> Nou</button>
          </div>
        </div>
        <div style={{ padding: "10px 20px" }}>
          {!synced && <div style={S.emptyState}><div className="spin" style={{ width: 24, height: 24, border: `2px solid ${C.border}`, borderTopColor: C.accent, borderRadius: "50%", margin: "0 auto" }} /></div>}
          {synced && tasks.length === 0 && <div style={S.emptyState}><div style={{ fontSize: 28, marginBottom: 8, opacity: 0.3 }}>📋</div><div style={{ fontSize: 13 }}>Adaugă primul task.</div></div>}
          {Object.entries(grouped).map(([tf, items]) => (
            <div key={tf} style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: "0.1em", color: C.gold, marginBottom: 6, paddingLeft: 2, fontWeight: 700 }}>◆ {tf}</div>
              {items.map((t, i) => (
                <div key={t.id} style={{ ...S.row2(i), gap: 10 }}>
                  <div style={S.checkCircle(t.done)} onClick={() => toggle(t.id)}>
                    {t.done && <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="3"><polyline points="20 6 9 17 4 12" /></svg>}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 14, fontWeight: 500, textDecoration: t.done ? "line-through" : "none", color: t.done ? C.textDim : C.textPri }}>{t.title}</div>
                    <div style={{ display: "flex", gap: 5, marginTop: 3, flexWrap: "wrap" }}>
                      <span style={S.badge("rose")}>{t.category}</span>
                      {t.priority !== "normal" && <span style={S.badge(pColor(t.priority))}>{t.priority}</span>}
                      {t.notes && <span style={{ fontSize: 11, color: C.textDim }}>{t.notes}</span>}
                    </div>
                  </div>
                  <button style={{ ...S.btn("ghost"), padding: "4px 6px" }} onClick={() => openEdit(t)}><EditIcon /></button>
                  <button style={{ ...S.btn("ghost"), padding: "4px 6px", color: C.danger }} onClick={() => remove(t.id)}><TrashIcon /></button>
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
      {modal && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setModal(null) }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div style={S.modalTitle}>{modal === "add" ? "Task nou" : "Editează task"}</div>
              <button style={{ ...S.btn("ghost"), padding: 4 }} onClick={() => setModal(null)}><XIcon /></button>
            </div>
            <div style={S.formGroup}><label style={S.label}>Descriere *</label><input style={S.input} placeholder="ex: Rezervare restaurant" value={form.title} onChange={e => setForm(p => ({ ...p, title: e.target.value }))} /></div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Categorie</label><select style={S.select} value={form.category} onChange={e => setForm(p => ({ ...p, category: e.target.value }))}>{TODO_CATS.map(c => <option key={c}>{c}</option>)}</select></div>
              <div style={S.col(1)}><label style={S.label}>Perioadă</label><select style={S.select} value={form.timeframe} onChange={e => setForm(p => ({ ...p, timeframe: e.target.value }))}>{TIMEFRAMES.map(m => <option key={m}>{m}</option>)}</select></div>
            </div>
            <div style={S.formGroup}><label style={S.label}>Prioritate</label><select style={S.select} value={form.priority} onChange={e => setForm(p => ({ ...p, priority: e.target.value }))}><option value="normal">Normal</option><option value="important">Important</option><option value="urgent">Urgent</option></select></div>
            <div style={S.formGroup}><label style={S.label}>Note</label><textarea style={{ ...S.input, height: 64, resize: "vertical" }} value={form.notes} onChange={e => setForm(p => ({ ...p, notes: e.target.value }))} /></div>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={() => setModal(null)}>Anulează</button>
              <button style={S.btn("primary")} onClick={save}>Salvează</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── VENDORS MODULE ───────────────────────────────────────────────────────────
const VENDOR_CATS = ["Sală / Restaurant", "Catering", "Foto & Video", "Muzică", "Flori & Decor", "Oficiant / Formalități", "Transport", "Frumusețe", "Altele"]
const VENDOR_STATUS = ["de contactat", "contactat", "ofertă primită", "rezervat", "confirmat"]

function VendorsModule() {
  const [vendors, setVendors, synced] = useSharedState("wedding_vendors", [])
  const [budgetItems] = useSharedState("wedding_budget", [])
  const [modal, setModal] = useState(null)
  const emptyVendor = { name: "", category: "Sală / Restaurant", contactPerson: "", phone: "", email: "", status: "de contactat", notes: "", budgetItemId: "" }
  const [form, setForm] = useState(emptyVendor)

  useLockBodyScroll(Boolean(modal))

  const openAdd = () => { setForm(emptyVendor); setModal("add") }
  const openEdit = (v) => { setForm({ ...emptyVendor, ...v }); setModal(v) }
  const save = () => {
    if (!form.name.trim()) return
    if (modal === "add") setVendors(prev => [...prev, { ...form, id: Date.now() }])
    else setVendors(prev => prev.map(v => v.id === modal.id ? { ...modal, ...form } : v))
    setModal(null)
  }
  const remove = (id) => { if (window.confirm("Ștergi furnizorul?")) setVendors(prev => prev.filter(v => v.id !== id)) }

  const budgetItemLabel = (id) => {
    const it = budgetItems.find(b => String(b.id) === String(id))
    if (!it) return null
    const fmt = new Intl.NumberFormat("ro-RO").format(Math.round(parseFloat(it.estimated) || 0))
    return `${it.name} — ${fmt} RON`
  }

  const exportVendors = () => {
    const rows = vendors.map(v => ({ "Nume": v.name, "Categorie": v.category, "Persoană contact": v.contactPerson || "", "Telefon": v.phone || "", "Email": v.email || "", "Status": v.status, "Cheltuială asociată": budgetItemLabel(v.budgetItemId) || "", "Note": v.notes || "" }))
    exportToExcel([{ name: "Furnizori", data: rows.length ? rows : [{ Info: "Niciun furnizor" }] }], "furnizori-nunta.xlsx")
  }

  const statusColor = (s) => s === "confirmat" ? "sage" : s === "rezervat" ? "purple" : s === "ofertă primită" ? "gold" : s === "contactat" ? "rose" : "neutral"
  const grouped = VENDOR_CATS.reduce((acc, cat) => { const its = vendors.filter(v => v.category === cat); if (its.length) acc[cat] = its; return acc }, {})

  return (
    <div className="fadeup">
      <div style={S.statsRow}>
        {[{ num: vendors.length, label: "Furnizori" }, { num: vendors.filter(v => v.status === "confirmat").length, label: "Confirmați" }, { num: vendors.filter(v => v.status === "rezervat").length, label: "Rezervați" }, { num: vendors.filter(v => v.status === "de contactat").length, label: "De contactat" }].map(s => (
          <div key={s.label} style={S.statCard}><div style={S.statNum}>{s.num}</div><div style={S.statLabel}>{s.label}</div></div>
        ))}
      </div>

      <div style={S.card}>
        <div style={S.cardHeader}>
          <span style={S.cardTitle}>Furnizori</span>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={S.btn("gold")} onClick={exportVendors}><DownloadIcon /></button>
            <button style={S.btn("primary")} onClick={openAdd}><PlusIcon /> Adaugă</button>
          </div>
        </div>
        <div style={{ padding: "10px 20px" }}>
          {!synced && <div style={S.emptyState}><div className="spin" style={{ width: 24, height: 24, border: `2px solid ${C.border}`, borderTopColor: C.accent, borderRadius: "50%", margin: "0 auto" }} /></div>}
          {synced && vendors.length === 0 && <div style={S.emptyState}><div style={{ fontSize: 28, marginBottom: 8, opacity: 0.3 }}>🤝</div><div style={{ fontSize: 13 }}>Adaugă primul furnizor.</div></div>}
          {Object.entries(grouped).map(([cat, its]) => (
            <div key={cat} style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: "0.1em", color: C.gold, marginBottom: 6, paddingLeft: 2, fontWeight: 700 }}>◆ {cat}</div>
              {its.map((v, i) => (
                <div key={v.id} style={{ ...S.row2(i), flexDirection: "column", alignItems: "stretch", gap: 6 }}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 14 }}>{v.name}</div>
                      {(v.contactPerson || v.phone) && (
                        <div style={{ fontSize: 11, color: C.textDim, marginTop: 2 }}>
                          {v.contactPerson}{v.contactPerson && v.phone ? " · " : ""}
                          {v.phone && <a href={`tel:${v.phone}`} style={{ color: C.textSec }}>{v.phone}</a>}
                        </div>
                      )}
                    </div>
                    <div style={{ display: "flex", gap: 2, flexShrink: 0 }}>
                      <button style={{ ...S.btn("ghost"), padding: "4px 6px" }} onClick={() => openEdit(v)}><EditIcon /></button>
                      <button style={{ ...S.btn("ghost"), padding: "4px 6px", color: C.danger }} onClick={() => remove(v.id)}><TrashIcon /></button>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center" }}>
                    <span style={S.badge(statusColor(v.status))}>{v.status}</span>
                    {budgetItemLabel(v.budgetItemId) && <span style={S.badge("gold")}>💰 {budgetItemLabel(v.budgetItemId)}</span>}
                    {v.email && <span style={{ fontSize: 11, color: C.textDim }}>{v.email}</span>}
                  </div>
                  {v.notes && <div style={{ fontSize: 11, color: C.textDim }}>{v.notes}</div>}
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>

      {modal && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setModal(null) }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div style={S.modalTitle}>{modal === "add" ? "Furnizor nou" : "Editează furnizor"}</div>
              <button style={{ ...S.btn("ghost"), padding: 4 }} onClick={() => setModal(null)}><XIcon /></button>
            </div>
            <div style={S.formGroup}><label style={S.label}>Nume firmă *</label><input style={S.input} placeholder="ex: Sala Regal" value={form.name} onChange={e => setForm(p => ({ ...p, name: e.target.value }))} /></div>
            <div style={S.formGroup}><label style={S.label}>Categorie</label><select style={S.select} value={form.category} onChange={e => setForm(p => ({ ...p, category: e.target.value }))}>{VENDOR_CATS.map(c => <option key={c}>{c}</option>)}</select></div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Persoană de contact</label><input style={S.input} placeholder="ex: Ana Popescu" value={form.contactPerson} onChange={e => setForm(p => ({ ...p, contactPerson: e.target.value }))} /></div>
              <div style={S.col(1)}><label style={S.label}>Telefon</label><input style={S.input} placeholder="07xx..." value={form.phone} onChange={e => setForm(p => ({ ...p, phone: e.target.value }))} /></div>
            </div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Email</label><input style={S.input} placeholder="contact@..." value={form.email} onChange={e => setForm(p => ({ ...p, email: e.target.value }))} /></div>
              <div style={S.col(1)}><label style={S.label}>Status</label><select style={S.select} value={form.status} onChange={e => setForm(p => ({ ...p, status: e.target.value }))}>{VENDOR_STATUS.map(s => <option key={s}>{s}</option>)}</select></div>
            </div>
            <div style={S.formGroup}>
              <label style={S.label}>Cheltuială asociată (din Buget)</label>
              <select style={S.select} value={form.budgetItemId} onChange={e => setForm(p => ({ ...p, budgetItemId: e.target.value }))}>
                <option value="">— Fără —</option>
                {budgetItems.map(it => <option key={it.id} value={it.id}>{it.name}</option>)}
              </select>
            </div>
            <div style={S.formGroup}><label style={S.label}>Note</label><textarea style={{ ...S.input, height: 60, resize: "vertical" }} value={form.notes} onChange={e => setForm(p => ({ ...p, notes: e.target.value }))} /></div>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={() => setModal(null)}>Anulează</button>
              <button style={S.btn("primary")} onClick={save}>Salvează</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── SEATING MODULE ───────────────────────────────────────────────────────────
const SEAT = 34
const SEAT_GAP = 14
const TABLE_GRID_STEP = 240

function tableDefaultPos(i) {
  return { x: 40 + (i % 4) * TABLE_GRID_STEP, y: 40 + Math.floor(i / 4) * TABLE_GRID_STEP }
}

function normalizeTable(t, i) {
  const shape = t.shape === "rect" ? "rect" : "round"
  const capacity = Math.max(1, t.capacity || 8)
  const rawSeats = Array.isArray(t.seats) ? t.seats : []
  const seats = Array.from({ length: capacity }, (_, idx) => rawSeats[idx] ?? null)
  const pos = tableDefaultPos(i)
  const x = typeof t.x === "number" ? t.x : pos.x
  const y = typeof t.y === "number" ? t.y : pos.y
  return { ...t, shape, capacity, seats, x, y }
}

function getTableLayout(shape, capacity) {
  if (shape === "rect") {
    const top = Math.ceil(capacity / 2)
    const bottom = capacity - top
    const perSide = Math.max(top, bottom, 1)
    const rectW = Math.max(150, perSide * 56)
    const rectH = 86
    const boxW = rectW + SEAT
    const boxH = rectH + 2 * (SEAT + SEAT_GAP)
    const shapeX = (boxW - rectW) / 2
    const shapeY = SEAT + SEAT_GAP
    const seats = []
    for (let i = 0; i < top; i++) seats.push({ x: shapeX + (rectW / top) * (i + 0.5), y: shapeY - SEAT_GAP - SEAT / 2 })
    for (let i = 0; i < bottom; i++) seats.push({ x: shapeX + (rectW / (bottom || 1)) * (i + 0.5), y: shapeY + rectH + SEAT_GAP + SEAT / 2 })
    return { boxW, boxH, shapeBox: { x: shapeX, y: shapeY, w: rectW, h: rectH }, seats }
  }
  const d = Math.min(190, 70 + Math.max(0, capacity - 4) * 8)
  const r = d / 2
  const seatR = r + SEAT_GAP + SEAT / 2
  const boxSize = 2 * (seatR + SEAT / 2)
  const cx = boxSize / 2, cy = boxSize / 2
  const seats = []
  for (let i = 0; i < capacity; i++) {
    const angle = (2 * Math.PI * i) / capacity - Math.PI / 2
    seats.push({ x: cx + seatR * Math.cos(angle), y: cy + seatR * Math.sin(angle) })
  }
  return { boxW: boxSize, boxH: boxSize, shapeBox: { x: cx - r, y: cy - r, w: d, h: d }, seats }
}

function SeatingModule() {
  const [guests] = useSharedState("wedding_guests", [])
  const [tables, setTables, synced] = useSharedState("wedding_tables", [])
  const [modal, setModal] = useState(null)
  const [form, setForm] = useState({ name: "", shape: "round", capacity: 8 })
  const [capacityError, setCapacityError] = useState("")
  const [seatModal, setSeatModal] = useState(null)
  const [livePos, setLivePos] = useState(null)

  useLockBodyScroll(Boolean(modal) || Boolean(seatModal))

  const normTables = tables.map((t, i) => normalizeTable(t, i))

  const updateTable = (id, patch) => {
    setTables(prev => {
      const norm = prev.map((t, i) => normalizeTable(t, i))
      return norm.map(t => t.id === id ? { ...t, ...patch } : t)
    })
  }

  const openAddTable = () => { setForm({ name: "", shape: "round", capacity: 8 }); setCapacityError(""); setModal("add") }
  const openEditTable = (t) => { setForm({ name: t.name, shape: t.shape, capacity: t.capacity }); setCapacityError(""); setModal(t) }

  const addTable = () => {
    if (!form.name.trim()) return
    const capacity = Math.max(1, parseInt(form.capacity) || 8)
    const pos = tableDefaultPos(tables.length)
    const newTable = { id: Date.now(), name: form.name, shape: form.shape, capacity, seats: Array(capacity).fill(null), x: pos.x, y: pos.y }
    setTables(prev => [...prev.map((t, i) => normalizeTable(t, i)), newTable])
    setModal(null)
  }
  const saveEditTable = () => {
    if (!form.name.trim()) return
    const capacity = Math.max(1, parseInt(form.capacity) || 8)
    const t = normTables.find(tt => tt.id === modal.id)
    const occupied = t.seats.filter(Boolean).length
    if (capacity < occupied) { setCapacityError(`Eliberează ${occupied - capacity} scaune înainte să reduci capacitatea.`); return }
    const seats = Array.from({ length: capacity }, (_, idx) => t.seats[idx] ?? null)
    updateTable(modal.id, { name: form.name, shape: form.shape, capacity, seats })
    setModal(null); setCapacityError("")
  }
  const saveTable = () => modal === "add" ? addTable() : saveEditTable()

  const removeTable = (id) => {
    if (!window.confirm("Ștergi masa?")) return
    setTables(prev => prev.filter(t => t.id !== id))
    if (seatModal?.tableId === id) setSeatModal(null)
  }

  const assignSeat = (tableId, seatIndex, personId) => {
    const t = normTables.find(tt => tt.id === tableId)
    if (!t) return
    const seats = t.seats.slice()
    seats[seatIndex] = personId
    updateTable(tableId, { seats })
    setSeatModal(null)
  }
  const freeSeat = (tableId, seatIndex) => {
    const t = normTables.find(tt => tt.id === tableId)
    if (!t) return
    const seats = t.seats.slice()
    seats[seatIndex] = null
    updateTable(tableId, { seats })
    setSeatModal(null)
  }

  const startDrag = (e, table) => {
    e.preventDefault()
    const startClientX = e.clientX, startClientY = e.clientY
    const origX = table.x, origY = table.y
    const onMove = (ev) => {
      const nx = Math.max(0, origX + (ev.clientX - startClientX))
      const ny = Math.max(0, origY + (ev.clientY - startClientY))
      setLivePos({ id: table.id, x: nx, y: ny })
    }
    const onUp = (ev) => {
      window.removeEventListener("pointermove", onMove)
      window.removeEventListener("pointerup", onUp)
      const nx = Math.max(0, origX + (ev.clientX - startClientX))
      const ny = Math.max(0, origY + (ev.clientY - startClientY))
      updateTable(table.id, { x: nx, y: ny })
      setLivePos(null)
    }
    window.addEventListener("pointermove", onMove)
    window.addEventListener("pointerup", onUp)
  }

  const allPeople = guests.flatMap(g => [
    { id: g.id, name: g.name, rsvp: g.rsvp, isComp: false },
    ...(g.companions || []).map(c => ({ id: `${g.id}_${c.id}`, name: `${c.name} (cu ${g.name})`, rsvp: g.rsvp, isComp: true }))
  ])
  const assignedIds = new Set(normTables.flatMap(t => t.seats.filter(Boolean)))
  const confirmed = allPeople.filter(p => p.rsvp === "confirmat")
  const unassigned = confirmed.filter(p => !assignedIds.has(p.id))
  const personName = (id) => allPeople.find(p => p.id === id)?.name || String(id)
  const personRsvp = (id) => allPeople.find(p => p.id === id)?.rsvp

  const canvasSize = normTables.reduce((acc, t) => {
    const { x, y } = livePos?.id === t.id ? livePos : t
    const { boxW, boxH } = getTableLayout(t.shape, t.capacity)
    return { w: Math.max(acc.w, x + boxW + 300), h: Math.max(acc.h, y + boxH + 300) }
  }, { w: 1200, h: 800 })

  return (
    <div className="fadeup">
      <div style={S.statsRow}>
        {[{ num: tables.length, label: "Mese" }, { num: confirmed.length, label: "Confirmați" }, { num: confirmed.length - unassigned.length, label: "Plasați" }, { num: unassigned.length, label: "Neplasați" }].map(s => (
          <div key={s.label} style={S.statCard}><div style={S.statNum}>{s.num}</div><div style={S.statLabel}>{s.label}</div></div>
        ))}
      </div>
      {unassigned.length > 0 && (
        <div style={{ ...S.card, border: `1px solid ${C.gold}44` }}>
          <div style={S.cardHeader}><span style={{ ...S.cardTitle, color: C.gold }}>⚠️ Neplasați ({unassigned.length})</span></div>
          <div style={{ padding: "10px 20px", display: "flex", flexWrap: "wrap", gap: 6 }}>
            {unassigned.map(p => <span key={p.id} style={S.badge("gold")}>{p.name}</span>)}
          </div>
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <button style={S.btn("primary")} onClick={openAddTable}><PlusIcon /> Masă nouă</button>
      </div>
      {!synced && <div style={S.emptyState}><div className="spin" style={{ width: 24, height: 24, border: `2px solid ${C.border}`, borderTopColor: C.accent, borderRadius: "50%", margin: "0 auto" }} /></div>}
      {synced && tables.length === 0 && <div style={{ ...S.card, ...S.emptyState }}><div style={{ fontSize: 28, marginBottom: 8, opacity: 0.3 }}>🪑</div><div style={{ fontSize: 13 }}>Adaugă mese și plasează invitații.</div></div>}
      {synced && tables.length > 0 && (
        <div style={{ ...S.card, padding: 0 }}>
          <div style={S.cardHeader}>
            <span style={S.cardTitle}>Sala</span>
            <span style={{ fontSize: 11, color: C.textDim }}>Trage din mânerul mesei ca s-o muți</span>
          </div>
          <div style={{ overflow: "auto", maxHeight: "70vh", background: `radial-gradient(circle, ${C.border} 1px, transparent 1px)`, backgroundSize: "24px 24px" }}>
            <div style={{ position: "relative", width: canvasSize.w, height: canvasSize.h }}>
              {normTables.map(table => {
                const pos = livePos?.id === table.id ? livePos : table
                const { boxW, boxH, shapeBox, seats } = getTableLayout(table.shape, table.capacity)
                const occupied = table.seats.filter(Boolean).length
                return (
                  <div key={table.id} style={{ position: "absolute", left: pos.x, top: pos.y }}>
                    <div
                      onPointerDown={e => startDrag(e, table)}
                      style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8, cursor: "grab", touchAction: "none", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 20, padding: "4px 6px 4px 12px", width: "fit-content" }}
                    >
                      <span style={{ fontSize: 12, fontWeight: 700 }}>{table.name}</span>
                      <span style={{ fontSize: 11, color: C.textDim }}>{occupied}/{table.capacity}</span>
                      <button onPointerDown={e => e.stopPropagation()} onClick={() => openEditTable(table)} style={{ ...S.btn("ghost"), padding: "2px 4px" }}><EditIcon size={13} /></button>
                      <button onPointerDown={e => e.stopPropagation()} onClick={() => removeTable(table.id)} style={{ ...S.btn("ghost"), padding: "2px 4px", color: C.danger }}><XIcon size={13} /></button>
                    </div>
                    <div style={{ position: "relative", width: boxW, height: boxH }}>
                      <div style={{
                        position: "absolute", left: shapeBox.x, top: shapeBox.y, width: shapeBox.w, height: shapeBox.h,
                        borderRadius: table.shape === "rect" ? 16 : "50%",
                        background: C.surfaceHi, border: `1.5px solid ${C.borderHi}`
                      }} />
                      {seats.map((s, idx) => {
                        const personId = table.seats[idx]
                        const filled = Boolean(personId)
                        return (
                          <button
                            key={idx}
                            title={filled ? personName(personId) : "Loc liber"}
                            onClick={() => setSeatModal({ tableId: table.id, seatIndex: idx })}
                            style={{
                              position: "absolute", left: s.x - SEAT / 2, top: s.y - SEAT / 2, width: SEAT, height: SEAT, borderRadius: "50%",
                              border: `1.5px solid ${filled ? C.accent : C.border}`,
                              background: filled ? C.accentGrad : C.surfaceHi,
                              color: filled ? "#fff" : C.textDim,
                              fontSize: 10, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
                              padding: 0, lineHeight: 1, overflow: "hidden"
                            }}
                          >
                            {filled ? personName(personId).slice(0, 2).toUpperCase() : "+"}
                          </button>
                        )
                      })}
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        </div>
      )}
      {modal && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setModal(null) }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={S.modalTitle}>{modal === "add" ? "Masă nouă" : "Editează masă"}</div>
            <div style={S.formGroup}><label style={S.label}>Numele mesei</label><input style={S.input} placeholder="ex: Masa mirilor, Masa 1..." value={form.name} onChange={e => setForm(p => ({ ...p, name: e.target.value }))} /></div>
            <div style={S.formGroup}>
              <label style={S.label}>Tip masă</label>
              <div style={{ display: "flex", gap: 8 }}>
                <button style={{ ...S.btn(form.shape === "round" ? "primary" : "outline"), flex: 1, justifyContent: "center" }} onClick={() => setForm(p => ({ ...p, shape: "round" }))}>◯ Rotundă</button>
                <button style={{ ...S.btn(form.shape === "rect" ? "primary" : "outline"), flex: 1, justifyContent: "center" }} onClick={() => setForm(p => ({ ...p, shape: "rect" }))}>▭ Dreptunghiulară</button>
              </div>
            </div>
            <div style={S.formGroup}><label style={S.label}>Capacitate (locuri)</label><input style={S.input} type="number" min={1} max={30} value={form.capacity} onChange={e => setForm(p => ({ ...p, capacity: parseInt(e.target.value) || 8 }))} /></div>
            {capacityError && <div style={{ fontSize: 12, color: C.danger, marginBottom: 12 }}>{capacityError}</div>}
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={() => setModal(null)}>Anulează</button>
              <button style={S.btn("primary")} onClick={saveTable}>{modal === "add" ? "Creează" : "Salvează"}</button>
            </div>
          </div>
        </div>
      )}
      {seatModal && (() => {
        const t = normTables.find(tt => tt.id === seatModal.tableId)
        if (!t) return null
        const personId = t.seats[seatModal.seatIndex]
        const available = allPeople.filter(p => !assignedIds.has(p.id))
        return (
          <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setSeatModal(null) }}>
            <div style={S.modalBox}>
              <div style={S.modalHandle} />
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div style={S.modalTitle}>{t.name} · scaun {seatModal.seatIndex + 1}</div>
                <button style={{ ...S.btn("ghost"), padding: 4 }} onClick={() => setSeatModal(null)}><XIcon /></button>
              </div>
              {personId ? (
                <>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 20 }}>
                    <span style={{ fontSize: 14, fontWeight: 600 }}>{personName(personId)}</span>
                    <span style={S.badge(personRsvp(personId) === "confirmat" ? "sage" : "neutral")}>{personRsvp(personId)}</span>
                  </div>
                  <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                    <button style={S.btn("outline")} onClick={() => setSeatModal(null)}>Închide</button>
                    <button style={S.btn("danger")} onClick={() => freeSeat(t.id, seatModal.seatIndex)}>Eliberează scaunul</button>
                  </div>
                </>
              ) : (
                <>
                  {available.length === 0 && <div style={{ color: C.textDim, fontSize: 13 }}>Toți sunt plasați.</div>}
                  <div style={{ maxHeight: 360, overflowY: "auto" }}>
                    {available.map(p => (
                      <div key={p.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 4px", borderBottom: `1px solid ${C.border}` }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ fontSize: 13 }}>{p.name}</span>
                          <span style={S.badge(p.rsvp === "confirmat" ? "sage" : "neutral")}>{p.rsvp}</span>
                          {p.isComp && <span style={S.badge("purple")}>însoțitor</span>}
                        </div>
                        <button style={S.btn("sage")} onClick={() => assignSeat(t.id, seatModal.seatIndex, p.id)}>Așază aici</button>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
          </div>
        )
      })()}
    </div>
  )
}

// ─── BUDGET MODULE ────────────────────────────────────────────────────────────
const BUDGET_CATS = ["Venue / Restaurant", "Catering & Băuturi", "Foto & Video", "Muzică", "Flori & Decor", "Haine & Accesorii", "Invitații & Print", "Lună de miere", "Transport", "Altele"]
const PAY_STATUS = ["neplătit", "parțial plătit", "plătit integral"]

function BudgetModule() {
  const [items, setItems, synced] = useSharedState("wedding_budget", [])
  const [totalBudget, setTotalBudget] = useSharedState("wedding_budget_total", 0)
  const [modal, setModal] = useState(null)
  const [editTotal, setEditTotal] = useState(false)
  const [totalInput, setTotalInput] = useState("")
  const [form, setForm] = useState({ name: "", category: "Venue / Restaurant", estimated: 0, paid: 0, vendor: "", paymentStatus: "neplătit", notes: "" })

  useLockBodyScroll(Boolean(modal))

  const openAdd = () => { setForm({ name: "", category: "Venue / Restaurant", estimated: 0, paid: 0, vendor: "", paymentStatus: "neplătit", notes: "" }); setModal("add") }
  const openEdit = (it) => { setForm({ ...it }); setModal(it) }
  const save = () => {
    if (!form.name.trim()) return
    const parsed = { ...form, estimated: parseFloat(form.estimated) || 0, paid: parseFloat(form.paid) || 0 }
    if (modal === "add") setItems(prev => [...prev, { ...parsed, id: Date.now() }])
    else setItems(prev => prev.map(it => it.id === modal.id ? { ...modal, ...parsed } : it))
    setModal(null)
  }
  const remove = (id) => { if (window.confirm("Ștergi cheltuiala?")) setItems(prev => prev.filter(it => it.id !== id)) }

  const exportBudget = () => {
    const rows = items.map(it => ({ "Descriere": it.name, "Categorie": it.category, "Furnizor": it.vendor || "", "Estimat (RON)": parseFloat(it.estimated) || 0, "Plătit (RON)": parseFloat(it.paid) || 0, "Diferență": (parseFloat(it.estimated) || 0) - (parseFloat(it.paid) || 0), "Status": it.paymentStatus, "Note": it.notes || "" }))
    const summary = [{ "": "Buget total", "Valoare (RON)": parseFloat(totalBudget) || 0 }, { "": "Total estimat", "Valoare (RON)": totalEst }, { "": "Total plătit", "Valoare (RON)": totalPaid }, { "": "Rămas de plătit", "Valoare (RON)": remaining }]
    exportToExcel([{ name: "Cheltuieli", data: rows.length ? rows : [{ Info: "Nicio cheltuială" }] }, { name: "Sumar", data: summary }], "buget-nunta.xlsx")
  }

  const totalEst = items.reduce((s, it) => s + (parseFloat(it.estimated) || 0), 0)
  const totalPaid = items.reduce((s, it) => s + (parseFloat(it.paid) || 0), 0)
  const remaining = (parseFloat(totalBudget) || 0) - totalPaid
  const pctUsed = totalBudget > 0 ? Math.round(totalPaid / totalBudget * 100) : 0
  const pctEst = totalBudget > 0 ? Math.round(totalEst / totalBudget * 100) : 0
  const fmt = (n) => new Intl.NumberFormat("ro-RO").format(Math.round(n)) + " RON"
  const payColor = (s) => s === "plătit integral" ? "sage" : s === "parțial plătit" ? "gold" : "neutral"

  const grouped = BUDGET_CATS.reduce((acc, cat) => { const its = items.filter(it => it.category === cat); if (its.length) acc[cat] = its; return acc }, {})

  return (
    <div className="fadeup">
      <div style={S.statsRow}>
        <div style={{ ...S.statCard, gridColumn: "span 2" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <div style={S.statLabel}>Buget total</div>
            <button style={{ ...S.btn("ghost"), fontSize: 11, padding: "2px 6px" }} onClick={() => { setTotalInput(totalBudget); setEditTotal(true) }}><EditIcon size={12} /> Editează</button>
          </div>
          {editTotal ? (
            <div style={{ display: "flex", gap: 8 }}>
              <input style={{ ...S.input, fontSize: 18, fontWeight: 800 }} type="number" value={totalInput} onChange={e => setTotalInput(e.target.value)} />
              <button style={S.btn("sage")} onClick={() => { setTotalBudget(parseFloat(totalInput) || 0); setEditTotal(false) }}>OK</button>
            </div>
          ) : <div style={{ ...S.statNum, textAlign: "left" }}>{fmt(totalBudget)}</div>}
        </div>
        {[{ num: fmt(totalEst), label: `Estimat (${pctEst}%)` }, { num: fmt(totalPaid), label: `Plătit (${pctUsed}%)` }, { num: fmt(remaining), label: remaining < 0 ? "⚠️ Depășit!" : "Rămas" }].map(s => (
          <div key={s.label} style={{ ...S.statCard, borderColor: s.label.includes("Depășit") ? C.danger : C.border }}>
            <div style={{ ...S.statNum, fontSize: 18 }}>{s.num}</div>
            <div style={S.statLabel}>{s.label}</div>
          </div>
        ))}
      </div>

      <div style={S.card}>
        <div style={S.cardHeader}><span style={S.cardTitle}>Utilizare buget</span></div>
        <div style={S.cardBody}>
          {[{ label: `Estimat — ${fmt(totalEst)}`, pct: pctEst, color: C.gold }, { label: `Plătit — ${fmt(totalPaid)}`, pct: pctUsed, color: pctUsed > 100 ? C.danger : C.sage }].map(b => (
            <div key={b.label} style={{ marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: C.textSec, marginBottom: 4 }}><span>{b.label}</span><span style={{ fontWeight: 700 }}>{b.pct}%</span></div>
              <div style={S.progressBar}><div style={S.progressFill(b.pct, b.color)} /></div>
            </div>
          ))}
        </div>
      </div>

      <div style={S.card}>
        <div style={S.cardHeader}>
          <span style={S.cardTitle}>Cheltuieli</span>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={S.btn("gold")} onClick={exportBudget}><DownloadIcon /></button>
            <button style={S.btn("primary")} onClick={openAdd}><PlusIcon /> Adaugă</button>
          </div>
        </div>
        <div style={{ padding: "10px 20px" }}>
          {!synced && <div style={S.emptyState}><div className="spin" style={{ width: 24, height: 24, border: `2px solid ${C.border}`, borderTopColor: C.accent, borderRadius: "50%", margin: "0 auto" }} /></div>}
          {synced && items.length === 0 && <div style={S.emptyState}><div style={{ fontSize: 28, marginBottom: 8, opacity: 0.3 }}>💰</div><div style={{ fontSize: 13 }}>Adaugă prima cheltuială.</div></div>}
          {Object.entries(grouped).map(([cat, its]) => {
            const catEst = its.reduce((s, it) => s + (parseFloat(it.estimated) || 0), 0)
            const catPaid = its.reduce((s, it) => s + (parseFloat(it.paid) || 0), 0)
            return (
              <div key={cat} style={{ marginBottom: 18 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                  <div style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: "0.1em", color: C.gold, fontWeight: 700 }}>◆ {cat}</div>
                  <div style={{ fontSize: 11, color: C.textDim }}>Est: {fmt(catEst)} · Plătit: {fmt(catPaid)}</div>
                </div>
                {its.map((it, i) => (
                  <div key={it.id} style={{ ...S.row2(i), flexWrap: "wrap" }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 14 }}>{it.name}</div>
                      {it.vendor && <div style={{ fontSize: 11, color: C.textDim, marginTop: 1 }}>Furnizor: {it.vendor}</div>}
                    </div>
                    <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                      <div style={{ textAlign: "right" }}>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{fmt(it.estimated)}</div>
                        <div style={{ fontSize: 11, color: C.sage }}>plătit: {fmt(it.paid)}</div>
                      </div>
                      <span style={S.badge(payColor(it.paymentStatus))}>{it.paymentStatus}</span>
                    </div>
                    <button style={{ ...S.btn("ghost"), padding: "4px 6px" }} onClick={() => openEdit(it)}><EditIcon /></button>
                    <button style={{ ...S.btn("ghost"), padding: "4px 6px", color: C.danger }} onClick={() => remove(it.id)}><TrashIcon /></button>
                  </div>
                ))}
              </div>
            )
          })}
        </div>
      </div>

      {modal && (
        <div style={S.modal} onClick={e => { if (e.target === e.currentTarget) setModal(null) }}>
          <div style={S.modalBox}>
            <div style={S.modalHandle} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div style={S.modalTitle}>{modal === "add" ? "Cheltuială nouă" : "Editează"}</div>
              <button style={{ ...S.btn("ghost"), padding: 4 }} onClick={() => setModal(null)}><XIcon /></button>
            </div>
            <div style={S.formGroup}><label style={S.label}>Descriere *</label><input style={S.input} placeholder="ex: Rezervare restaurant" value={form.name} onChange={e => setForm(p => ({ ...p, name: e.target.value }))} /></div>
            <div style={S.formGroup}><label style={S.label}>Categorie</label><select style={S.select} value={form.category} onChange={e => setForm(p => ({ ...p, category: e.target.value }))}>{BUDGET_CATS.map(c => <option key={c}>{c}</option>)}</select></div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Estimat (RON)</label><input style={S.input} type="number" min={0} value={form.estimated} onChange={e => setForm(p => ({ ...p, estimated: e.target.value }))} /></div>
              <div style={S.col(1)}><label style={S.label}>Plătit (RON)</label><input style={S.input} type="number" min={0} value={form.paid} onChange={e => setForm(p => ({ ...p, paid: e.target.value }))} /></div>
            </div>
            <div style={{ ...S.row, marginBottom: 14 }}>
              <div style={S.col(1)}><label style={S.label}>Furnizor</label><input style={S.input} placeholder="Numele furnizorului" value={form.vendor} onChange={e => setForm(p => ({ ...p, vendor: e.target.value }))} /></div>
              <div style={S.col(1)}><label style={S.label}>Status plată</label><select style={S.select} value={form.paymentStatus} onChange={e => setForm(p => ({ ...p, paymentStatus: e.target.value }))}>{PAY_STATUS.map(s => <option key={s}>{s}</option>)}</select></div>
            </div>
            <div style={S.formGroup}><label style={S.label}>Note</label><textarea style={{ ...S.input, height: 60, resize: "vertical" }} value={form.notes} onChange={e => setForm(p => ({ ...p, notes: e.target.value }))} /></div>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button style={S.btn("outline")} onClick={() => setModal(null)}>Anulează</button>
              <button style={S.btn("primary")} onClick={save}>Salvează</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────
const TABS = [
  { id: "guests",  label: "Invitați",  Icon: UsersIcon    },
  { id: "todo",    label: "To-do",     Icon: CheckIcon    },
  { id: "vendors", label: "Furnizori", Icon: BriefcaseIcon },
  { id: "seating", label: "Plan mese", Icon: TableIcon    },
  { id: "budget",  label: "Buget",     Icon: WalletIcon   },
]

export default function App() {
  const [tab, setTab] = useState("guests")
  const [syncing, setSyncing] = useState(false)
  const isMobile = useIsMobile()

  // inject global styles once
  useEffect(() => {
    if (!document.getElementById("wp-styles")) {
      const el = document.createElement("style")
      el.id = "wp-styles"
      el.textContent = globalStyles
      document.head.appendChild(el)
    }
  }, [])

  const activeTab = TABS.find(t => t.id === tab)

  return (
    <div style={S.app}>
      {/* Top bar */}
      <header style={S.topbar}>
        <div style={S.topbarInner}>
          <div style={S.logo}>
            <HeartIcon size={18} />
            <span>Nuntă</span>
          </div>

          {/* Desktop nav */}
          {!isMobile && (
            <nav style={{ display: "flex", gap: 2 }}>
              {TABS.map(t => {
                const a = tab === t.id
                return (
                  <button key={t.id} onClick={() => setTab(t.id)} style={{ display: "flex", alignItems: "center", gap: 7, padding: "8px 16px", borderRadius: 10, border: "none", cursor: "pointer", fontSize: 13, fontWeight: a ? 700 : 400, fontFamily: "inherit", background: a ? "rgba(192,132,168,0.14)" : "transparent", color: a ? C.accent : C.textSec, transition: "all 0.15s" }}>
                    <t.Icon size={15} />{t.label}
                  </button>
                )
              })}
            </nav>
          )}

          {isMobile && <div style={{ fontSize: 15, fontWeight: 700, letterSpacing: "-0.02em" }}>{activeTab?.label}</div>}

          <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11, color: C.textDim }}>
            <div style={S.syncDot(true)} />
            <span>live</span>
          </div>
        </div>
      </header>

      {/* Main content */}
      <main style={S.main(isMobile)} key={tab} className="fadeup">
        {tab === "guests"  && <GuestsModule />}
        {tab === "todo"    && <TodoModule />}
        {tab === "vendors" && <VendorsModule />}
        {tab === "seating" && <SeatingModule />}
        {tab === "budget"  && <BudgetModule />}
      </main>

      {/* Bottom nav — mobile only */}
      {isMobile && (
        <nav style={S.bottomNav}>
          {TABS.map(t => {
            const a = tab === t.id
            return (
              <button key={t.id} style={S.bottomNavBtn(a)} onClick={() => setTab(t.id)}>
                <div style={{ position: "relative" }}>
                  {a && <div style={{ position: "absolute", inset: -7, borderRadius: 12, background: "rgba(192,132,168,0.14)" }} />}
                  <div style={{ position: "relative" }}><t.Icon size={22} /></div>
                </div>
                <span style={S.bottomNavLabel(a)}>{t.label}</span>
                {a && <div style={{ width: 4, height: 4, borderRadius: "50%", background: C.accent, marginTop: 1 }} />}
              </button>
            )
          })}
        </nav>
      )}
    </div>
  )
}
