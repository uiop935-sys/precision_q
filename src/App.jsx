import { useState, useRef, useEffect } from "react";

// ============================================================
// 1. 변환시트 로직 (OP10_프로그램_수정.xlsm 기반)
// ============================================================

const TRANSFORM_RULES = {
  _default: {
    X: { targetAxis: "X", coeff: -1 },
    Y: { targetAxis: null, coeff: 0 },
    Z: { targetAxis: "Y", coeff: 1 },
    D: { targetAxis: null, coeff: 0 },
  },
  바텀페이스: {
    X: { targetAxis: null, coeff: 0 },
    Y: { targetAxis: "Z", coeff: -1 },
    Z: { targetAxis: null, coeff: 0 },
    D: { targetAxis: null, coeff: 0 },
  },
  "베어링 캡 폭": {
    X: { targetAxis: "Y", coeff: 1 },
    Y: { targetAxis: "Z", coeff: 1 },
    Z: { targetAxis: null, coeff: 0 },
    D: { targetAxis: null, coeff: 0 },
  },
  "베어링 캡 깊이": {
    X: { targetAxis: "Y", coeff: -1 },
    Y: { targetAxis: "Z", coeff: 1 },
    Z: { targetAxis: null, coeff: 0 },
    D: { targetAxis: null, coeff: 0 },
  },
  "언더컷(아래)": {
    X: { targetAxis: "X", coeff: -1 },
    Y: { targetAxis: "Y", coeff: -1 },
    Z: { targetAxis: null, coeff: 0 },
    D: { targetAxis: null, coeff: 0 },
  },
  "언더컷(위)": {
    X: { targetAxis: "X", coeff: -1 },
    Y: { targetAxis: "Y", coeff: -1 },
    Z: { targetAxis: null, coeff: 0 },
    D: { targetAxis: null, coeff: 0 },
  },
};

const PROGRAM_MAP_THETA3 = {
  T101: { holes: ["BT15","BT16","BT17","BT18","BT19","BT38","BT23","BT22","BT21","BT20"], coord: "G55", program: "O5001" },
  T105: { holes: ["BT42"], coord: "G55", program: "O5005" },
  T106: { holes: ["BT12"], coord: "G55", program: "O5006" },
  T116: { holes: ["BT24","BT25","BT26","BT27"], coord: "G55", program: "O5016" },
  T119: { holes: ["BT28","BT29","BT30","BT31"], coord: "G55", program: "O5019" },
  T121: { holes: ["BT14"], coord: "G55", program: "O5021" },
  T122: { holes: ["BT37"], coord: "G55", program: "O5022" },
  T124: { holes: ["BT60"], coord: "G55", program: "O5024" },
};

const PROGRAM_MAP_THETA2 = {
  T101: { holes: ["D18","D19","D20","D21","D22","D33","D34","D35","D36","D37"], coord: "G55", program: "O1001" },
  T102: { holes: ["D23","D25","D27","D28","D30","D32"], coord: "G55", program: "O1002" },
  T105: { holes: ["D60"], coord: "G55", program: "O1005" },
  T106: { holes: ["D61"], coord: "G55", program: "O1006" },
  T107: { holes: ["바텀페이스"], coord: "G55", program: "O1007" },
  T108: { holes: ["베어링캡 언더컷(아래)"], coord: "G57", program: "O1008" },
  T109: { holes: ["베어링 캡 폭"], coord: "G55", program: "O1009" },
  T111: { holes: ["베어링 캡 깊이"], coord: "G55", program: "O1011" },
};

const ENGINE_TYPES = [
  { id: "theta2", label: "세타2", models: ["세타2 GDI 2.0", "세타2 GDI 2.4", "세타2 MPI 2.0"] },
  { id: "theta3", label: "세타3", models: ["세타3 BLOCK GDI 2.5", "세타3 BLOCK MPI 2.5"] },
  { id: "nu", label: "누우", models: ["누우 GDI 1.6", "누우 MPI 1.6"] },
  { id: "gamma", label: "감마", models: ["감마 1.4 T-GDI", "감마 1.6 MPI"] },
];


// ============================================================
// 2. 보정값 산출
// ============================================================

function computeCorrections(parsedRows, engineType) {
  const ngRows = parsedRows.filter(r => r.ngOk === "NG");
  const corrections = {};

  ngRows.forEach(row => {
    const rule = TRANSFORM_RULES._default;
    const transform = rule[row.item];
    if (!transform || !transform.targetAxis) return;

    const corrValue = parseFloat((row.deviation * transform.coeff).toFixed(4));
    const pointKey = row.no;

    if (!corrections[pointKey]) {
      corrections[pointKey] = { point: row.no, type: row.type, X: 0, Y: 0, Z: 0, measurements: {} };
    }
    corrections[pointKey][transform.targetAxis] = corrValue;
    corrections[pointKey].measurements[row.item] = {
      actual: row.actual, nominal: row.nominal, deviation: row.deviation,
      corrAxis: transform.targetAxis, corrValue,
    };
  });

  return Object.values(corrections).map((corr, idx) => ({
    ...corr,
    program: `O${5000 + idx + 1}`,
    tool: `T${101 + idx}`,
    coord: "G55",
  }));
}

// ============================================================
// 3. OCR 프롬프트
// ============================================================

// ============================================================
// 4. 색상
// ============================================================

const C = {
  bg: "#0a0a0f", surface: "#12121a", surfaceAlt: "#1a1a28",
  border: "#2a2a3a", borderLight: "#3a3a4f",
  text: "#e8e8f0", textDim: "#8888a0", textMuted: "#555570",
  accent: "#c8102e", accentDark: "#9a0c22", accentGlow: "rgba(200,16,46,0.15)",
  ok: "#22c55e", okDim: "rgba(34,197,94,0.12)",
  ng: "#ef4444", ngDim: "rgba(239,68,68,0.12)",
  blue: "#3b82f6", blueDim: "rgba(59,130,246,0.12)",
  gold: "#f59e0b", goldDim: "rgba(245,158,11,0.12)",
};

// ============================================================
// 5. 앱
// ============================================================

export default function PrecisionQ() {
  const [step, setStep] = useState(0);
  const [engineType, setEngineType] = useState(null);
  const [engineModel, setEngineModel] = useState(null);
  const [equipment, setEquipment] = useState(null);
  const [equipmentGroup, setEquipmentGroup] = useState(null);
  const [imageFile, setImageFile] = useState(null);
  const [imagePreview, setImagePreview] = useState(null);
  const [ocrResult, setOcrResult] = useState(null);
  const [ocrLoading, setOcrLoading] = useState(false);
  const [ocrError, setOcrError] = useState(null);
  const [corrections, setCorrections] = useState(null);
  const [showHistory, setShowHistory] = useState(false);
  const [logs, setLogs] = useState([]);
  const [animateIn, setAnimateIn] = useState(true);
  const fileInputRef = useRef(null);

  // persistent storage (localStorage)
  useEffect(() => {
    try {
      const stored = localStorage.getItem("precision-q-logs");
      if (stored) setLogs(JSON.parse(stored));
    } catch {}
  }, []);

  const saveLogs = (next) => {
    setLogs(next);
    try { localStorage.setItem("precision-q-logs", JSON.stringify(next)); } catch {}
  };

  const DEMO_DATA = {
    "품명": "세타3 BLOCK GDI 2.5",
    rows: [
      { type:"RH BOSS MILLING", no:"E1", item:"X", actual:67.686, nominal:67.55, upTol:0.1, lowTol:-0.1, deviation:0.136, evaluation:"0.036", ngOk:"NG" },
      { type:"RH BOSS MILLING", no:"E1", item:"Y/X", actual:0.088, nominal:0, upTol:0.5, lowTol:-0.5, deviation:0.088, evaluation:"|+", ngOk:"OK" },
      { type:"RH BOSS MILLING", no:"E1", item:"Z/X", actual:-0.006, nominal:0, upTol:0.5, lowTol:-0.5, deviation:-0.006, evaluation:"-|", ngOk:"OK" },
      { type:"RH BOSS MILLING", no:"E2", item:"X", actual:67.686, nominal:67.55, upTol:0.1, lowTol:-0.1, deviation:0.136, evaluation:"0.036", ngOk:"NG" },
      { type:"RH BOSS MILLING", no:"E3", item:"X", actual:39.045, nominal:39, upTol:0.2, lowTol:-0.2, deviation:0.045, evaluation:"|+", ngOk:"OK" },
      { type:"RH BOSS MILLING", no:"E4", item:"X", actual:39.098, nominal:39, upTol:0.2, lowTol:-0.2, deviation:0.098, evaluation:"|++", ngOk:"OK" },
      { type:"RH BOSS MILLING", no:"E8", item:"X", actual:64.104, nominal:64, upTol:0.1, lowTol:-0.1, deviation:0.104, evaluation:"0.004", ngOk:"NG" },
      { type:"RH BOSS MILLING", no:"E9", item:"X", actual:64.103, nominal:64, upTol:0.1, lowTol:-0.1, deviation:0.103, evaluation:"0.003", ngOk:"NG" },
      { type:"RH BOSS MILLING", no:"E11", item:"X", actual:64.103, nominal:64, upTol:0.1, lowTol:-0.1, deviation:0.103, evaluation:"0.003", ngOk:"NG" },
      { type:"RH BOSS MILLING", no:"E10", item:"X", actual:64.068, nominal:64, upTol:0.1, lowTol:-0.1, deviation:0.068, evaluation:"|+++", ngOk:"OK" },
      { type:"RH BOSS MILLING", no:"E12", item:"X", actual:45.568, nominal:45.5, upTol:0.2, lowTol:-0.2, deviation:0.068, evaluation:"|++", ngOk:"OK" },
      { type:"D/SHAFT MTG HOLE DR", no:"E8", item:"Y", actual:35.163, nominal:35, upTol:0.142, lowTol:-0.142, deviation:0.163, evaluation:"0.021", ngOk:"NG" },
      { type:"D/SHAFT MTG HOLE DR", no:"E8", item:"Z", actual:130.024, nominal:130, upTol:0.142, lowTol:-0.142, deviation:0.024, evaluation:"|+", ngOk:"OK" },
      { type:"D/SHAFT MTG HOLE DR", no:"E8", item:"D", actual:8.803, nominal:8.8, upTol:0.1, lowTol:-0.1, deviation:0.003, evaluation:"|+", ngOk:"OK" },
      { type:"D/SHAFT MTG HOLE DR", no:"E9", item:"Y", actual:35.184, nominal:35, upTol:0.142, lowTol:-0.142, deviation:0.184, evaluation:"0.042", ngOk:"NG" },
      { type:"D/SHAFT MTG HOLE DR", no:"E10", item:"Y", actual:55.188, nominal:55, upTol:0.142, lowTol:-0.142, deviation:0.188, evaluation:"0.046", ngOk:"NG" },
      { type:"LH BOSS MILLING", no:"N1", item:"X", actual:-196.026, nominal:-196, upTol:0.2, lowTol:-0.2, deviation:0.026, evaluation:"|+", ngOk:"OK" },
      { type:"LH BOSS MILLING", no:"N12", item:"X", actual:-164.056, nominal:-164, upTol:0.1, lowTol:-0.1, deviation:0.056, evaluation:"|+++", ngOk:"OK" },
    ]
  };

  const goToStep = (s) => { setAnimateIn(false); setTimeout(() => { setStep(s); setAnimateIn(true); }, 150); };

  const handleImageSelect = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImageFile(file);
    const reader = new FileReader();
    reader.onload = (ev) => setImagePreview(ev.target.result);
    reader.readAsDataURL(file);
  };

  const runOCR = async (useDemo = false) => {
    if (useDemo) {
      setOcrLoading(true); setOcrError(null);
      await new Promise(r => setTimeout(r, 1500));
      setOcrResult(DEMO_DATA); setOcrLoading(false); goToStep(2); return;
    }
    if (!imageFile) return;
    setOcrLoading(true); setOcrError(null);
    try {
      const base64 = imagePreview.split(",")[1];
      const res = await fetch("/api/ocr", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ image: base64, mimeType: imageFile.type || "image/png" }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setOcrResult(data);
      goToStep(2);
    } catch (err) {
      console.error(err);
      setOcrError("OCR 오류 — 데모 데이터로 진행하시겠습니까?");
    } finally { setOcrLoading(false); }
  };

  const calculateCorrections = () => {
    if (!ocrResult?.rows) return;
    const corr = computeCorrections(ocrResult.rows, engineType?.id);
    setCorrections(corr);
    const entry = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2,6),
      ts: new Date().toISOString(),
      engine: engineModel, equipment,
      품명: ocrResult?.품명 || engineModel,
      total: ocrResult?.rows?.length || 0,
      ng: ocrResult?.rows?.filter(r => r.ngOk === "NG").length || 0,
      corr: corr.map(c => ({ pt: c.point, tp: c.type, X: c.X, Y: c.Y, Z: c.Z, prg: c.program, tl: c.tool, cd: c.coord })),
      status: "calculated",
    };
    saveLogs([entry, ...logs].slice(0, 100));
    goToStep(3);
  };

  const reset = () => {
    setStep(0); setEngineType(null); setEngineModel(null); setEquipment(null); setEquipmentGroup(null);
    setImageFile(null); setImagePreview(null); setOcrResult(null); setOcrError(null);
    setCorrections(null); setAnimateIn(true);
  };

  const totalRows = ocrResult?.rows?.length || 0;
  const ngCount = ocrResult?.rows?.filter(r => r.ngOk === "NG").length || 0;
  const okCount = totalRows - ngCount;

  // ===================== 분석 이력 화면 =====================
  if (showHistory) {
    return (
      <div style={shell}>
        <header style={hdr}>
          <button onClick={() => setShowHistory(false)} style={{ ...ghostBtn, fontSize: 14 }}>← 돌아가기</button>
          {logs.length > 0 && <button onClick={() => saveLogs([])} style={{ ...ghostBtn, fontSize: 11, border: `1px solid ${C.border}`, padding: "4px 12px", borderRadius: 6 }}>전체 삭제</button>}
        </header>
        <div style={{ padding: 20 }}>
          <h2 style={titleStyle}>분석 이력</h2>
          <p style={{ ...subtitleStyle, marginBottom: 20 }}>총 {logs.length}건</p>

          {logs.length === 0 ? (
            <div style={{ textAlign: "center", padding: "60px 20px", color: C.textMuted }}>
              <div style={{ fontSize: 40, marginBottom: 12 }}>📋</div>
              <div>아직 분석 이력이 없습니다</div>
            </div>
          ) : logs.map(log => {
            const d = new Date(log.ts);
            const ds = `${d.getFullYear()}.${p2(d.getMonth()+1)}.${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`;
            return (
              <div key={log.id} style={{ borderRadius: 10, border: `1px solid ${C.border}`, background: C.surface, marginBottom: 10, overflow: "hidden" }}>
                <div style={{ padding: "14px 16px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 600 }}>{log.품명 || log.engine}</div>
                    <div style={{ fontSize: 12, color: C.textDim, marginTop: 2 }}>{log.equipment} · {ds}</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <span style={{ ...badge, background: C.ngDim, color: C.ng }}>NG {log.ng}</span>
                    <span style={{ ...badge, background: C.okDim, color: C.ok }}>보정값산출</span>
                  </div>
                </div>
                {log.corr?.length > 0 && (
                  <div style={{ padding: "0 16px 14px", display: "flex", flexWrap: "wrap", gap: 6 }}>
                    {log.corr.map((c, ci) => (
                      <div key={ci} style={{ padding: "6px 10px", borderRadius: 6, background: C.surfaceAlt, fontSize: 11, fontFamily: "monospace" }}>
                        <span style={{ color: C.textDim }}>{c.pt}</span>
                        {c.X !== 0 && <span style={{ color: C.ng, marginLeft: 6 }}>X{c.X > 0 ? "+" : ""}{c.X.toFixed(3)}</span>}
                        {c.Y !== 0 && <span style={{ color: C.blue, marginLeft: 6 }}>Y{c.Y > 0 ? "+" : ""}{c.Y.toFixed(3)}</span>}
                        {c.Z !== 0 && <span style={{ color: C.gold, marginLeft: 6 }}>Z{c.Z > 0 ? "+" : ""}{c.Z.toFixed(3)}</span>}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <style>{css}</style>
      </div>
    );
  }

  // ===================== 메인 화면 =====================
  return (
    <div style={shell}>
      <div style={{ position: "fixed", inset: 0, opacity: 0.03, pointerEvents: "none", backgroundImage: `linear-gradient(${C.text} 1px, transparent 1px), linear-gradient(90deg, ${C.text} 1px, transparent 1px)`, backgroundSize: "60px 60px" }} />

      <header style={hdr}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 36, height: 36, borderRadius: 8, background: `linear-gradient(135deg, ${C.accent}, ${C.accentDark})`, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ fontSize: 16, fontWeight: 900, color: "#fff", letterSpacing: -1, fontFamily: "'Arial Black', sans-serif" }}>M</span>
          </div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, letterSpacing: -0.3 }}>MAVIS</div>
            <div style={{ fontSize: 10, color: C.textMuted, marginTop: 1, letterSpacing: 0.5 }}>Machining Assistant for Visual Inspection & Solution</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <button onClick={() => setShowHistory(true)} style={{ ...iconBtn, position: "relative" }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M12 8v4l3 3"/><circle cx="12" cy="12" r="10"/></svg>
            {logs.length > 0 && <div style={{ position: "absolute", top: -4, right: -4, width: 16, height: 16, borderRadius: "50%", background: C.accent, color: "#fff", fontSize: 9, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" }}>{logs.length > 9 ? "9+" : logs.length}</div>}
          </button>
          {step > 0 && <button onClick={reset} style={{ ...ghostBtn, fontSize: 11, border: `1px solid ${C.border}`, padding: "5px 12px", borderRadius: 6 }}>초기화</button>}
          <div style={{ padding: "4px 8px", borderRadius: 4, background: C.accentDark, color: "#fff", fontSize: 9, fontWeight: 700, letterSpacing: 1 }}>사내한</div>
        </div>
      </header>

      {/* 스텝바 */}
      <div style={{ padding: "14px 20px", display: "flex", gap: 4, alignItems: "center" }}>
        {["엔진 선택","성적서 업로드","데이터 분석","보정값 산출"].map((l,i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: 4, flex: 1 }}>
            <div style={{ width: 24, height: 24, borderRadius: "50%", background: i <= step ? C.accent : C.surfaceAlt, border: `2px solid ${i <= step ? C.accent : C.border}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 700, color: i <= step ? "#fff" : C.textMuted, transition: "all 0.3s" }}>{i+1}</div>
            <span style={{ fontSize: 11, color: i <= step ? C.text : C.textMuted, fontWeight: i === step ? 600 : 400, display: i === step ? "block" : "none" }}>{l}</span>
            {i < 3 && <div style={{ flex: 1, height: 1, background: i < step ? C.accent : C.border }} />}
          </div>
        ))}
      </div>

      <main style={{ padding: "0 20px 100px", opacity: animateIn ? 1 : 0, transform: animateIn ? "translateY(0)" : "translateY(8px)", transition: "all 0.25s" }}>

        {/* STEP 0 */}
        {step === 0 && (<div>
          <h2 style={titleStyle}>엔진 수정 대상 선택</h2>
          <p style={subtitleStyle}>차종, 모델, 설비를 선택하세요</p>
          <div style={{ marginBottom: 20 }}>
            <label style={labelStyle}>차종</label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {ENGINE_TYPES.map(e => <Btn key={e.id} active={engineType?.id===e.id} onClick={() => { setEngineType(e); setEngineModel(null); }}>{e.label}</Btn>)}
            </div>
          </div>
          {engineType && <div style={{ marginBottom: 20 }}>
            <label style={labelStyle}>모델</label>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {engineType.models.map(m => <Btn key={m} active={engineModel===m} onClick={() => setEngineModel(m)} small>{m}</Btn>)}
            </div>
          </div>}
          {engineModel && <div style={{ marginBottom: 20 }}>
            <label style={labelStyle}>설비</label>
            <div style={{ borderRadius: 10, border: `1px solid ${C.border}`, overflow: "hidden" }}>
              {/* 상위: OP 번호 */}
              <div style={{ padding: "10px 12px", background: C.surfaceAlt, borderBottom: equipmentGroup ? `1px solid ${C.border}` : "none" }}>
                <div style={{ fontSize: 10, fontWeight: 600, color: C.textMuted, letterSpacing: 1, marginBottom: 8, textTransform: "uppercase" }}>공정</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 6 }}>
                  {["OP11","OP12","OP13","OP14","OP15","OP16"].map(g => (
                    <button key={g} onClick={() => { setEquipmentGroup(g); setEquipment(null); }} style={{
                      padding: "8px 4px", borderRadius: 6, border: `1px solid ${equipmentGroup===g ? C.accent : C.border}`,
                      background: equipmentGroup===g ? C.accentGlow : C.surface,
                      color: equipmentGroup===g ? C.text : C.textDim,
                      fontSize: 12, fontWeight: equipmentGroup===g ? 700 : 500,
                    }}>{g}</button>
                  ))}
                </div>
              </div>
              {/* 하위: A~E */}
              {equipmentGroup && (
                <div style={{ padding: "10px 12px", background: C.bg }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: C.textMuted, letterSpacing: 1, marginBottom: 8, textTransform: "uppercase" }}>라인 <span style={{ color: C.accent }}>{equipmentGroup}</span></div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 6 }}>
                    {["A","B","C","D","E"].map(l => (
                      <button key={l} onClick={() => setEquipment(equipmentGroup+l)} style={{
                        padding: "10px 4px", borderRadius: 6, border: `1px solid ${equipment===equipmentGroup+l ? C.accent : C.borderLight}`,
                        background: equipment===equipmentGroup+l ? C.accentGlow : C.surfaceAlt,
                        color: equipment===equipmentGroup+l ? C.text : C.textDim,
                        fontSize: 14, fontWeight: equipment===equipmentGroup+l ? 700 : 500,
                      }}>{l}</button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>}
          {equipment && <button onClick={() => goToStep(1)} style={primaryBtn}>다음 →</button>}
        </div>)}

        {/* STEP 1 */}
        {step === 1 && (<div>
          <h2 style={titleStyle}>품질 검사 성적서</h2>
          <p style={{ fontSize: 13, color: C.textDim, margin: "0 0 4px" }}>{engineModel} · {equipment}</p>
          <p style={{ fontSize: 12, color: C.textMuted, margin: "0 0 20px" }}>Inspection Report를 촬영하거나 업로드하세요</p>
          <input ref={fileInputRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={handleImageSelect} />

          {!imagePreview ? (
            <div style={{ border: `2px dashed ${C.border}`, borderRadius: 12, padding: "48px 20px", display: "flex", flexDirection: "column", alignItems: "center", gap: 16, background: C.surface }}>
              <div style={{ width: 64, height: 64, borderRadius: "50%", background: C.surfaceAlt, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28 }}>📷</div>
              <div style={{ textAlign: "center" }}>
                <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 4 }}>Inspection Report 촬영</div>
                <div style={{ fontSize: 12, color: C.textMuted }}>성적서를 촬영하거나 갤러리에서 선택</div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => { fileInputRef.current.setAttribute("capture","environment"); fileInputRef.current.click(); }} style={{ ...primaryBtn, width: "auto", padding: "10px 24px", fontSize: 13 }}>📷 촬영</button>
                <button onClick={() => { fileInputRef.current.removeAttribute("capture"); fileInputRef.current.click(); }} style={{ ...secBtn, width: "auto", padding: "10px 24px", fontSize: 13 }}>🖼 갤러리</button>
              </div>
            </div>
          ) : (<div>
            <div style={{ borderRadius: 12, overflow: "hidden", border: `1px solid ${C.border}`, marginBottom: 12 }}><img src={imagePreview} alt="성적서" style={{ width: "100%", display: "block" }} /></div>
            <button onClick={() => { setImageFile(null); setImagePreview(null); }} style={{ ...secBtn, padding: "10px", marginBottom: 16 }}>다시 촬영</button>
            {ocrError && <div style={{ padding: "12px 16px", borderRadius: 8, background: C.ngDim, border: `1px solid ${C.ng}33`, marginBottom: 12, fontSize: 13, color: C.ng }}>{ocrError}<button onClick={() => runOCR(true)} style={{ ...primaryBtn, marginTop: 8, padding: "8px 16px", fontSize: 12 }}>데모 데이터로 진행</button></div>}
            <button onClick={() => runOCR(false)} disabled={ocrLoading} style={{ ...primaryBtn, opacity: ocrLoading ? 0.7 : 1 }}>
              {ocrLoading ? <span style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><Spinner />분석 중...</span> : "🔍 분석하기"}
            </button>
            <button onClick={() => runOCR(true)} style={{ ...secBtn, marginTop: 8, fontSize: 12, color: C.textMuted }}>💡 데모 데이터로 테스트</button>
          </div>)}
          <button onClick={() => goToStep(0)} style={{ ...textBtn, marginTop: 20 }}>← 이전 단계</button>
        </div>)}

        {/* STEP 2 */}
        {step === 2 && ocrResult && (<div>
          <h2 style={titleStyle}>데이터 분석 결과</h2>
          <p style={subtitleStyle}>{ocrResult.품명 || engineModel}</p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 20 }}>
            <StatCard label="전체" value={totalRows} color={C.blue} bg={C.blueDim} />
            <StatCard label="OK" value={okCount} color={C.ok} bg={C.okDim} />
            <StatCard label="NG" value={ngCount} color={C.ng} bg={C.ngDim} />
          </div>

          {ngCount > 0 && <div style={{ padding: "14px 16px", borderRadius: 10, background: C.ngDim, border: `1px solid ${C.ng}22`, marginBottom: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: C.ng, marginBottom: 10 }}>⚠ 불량 항목 ({ngCount}건)</div>
            {ocrResult.rows.filter(r => r.ngOk === "NG").map((row, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderTop: i ? `1px solid ${C.ng}15` : "none" }}>
                <div><span style={{ fontSize: 12, color: C.textMuted }}>{row.type}</span><div style={{ fontSize: 14, fontWeight: 600 }}>{row.no} · {row.item}축</div></div>
                <div style={{ textAlign: "right" }}><div style={{ fontSize: 14, fontWeight: 700, color: C.ng }}>{row.deviation > 0 ? "+" : ""}{row.deviation?.toFixed(4)}</div><div style={{ fontSize: 11, color: C.textMuted }}>공차 ±{Math.abs(row.upTol)}</div></div>
              </div>
            ))}
          </div>}

          <div style={{ borderRadius: 10, overflow: "hidden", border: `1px solid ${C.border}`, marginBottom: 20 }}>
            <div style={{ padding: "10px 14px", background: C.surfaceAlt, fontSize: 13, fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>전체 측정 데이터</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead><tr style={{ background: C.surface }}>{["Type","No.","축","실측","기준","편차","판정"].map(h => <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 600, color: C.textDim, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
                <tbody>{ocrResult.rows.map((r,i) => {
                  const ng = r.ngOk === "NG";
                  return <tr key={i} style={{ background: ng ? C.ngDim : (i%2 ? C.surfaceAlt : "transparent") }}>
                    <td style={td}><span style={{ fontSize: 11, color: C.textMuted }}>{r.type?.substring(0,18)}</span></td>
                    <td style={{ ...td, fontWeight: 600 }}>{r.no}</td>
                    <td style={td}>{r.item}</td>
                    <td style={{ ...td, fontFamily: "monospace" }}>{r.actual}</td>
                    <td style={{ ...td, fontFamily: "monospace", color: C.textMuted }}>{r.nominal}</td>
                    <td style={{ ...td, fontFamily: "monospace", fontWeight: 600, color: ng ? C.ng : C.text }}>{r.deviation > 0 ? "+" : ""}{r.deviation?.toFixed(4)}</td>
                    <td style={td}><span style={{ display: "inline-block", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700, background: ng ? C.ng : C.ok, color: "#fff" }}>{r.ngOk}</span></td>
                  </tr>;
                })}</tbody>
              </table>
            </div>
          </div>
          {ngCount > 0 && <button onClick={calculateCorrections} style={primaryBtn}>🔧 보정값 산출 →</button>}
          <button onClick={() => goToStep(1)} style={{ ...textBtn, marginTop: 12 }}>← 이전 단계</button>
        </div>)}

        {/* STEP 3 */}
        {step === 3 && corrections && (<div>
          <h2 style={titleStyle}>분석 결과</h2>
          <p style={subtitleStyle}>{ocrResult?.품명 || engineModel} · {equipment} · 수정 대상 {corrections.length}건</p>

          {corrections.map((corr, i) => (
            <div key={i} style={{ borderRadius: 12, overflow: "hidden", border: `1px solid ${C.border}`, marginBottom: 12, background: C.surface }}>
              <div style={{ padding: "14px 16px", background: `linear-gradient(135deg, ${C.accentGlow}, transparent)`, borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ width: 32, height: 32, borderRadius: 8, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 800, color: "#fff" }}>{i+1}</div>
                  <div><div style={{ fontSize: 11, color: C.textMuted }}>수정 대상 포인트</div><div style={{ fontSize: 16, fontWeight: 700 }}>{corr.point}</div></div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 11, color: C.textMuted }}>{corr.type?.substring(0,20)}</div>
                  <div style={{ fontSize: 11, fontFamily: "monospace", color: C.textDim, marginTop: 2 }}>{corr.program} · {corr.tool} · {corr.coord}</div>
                </div>
              </div>
              <div style={{ padding: 16 }}>
                <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 10 }}>보정값</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                  {["X","Y","Z"].map(ax => {
                    const v = corr[ax], has = v !== 0;
                    return <div key={ax} style={{ padding: 12, borderRadius: 8, background: has ? C.surfaceAlt : `${C.surfaceAlt}66`, border: `1px solid ${has ? C.borderLight : C.border}`, textAlign: "center" }}>
                      <div style={{ fontSize: 11, fontWeight: 600, color: has ? C.text : C.textMuted, marginBottom: 4 }}>{ax}축</div>
                      <div style={{ fontSize: 20, fontWeight: 800, fontFamily: "monospace", color: !has ? C.textMuted : v > 0 ? C.ng : C.blue }}>{v > 0 ? "+" : ""}{v.toFixed(3)}</div>
                    </div>;
                  })}
                </div>
                {Object.keys(corr.measurements||{}).length > 0 && <div style={{ marginTop: 12 }}>
                  <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 6 }}>측정 상세</div>
                  {Object.entries(corr.measurements).map(([ax,m]) => (
                    <div key={ax} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", fontSize: 12, borderTop: `1px solid ${C.border}` }}>
                      <span style={{ color: C.textDim }}>{ax}축: {m.actual} (기준 {m.nominal})</span>
                      <span style={{ fontFamily: "monospace", color: C.ng }}>편차 {m.deviation > 0 ? "+" : ""}{m.deviation?.toFixed(4)} → {m.corrAxis}축 {m.corrValue > 0 ? "+" : ""}{m.corrValue?.toFixed(4)}</span>
                    </div>
                  ))}
                </div>}
              </div>
            </div>
          ))}

          <div style={{ marginTop: 12, display: "flex", gap: 8 }}>
            <button onClick={() => goToStep(2)} style={{ ...secBtn, flex: 1 }}>← 데이터 확인</button>
            <button onClick={reset} style={{ ...secBtn, flex: 1, color: C.textMuted }}>🏠 홈으로</button>
          </div>
        </div>)}

      </main>
      <style>{css}</style>
    </div>
  );
}

// ============================================================
// 서브 컴포넌트
// ============================================================

function StatCard({ label, value, color, bg }) {
  return <div style={{ padding: "14px 12px", borderRadius: 10, background: bg, border: `1px solid ${color}22`, textAlign: "center" }}>
    <div style={{ fontSize: 26, fontWeight: 800, color, fontFamily: "monospace" }}>{value}</div>
    <div style={{ fontSize: 11, color: `${color}cc`, marginTop: 2, fontWeight: 500 }}>{label}</div>
  </div>;
}

function Btn({ children, active, onClick, small, compact }) {
  return <button onClick={onClick} style={{
    display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: compact ? 10 : small ? "12px 16px" : "14px 16px",
    borderRadius: small ? 8 : 10,
    border: `1px solid ${active ? C.accent : C.border}`,
    background: active ? C.accentGlow : C.surface,
    color: C.text, fontSize: compact ? 13 : small ? 13 : 15,
    fontWeight: compact ? 500 : 600, textAlign: "left",
  }}>
    <span>{children}</span>
    {active && <span style={{ width: compact ? 18 : 22, height: compact ? 18 : 22, borderRadius: "50%", background: C.accent, color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: compact ? 10 : 12, fontWeight: 700, flexShrink: 0 }}>✓</span>}
  </button>;
}

function Spinner() {
  return <span style={{ display: "inline-block", width: 16, height: 16, border: "2px solid #fff4", borderTopColor: "#fff", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />;
}

function p2(n) { return String(n).padStart(2, "0"); }

// ============================================================
// 스타일
// ============================================================

const shell = { minHeight: "100vh", background: `linear-gradient(180deg, ${C.bg} 0%, #08080d 100%)`, color: C.text, fontFamily: "'Pretendard','Apple SD Gothic Neo',sans-serif", position: "relative", overflow: "hidden" };
const hdr = { padding: "16px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: `1px solid ${C.border}`, background: `${C.surface}cc`, backdropFilter: "blur(12px)", position: "sticky", top: 0, zIndex: 100 };
const titleStyle = { fontSize: 20, fontWeight: 700, margin: "20px 0 6px", letterSpacing: -0.5 };
const subtitleStyle = { fontSize: 13, color: C.textDim, margin: "0 0 20px" };
const labelStyle = { display: "block", fontSize: 12, fontWeight: 600, color: C.textDim, marginBottom: 8, textTransform: "uppercase", letterSpacing: 1 };
const primaryBtn = { width: "100%", padding: "14px", borderRadius: 10, border: "none", fontSize: 15, fontWeight: 700, background: `linear-gradient(135deg, ${C.accent}, ${C.accentDark})`, color: "#fff", letterSpacing: -0.3, boxShadow: `0 4px 20px ${C.accent}40` };
const secBtn = { width: "100%", padding: "12px", borderRadius: 10, border: `1px solid ${C.border}`, background: C.surface, color: C.text, fontSize: 14, fontWeight: 500 };
const textBtn = { background: "none", border: "none", color: C.textMuted, fontSize: 13, padding: "8px 0" };
const ghostBtn = { background: "none", border: "none", color: C.text };
const iconBtn = { width: 34, height: 34, borderRadius: 8, border: `1px solid ${C.border}`, background: "transparent", color: C.textDim, display: "flex", alignItems: "center", justifyContent: "center" };
const badge = { padding: "3px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700 };
const td = { padding: "8px 10px", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" };

const css = `
  @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
  @keyframes spin { to { transform: rotate(360deg); } }
  @keyframes popIn { 0% { transform: scale(0.5); opacity: 0; } 60% { transform: scale(1.15); } 100% { transform: scale(1); opacity: 1; } }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  button { cursor: pointer; transition: all 0.2s ease; }
  button:active { transform: scale(0.97); }
  ::-webkit-scrollbar { width: 4px; height: 4px; }
  ::-webkit-scrollbar-thumb { background: #2a2a3a; border-radius: 4px; }
`;
