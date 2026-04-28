"use client";

import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── 타입 ────────────────────────────────────────────────────────────────────

interface Module {
  cf: string;
  상위모듈: string;
  모듈코드: string;
  모듈명: string;
  발화목록: string[];
  중요: boolean; // 엑셀 "중요" 컬럼에 값이 있으면 true
}

interface Branch {
  cf: string;
  모듈명: string;
  응답유형: string;
  다음모듈: string;
  비고: string;
}

interface CustomImportant {
  제목: string;
  내용: string;
}

interface ParsedData {
  modules: Module[];
  branches: Branch[];
  customImportant: CustomImportant[]; // "중요발화" 시트에서 읽어온 직접 작성 발화
}

interface SavedScript {
  id: string;
  모듈명: string;
  모듈코드: string;
  cf: string;
  발화: string;
  savedAt: string;
}

interface CustomerInfo {
  customerName: string;
  gender: string;
  age: string;
  consultantName: string;
  renewalType: string;
  coverageRange: string;
  coverageAmount: string;
  coveragePeriod: string;
}

// ─── 상수 ────────────────────────────────────────────────────────────────────

const CF_LABELS: Record<string, string> = {
  "CF1. 상담 가치 형성(구체화)": "CF1 상담 가치 형성",
  "CF2. 문제 인식": "CF2 문제 인식",
  "CF3. 솔루션 확정": "CF3 솔루션 확정",
  "CF4. 청약 완료": "CF4 청약 완료",
  "CF5. 상담 마무리": "CF5 상담 마무리",
};

const INPUT_CLS =
  "w-full rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white";

const STORAGE_KEY = "savedScripts_v1";

// ─── 이름 치환 ────────────────────────────────────────────────────────────────

function applyReplacements(text: string, info: CustomerInfo): string {
  const { customerName, consultantName } = info;
  let r = text;
  if (consultantName) {
    r = r.replace(/보험전문가\s*\[이름\]/g, `보험전문가 ${consultantName}`);
    r = r.replace(/보험전문가\s*○{2,3}/g, `보험전문가 ${consultantName}`);
    r = r.replace(/보험전문가\s*O{2,3}/g, `보험전문가 ${consultantName}`);
    r = r.replace(/\[이름\](입니다|이에요|예요|이야)/g, `${consultantName}$1`);
  }
  if (customerName) {
    r = r.replace(/○{2,3}님/g, `${customerName}님`);
    r = r.replace(/O{2,3}님/g, `${customerName}님`);
    r = r.replace(/0{2}님/g, `${customerName}님`);
    r = r.replace(/\[이름\]님/g, `${customerName}님`);
    r = r.replace(/고객님/g, `${customerName}님`);
  }
  return r;
}

// ─── 엑셀 파싱 ────────────────────────────────────────────────────────────────

function parseExcel(file: File): Promise<ParsedData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });

        // ── 시트3: 발화 스크립트 ──
        const scriptSheet = wb.Sheets[wb.SheetNames[2]];
        const scriptRows = XLSX.utils.sheet_to_json<Record<string, string>>(scriptSheet, { defval: "" });
        // "중요" 컬럼은 발화 텍스트가 아니므로 스크립트 컬럼에서 제외
        const RESERVED_COLS = ["CF", "상위 모듈", "모듈 코드", "모듈명", "중요"];
        const scriptCols = Object.keys(scriptRows[0] || {}).filter(
          (k) => !RESERVED_COLS.includes(k)
        );
        const modules: Module[] = scriptRows
          .filter((r) => r["모듈명"])
          .map((r) => {
            // "중요" 컬럼: 비어있지 않고 N/no/false 가 아니면 중요로 간주
            const rawImportant = r["중요"]?.toString().trim() ?? "";
            const 중요 =
              rawImportant !== "" &&
              !["n", "no", "false", "아니오"].includes(rawImportant.toLowerCase());
            return {
              cf: r["CF"] || "",
              상위모듈: r["상위 모듈"] || "",
              모듈코드: r["모듈 코드"] || "",
              모듈명: r["모듈명"] || "",
              발화목록: scriptCols.map((c) => r[c]?.toString().trim()).filter(Boolean),
              중요,
            };
          })
          .filter((m) => m.발화목록.length > 0);

        // ── 시트1: 분기 테이블 ──
        const branchSheet = wb.Sheets[wb.SheetNames[0]];
        const branchRows = XLSX.utils.sheet_to_json<Record<string, string>>(branchSheet, { defval: "" });
        const branches: Branch[] = branchRows
          .filter((r) => r["현재 모듈명"] && r["고객 응답 유형"] && r["다음 모듈"])
          .map((r) => ({
            cf: r["CF"] || "",
            모듈명: r["현재 모듈명"],
            응답유형: r["고객 응답 유형"],
            다음모듈: r["다음 모듈"],
            비고: r["비고"] || "",
          }));

        // ── "중요발화" 시트: 직접 작성한 중요 발화 (없으면 빈 배열)
        const customSheet = wb.Sheets["중요발화"];
        const customImportant: CustomImportant[] = customSheet
          ? XLSX.utils.sheet_to_json<Record<string, string>>(customSheet, { defval: "" })
              .filter((r) => r["내용"]?.toString().trim())
              .map((r) => ({
                제목: r["제목"]?.toString().trim() || "",
                내용: r["내용"]?.toString().trim(),
              }))
          : [];

        resolve({ modules, branches, customImportant });
      } catch {
        reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
      }
    };
    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsArrayBuffer(file);
  });
}

// ─── 모듈 필터링 (선택 안한 항목 관련 모듈 제외) ────────────────────────────────

const FIELD_MODULE_PATTERNS: { field: keyof CustomerInfo; patterns: string[] }[] = [
  { field: "renewalType",    patterns: ["갱신형"] },
  { field: "coverageRange",  patterns: ["보장범위"] },
  { field: "coverageAmount", patterns: ["보장금액"] },
  { field: "coveragePeriod", patterns: ["보장기간", "납입기간"] },
];

function shouldShowModule(module: Module, info: CustomerInfo): boolean {
  for (const { field, patterns } of FIELD_MODULE_PATTERNS) {
    const isRelated = patterns.some((p) => module.모듈명.includes(p));
    if (isRelated && !info[field]) return false;
  }
  return true;
}

function findNextVisible(
  fromModule: Module,
  modules: Module[],
  info: CustomerInfo
): Module | undefined {
  const idx = modules.findIndex((m) => m.모듈명 === fromModule.모듈명);
  for (let i = idx + 1; i < modules.length; i++) {
    if (shouldShowModule(modules[i], info)) return modules[i];
  }
  return undefined;
}

// ─── 모듈 검색 (유사 이름 매칭) ──────────────────────────────────────────────

function findModule(targetName: string, modules: Module[]): Module | undefined {
  const name = targetName.includes("→")
    ? targetName.split("→").pop()!.trim()
    : targetName.trim();

  return (
    modules.find((m) => m.모듈명 === name) ||
    modules.find((m) => m.모듈명.includes(name) || name.includes(m.모듈명))
  );
}

function getBranches(moduleName: string, branches: Branch[]): Branch[] {
  return branches.filter(
    (b) => b.모듈명 === moduleName || moduleName.includes(b.모듈명) || b.모듈명.includes(moduleName)
  );
}

function pickScript(module: Module, info: CustomerInfo): string {
  const line = module.발화목록[Math.floor(Math.random() * module.발화목록.length)];
  return applyReplacements(line, info);
}

// ─── 날짜 포맷 ────────────────────────────────────────────────────────────────

function formatDate(iso: string): string {
  const d = new Date(iso);
  return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

// ─── 컴포넌트 ─────────────────────────────────────────────────────────────────

export default function Home() {
  const [parsedData, setParsedData] = useState<ParsedData | null>(null);
  const [fileName, setFileName] = useState("");
  const [fileError, setFileError] = useState("");
  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [customerName, setCustomerName] = useState("");
  const [gender, setGender] = useState("");
  const [age, setAge] = useState("");
  const [consultantName, setConsultantName] = useState("");
  const [renewalType, setRenewalType] = useState("");
  const [coverageRange, setCoverageRange] = useState("");
  const [coverageAmount, setCoverageAmount] = useState("");
  const [coveragePeriod, setCoveragePeriod] = useState("");

  const info: CustomerInfo = {
    customerName, gender, age, consultantName,
    renewalType, coverageRange, coverageAmount, coveragePeriod,
  };

  const [started, setStarted] = useState(false);
  const [currentModule, setCurrentModule] = useState<Module | null>(null);
  const [currentScript, setCurrentScript] = useState("");
  const [history, setHistory] = useState<{ module: Module; script: string }[]>([]);
  const [done, setDone] = useState(false);

  // ── 홈 탭 (상담 | 저장된 발화) ──
  const [homeTab, setHomeTab] = useState<"상담" | "저장">("상담");

  // ── 저장된 발화 (localStorage 동기화) ──
  const [savedScripts, setSavedScripts] = useState<SavedScript[]>([]);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setSavedScripts(JSON.parse(raw));
    } catch {
      // ignore
    }
  }, []);

  function persistSaved(next: SavedScript[]) {
    setSavedScripts(next);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch {
      // ignore
    }
  }

  function handleSave() {
    if (!currentModule || !currentScript) return;
    // 동일 발화 중복 저장 방지
    if (savedScripts.some((s) => s.발화 === currentScript && s.모듈명 === currentModule.모듈명)) return;
    const newItem: SavedScript = {
      id: Date.now().toString(),
      모듈명: currentModule.모듈명,
      모듈코드: currentModule.모듈코드,
      cf: currentModule.cf,
      발화: currentScript,
      savedAt: new Date().toISOString(),
    };
    persistSaved([newItem, ...savedScripts]);
  }

  function handleDeleteSaved(id: string) {
    persistSaved(savedScripts.filter((s) => s.id !== id));
  }

  function handleClearAll() {
    if (confirm("저장된 발화를 모두 삭제할까요?")) persistSaved([]);
  }

  // 현재 스크립트가 이미 저장됐는지
  const isAlreadySaved =
    !!currentModule &&
    savedScripts.some((s) => s.발화 === currentScript && s.모듈명 === currentModule.모듈명);

  // ── 파일 ──
  const handleFile = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      setFileError("엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.");
      return;
    }
    setFileError("");
    try {
      const data = await parseExcel(file);
      setParsedData(data);
      setFileName(file.name);
      setStarted(false);
      setDone(false);
    } catch (e) {
      setFileError((e as Error).message);
    }
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    },
    [handleFile]
  );

  function handleStart() {
    if (!parsedData) return;
    const first = parsedData.modules.find((m) => shouldShowModule(m, info));
    if (!first) return;
    setCurrentModule(first);
    setCurrentScript(pickScript(first, info));
    setHistory([]);
    setDone(false);
    setStarted(true);
  }

  function handleResponse(branch: Branch) {
    if (!parsedData || !currentModule) return;
    setHistory((h) => [...h, { module: currentModule, script: currentScript }]);

    const nextName = branch.다음모듈;
    if (
      nextName.includes("종료") ||
      nextName.includes("CF1부터") ||
      currentModule.모듈명 === "마무리 인사"
    ) {
      setDone(true);
      return;
    }

    const next = findModule(nextName, parsedData.modules);
    const target = next && shouldShowModule(next, info)
      ? next
      : next
        ? findNextVisible(next, parsedData.modules, info)
        : findNextVisible(currentModule, parsedData.modules, info);

    if (target) {
      setCurrentModule(target);
      setCurrentScript(pickScript(target, info));
    } else {
      setDone(true);
    }
  }

  function handleNext() {
    if (!parsedData || !currentModule) return;
    setHistory((h) => [...h, { module: currentModule, script: currentScript }]);
    const next = findNextVisible(currentModule, parsedData.modules, info);
    if (next) {
      setCurrentModule(next);
      setCurrentScript(pickScript(next, info));
    } else {
      setDone(true);
    }
  }

  function handleBack() {
    if (history.length === 0) return;
    const prev = history[history.length - 1];
    setHistory((h) => h.slice(0, -1));
    setCurrentModule(prev.module);
    setCurrentScript(prev.script);
    setDone(false);
  }

  function handleReroll() {
    if (!currentModule) return;
    setCurrentScript(pickScript(currentModule, info));
  }

  const currentBranches = parsedData && currentModule
    ? getBranches(currentModule.모듈명, parsedData.branches)
    : [];

  return (
    <main className="min-h-screen py-10 px-4">
      <div className="max-w-2xl mx-auto">

        <div className="mb-8 text-center">
          <h1 className="text-2xl font-bold text-black mb-1">보험 상담 스크립트</h1>
          <p className="text-sm text-black">고객 반응에 따라 다음 발화를 안내해드립니다.</p>
        </div>

        {/* ── 준비 화면 ── */}
        {!started && (
          <>
            {/* 탭 */}
            <div className="flex gap-1 mb-4 bg-gray-200 rounded-xl p-1">
              <button
                onClick={() => setHomeTab("상담")}
                className={`flex-1 py-2 rounded-lg text-sm font-semibold transition-colors ${
                  homeTab === "상담"
                    ? "bg-white text-black shadow-sm"
                    : "text-gray-500 hover:text-black"
                }`}
              >
                상담 시작
              </button>
              <button
                onClick={() => setHomeTab("저장")}
                className={`flex-1 py-2 rounded-lg text-sm font-semibold transition-colors relative ${
                  homeTab === "저장"
                    ? "bg-white text-black shadow-sm"
                    : "text-gray-500 hover:text-black"
                }`}
              >
                저장된 발화
                {savedScripts.length > 0 && (
                  <span className="ml-1.5 inline-flex items-center justify-center w-5 h-5 rounded-full bg-blue-500 text-white text-xs font-bold">
                    {savedScripts.length}
                  </span>
                )}
              </button>
            </div>

            {/* ── 상담 탭 ── */}
            {homeTab === "상담" && (
              <>
                <div className="bg-white rounded-2xl border border-gray-300 p-6 mb-4">
                  <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">상담 정보 입력</h2>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">고객 이름</label>
                      <input type="text" placeholder="홍길동" value={customerName}
                        onChange={(e) => setCustomerName(e.target.value)} className={INPUT_CLS} />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">설계사 이름</label>
                      <input type="text" placeholder="내 이름" value={consultantName}
                        onChange={(e) => setConsultantName(e.target.value)} className={INPUT_CLS} />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">성별</label>
                      <select value={gender} onChange={(e) => setGender(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="남성">남성</option>
                        <option value="여성">여성</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">나이</label>
                      <input type="text" placeholder="예: 35세" value={age}
                        onChange={(e) => setAge(e.target.value)} className={INPUT_CLS} />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">갱신형</label>
                      <select value={renewalType} onChange={(e) => setRenewalType(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="갱신형">갱신형</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">보장범위</label>
                      <select value={coverageRange} onChange={(e) => setCoverageRange(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="좁음">좁음</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">보장금액</label>
                      <select value={coverageAmount} onChange={(e) => setCoverageAmount(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="부족">부족</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">보장기간</label>
                      <select value={coveragePeriod} onChange={(e) => setCoveragePeriod(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="짧음">짧음</option>
                      </select>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-gray-300 p-6 mb-6">
                  <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">엑셀 파일 업로드</h2>
                  <div
                    onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
                    onDragLeave={() => setDragging(false)}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                    className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors
                      ${dragging ? "border-blue-400 bg-blue-50" : "border-gray-300 hover:border-gray-400 hover:bg-gray-50"}`}
                  >
                    <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden"
                      onChange={(e) => { if (e.target.files?.[0]) handleFile(e.target.files[0]); }} />
                    {fileName ? (
                      <>
                        <p className="text-sm font-semibold text-black mb-1">{fileName}</p>
                        <p className="text-xs text-black">
                          발화 {parsedData?.modules.length}개 · 분기 {parsedData?.branches.length}개 로드됨 — 클릭하면 다시 업로드
                        </p>
                      </>
                    ) : (
                      <>
                        <p className="text-sm text-black mb-1">엑셀 파일을 드래그하거나 클릭해서 업로드</p>
                        <p className="text-xs text-black">.xlsx, .xls 지원 (시트1: 분기, 시트3: 발화)</p>
                      </>
                    )}
                  </div>
                  {fileError && <p className="mt-3 text-sm text-red-500">{fileError}</p>}

                  <button
                    onClick={handleStart}
                    disabled={!parsedData}
                    className="mt-5 w-full py-3 rounded-lg text-sm font-semibold transition-colors
                      bg-blue-600 text-white hover:bg-blue-700 disabled:bg-gray-200 disabled:text-black disabled:cursor-not-allowed"
                  >
                    {parsedData ? "상담 시작" : "파일을 먼저 업로드해주세요"}
                  </button>
                </div>
              </>
            )}

            {/* ── 저장된 발화 탭 ── */}
            {homeTab === "저장" && (
              <div className="space-y-4 mb-6">

                {/* 엑셀 중요 발화 섹션 */}
                {parsedData && parsedData.modules.some((m) => m.중요) && (
                  <div className="bg-white rounded-2xl border border-yellow-300 p-6">
                    <div className="flex items-center gap-2 mb-4">
                      <span className="text-yellow-400 text-sm">★</span>
                      <h2 className="text-sm font-semibold text-black">엑셀 중요 발화</h2>
                      <span className="ml-auto text-xs text-gray-400">
                        파일에서 자동으로 불러온 중요 발화입니다.
                      </span>
                    </div>
                    <div className="space-y-3">
                      {parsedData.modules
                        .filter((m) => m.중요)
                        .map((m) => (
                          <div key={m.모듈코드 || m.모듈명} className="rounded-xl border border-yellow-100 bg-yellow-50 p-4">
                            <div className="flex items-center gap-2 mb-2">
                              <span className="text-xs font-semibold text-yellow-700">{m.모듈명}</span>
                              {m.모듈코드 && (
                                <span className="text-xs text-gray-400">{m.모듈코드}</span>
                              )}
                            </div>
                            <div className="space-y-2">
                              {m.발화목록.map((line, i) => (
                                <p
                                  key={i}
                                  className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-white rounded-lg p-3 border border-yellow-100"
                                >
                                  {applyReplacements(line, info)}
                                </p>
                              ))}
                            </div>
                          </div>
                        ))}
                    </div>
                  </div>
                )}

                {/* 중요발화 시트 섹션 */}
                {parsedData && parsedData.customImportant.length > 0 && (
                  <div className="bg-white rounded-2xl border border-blue-200 p-6">
                    <div className="flex items-center gap-2 mb-4">
                      <span className="text-blue-400 text-sm">📋</span>
                      <h2 className="text-sm font-semibold text-black">직접 작성한 중요 발화</h2>
                      <span className="ml-auto text-xs text-gray-400">엑셀 &quot;중요발화&quot; 시트</span>
                    </div>
                    <div className="space-y-3">
                      {parsedData.customImportant.map((item, i) => (
                        <div key={i} className="rounded-xl border border-blue-100 bg-blue-50 p-4">
                          {item.제목 && (
                            <p className="text-xs font-semibold text-blue-700 mb-2">{item.제목}</p>
                          )}
                          <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-white rounded-lg p-3 border border-blue-100">
                            {applyReplacements(item.내용, info)}
                          </p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* ★ 직접 저장한 발화 섹션 */}
                <div className="bg-white rounded-2xl border border-gray-300 p-6">
                  <div className="flex items-center justify-between mb-4">
                    <h2 className="text-sm font-semibold text-black">직접 저장한 발화</h2>
                    {savedScripts.length > 0 && (
                      <button
                        onClick={handleClearAll}
                        className="text-xs text-red-400 hover:text-red-600 transition-colors"
                      >
                        전체 삭제
                      </button>
                    )}
                  </div>

                  {savedScripts.length === 0 ? (
                    <div className="py-10 text-center">
                      <p className="text-3xl mb-3">★</p>
                      <p className="text-sm text-gray-400">아직 저장된 발화가 없습니다.</p>
                      <p className="text-xs text-gray-400 mt-1">상담 중 ★ 버튼을 눌러 발화를 저장하세요.</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {savedScripts.map((s) => (
                        <div key={s.id} className="rounded-xl border border-gray-200 p-4">
                          <div className="flex items-start justify-between gap-2 mb-2">
                            <div>
                              <span className="text-xs font-semibold text-blue-600">{s.모듈명}</span>
                              {s.모듈코드 && (
                                <span className="ml-2 text-xs text-gray-400">{s.모듈코드}</span>
                              )}
                            </div>
                            <div className="flex items-center gap-2 flex-shrink-0">
                              <span className="text-xs text-gray-400">{formatDate(s.savedAt)}</span>
                              <button
                                onClick={() => handleDeleteSaved(s.id)}
                                className="text-gray-300 hover:text-red-400 transition-colors text-base leading-none"
                                title="삭제"
                              >
                                ✕
                              </button>
                            </div>
                          </div>
                          <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-gray-50 rounded-lg p-3">
                            {s.발화}
                          </p>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

              </div>
            )}
          </>
        )}

        {/* ── 상담 진행 화면 ── */}
        {started && !done && currentModule && (
          <>
            <div className="flex items-center gap-2 mb-4">
              {history.length > 0 && (
                <button
                  onClick={handleBack}
                  className="text-sm text-black bg-white border border-gray-300 rounded-lg px-3 py-1.5 hover:bg-gray-100 transition-colors whitespace-nowrap"
                >
                  ← 이전 모듈
                </button>
              )}
              <button
                onClick={() => { setStarted(false); setDone(false); setHistory([]); }}
                className="text-sm text-black bg-white border border-gray-300 rounded-lg px-3 py-1.5 hover:bg-gray-100 transition-colors whitespace-nowrap"
              >
                ⌂ 첫 화면
              </button>
              {history.length > 0 && (
                <p className="text-xs text-black truncate">
                  {history.map((h) => h.module.모듈명).join(" › ")} › <span className="font-semibold">{currentModule.모듈명}</span>
                </p>
              )}
            </div>

            <div className="bg-white rounded-2xl border border-gray-300 overflow-hidden mb-4">
              <div className="bg-blue-600 px-5 py-3 flex items-center justify-between">
                <div>
                  <p className="text-xs text-blue-100">{CF_LABELS[currentModule.cf] ?? currentModule.cf}</p>
                  <p className="text-base font-bold text-white">{currentModule.모듈명}</p>
                  <p className="text-xs text-blue-200">{currentModule.모듈코드}</p>
                </div>
                {(customerName || gender || age) && (
                  <div className="text-right text-xs text-blue-100 space-y-0.5">
                    {customerName && <p>{customerName}님</p>}
                    {gender && age && <p>{gender} · {age}</p>}
                    {gender && !age && <p>{gender}</p>}
                    {!gender && age && <p>{age}</p>}
                  </div>
                )}
              </div>

              <div className="p-5">
                <div className="flex items-start justify-between gap-3 mb-1">
                  <p className="text-xs font-semibold text-black">발화 스크립트</p>
                  <div className="flex items-center gap-3 flex-shrink-0">
                    <button
                      onClick={handleSave}
                      disabled={isAlreadySaved}
                      title={isAlreadySaved ? "이미 저장됨" : "발화 저장"}
                      className={`text-sm transition-colors ${
                        isAlreadySaved
                          ? "text-yellow-400 cursor-default"
                          : "text-gray-300 hover:text-yellow-400"
                      }`}
                    >
                      ★
                    </button>
                    <button
                      onClick={handleReroll}
                      className="text-xs text-blue-600 hover:text-blue-800 whitespace-nowrap"
                    >
                      ↻ 다시 뽑기
                    </button>
                  </div>
                </div>
                <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-gray-50 rounded-lg p-4">
                  {currentScript}
                </p>
              </div>
            </div>

            <div className="bg-white rounded-2xl border border-gray-300 p-5">
              {currentBranches.length > 0 && (
                <>
                  <p className="text-xs font-semibold text-black mb-3">고객 반응은?</p>
                  <div className="space-y-2 mb-3">
                    {currentBranches.map((b, i) => (
                      <button
                        key={i}
                        onClick={() => handleResponse(b)}
                        className="w-full text-left rounded-xl border border-gray-200 px-4 py-3 hover:border-blue-400 hover:bg-blue-50 transition-colors group"
                      >
                        <p className="text-sm font-semibold text-black group-hover:text-blue-700">{b.응답유형}</p>
                        <p className="text-xs text-black mt-0.5">
                          → {b.다음모듈}
                          {b.비고 && <span className="ml-2 text-gray-400">({b.비고})</span>}
                        </p>
                      </button>
                    ))}
                  </div>
                  <hr className="border-gray-200 mb-3" />
                </>
              )}
              <button
                onClick={handleNext}
                className="w-full py-3 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700 transition-colors"
              >
                다음 모듈로 →
              </button>
            </div>
          </>
        )}

        {/* ── 상담 완료 화면 ── */}
        {started && done && (
          <div className="bg-white rounded-2xl border border-gray-300 p-10 text-center">
            <p className="text-2xl mb-2">🎉</p>
            <p className="text-lg font-bold text-black mb-1">상담 완료</p>
            {customerName && (
              <p className="text-sm text-black mb-6">{customerName}님과의 상담이 마무리되었습니다.</p>
            )}
            <div className="flex gap-3 justify-center">
              <button
                onClick={handleBack}
                className="px-5 py-2 rounded-lg border border-gray-300 text-sm font-medium text-black bg-white hover:bg-gray-100"
              >
                ← 이전으로
              </button>
              <button
                onClick={() => { setStarted(false); setDone(false); setHistory([]); }}
                className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700"
              >
                새 상담 시작
              </button>
            </div>
          </div>
        )}

      </div>
    </main>
  );
}
