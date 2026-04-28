"use client";

import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

interface Module {
  cf: string;
  상위모듈: string;
  모듈코드: string;
  모듈명: string;
  발화목록: string[];
}

interface ScriptLine {
  모듈코드: string;
  모듈명: string;
  발화: string;
}

interface StageScript {
  cf: string;
  lines: ScriptLine[];
}

interface CustomerInfo {
  customerName: string;
  gender: string;
  age: string;
  consultantName: string;
}

const CF_LABELS: Record<string, string> = {
  "CF1. 상담 가치 형성(구체화)": "1단계 — 상담 가치 형성",
  "CF2. 문제 인식": "2단계 — 문제 인식",
  "CF3. 솔루션 확정": "3단계 — 솔루션 확정",
  "CF4. 청약 완료": "4단계 — 청약 완료",
  "CF5. 상담 마무리": "5단계 — 상담 마무리",
};

function applyReplacements(text: string, info: CustomerInfo): string {
  const { customerName, consultantName } = info;
  let result = text;

  if (consultantName) {
    // 설계사 이름 — 고객명보다 먼저 처리
    result = result.replace(/보험전문가\s*\[이름\]/g, `보험전문가 ${consultantName}`);
    result = result.replace(/보험전문가\s*○{2,3}/g, `보험전문가 ${consultantName}`);
    result = result.replace(/보험전문가\s*O{2,3}/g, `보험전문가 ${consultantName}`);
    result = result.replace(/\[이름\](입니다|이에요|예요|이야)/g, `${consultantName}$1`);
  }

  if (customerName) {
    // ○○님 / ○○○님 (U+25CB 화이트서클)
    result = result.replace(/○{2,3}님/g, `${customerName}님`);
    // OOO님 / OO님 (영문 대문자 O)
    result = result.replace(/O{2,3}님/g, `${customerName}님`);
    // 00님 (숫자 0)
    result = result.replace(/0{2}님/g, `${customerName}님`);
    // [이름]님
    result = result.replace(/\[이름\]님/g, `${customerName}님`);
    // 고객님
    result = result.replace(/고객님/g, `${customerName}님`);
  }

  return result;
}

function parseExcel(file: File): Promise<Module[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<Record<string, string>>(ws);

        const scriptCols = Object.keys(rows[0] || {}).filter(
          (k) => !["CF", "상위 모듈", "모듈 코드", "모듈명"].includes(k)
        );

        const modules: Module[] = rows
          .filter((r) => r["모듈명"])
          .map((r) => ({
            cf: r["CF"] || "",
            상위모듈: r["상위 모듈"] || "",
            모듈코드: r["모듈 코드"] || "",
            모듈명: r["모듈명"] || "",
            발화목록: scriptCols
              .map((col) => (r[col] || "").toString().trim())
              .filter(Boolean),
          }))
          .filter((m) => m.발화목록.length > 0);

        resolve(modules);
      } catch {
        reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
      }
    };
    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsArrayBuffer(file);
  });
}

function generateScript(modules: Module[], info: CustomerInfo): StageScript[] {
  const cfOrder = Object.keys(CF_LABELS);
  const grouped: Record<string, Module[]> = {};

  for (const m of modules) {
    if (!grouped[m.cf]) grouped[m.cf] = [];
    grouped[m.cf].push(m);
  }

  return cfOrder
    .filter((cf) => grouped[cf]?.length)
    .map((cf) => ({
      cf,
      lines: grouped[cf].map((m) => ({
        모듈코드: m.모듈코드,
        모듈명: m.모듈명,
        발화: applyReplacements(
          m.발화목록[Math.floor(Math.random() * m.발화목록.length)],
          info
        ),
      })),
    }));
}

export default function Home() {
  const [modules, setModules] = useState<Module[] | null>(null);
  const [fileName, setFileName] = useState("");
  const [script, setScript] = useState<StageScript[] | null>(null);
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [customerName, setCustomerName] = useState("");
  const [gender, setGender] = useState("");
  const [age, setAge] = useState("");
  const [consultantName, setConsultantName] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const scriptRef = useRef<HTMLDivElement>(null);

  const info: CustomerInfo = { customerName, gender, age, consultantName };

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      setError("엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.");
      return;
    }
    setError("");
    try {
      const parsed = await parseExcel(file);
      setModules(parsed);
      setFileName(file.name);
      setScript(null);
    } catch (e) {
      setError((e as Error).message);
    }
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  function handleGenerate() {
    if (!modules) return;
    setScript(generateScript(modules, info));
  }

  function handleCopy() {
    if (!script) return;
    const header = [
      customerName && `고객명: ${customerName}`,
      gender && `성별: ${gender}`,
      age && `나이: ${age}`,
      consultantName && `설계사: ${consultantName}`,
    ].filter(Boolean).join(" | ");

    const text = [
      header,
      "",
      ...script.map(
        (stage) =>
          `【${CF_LABELS[stage.cf] ?? stage.cf}】\n` +
          stage.lines.map((l) => `[${l.모듈명}] ${l.발화}`).join("\n")
      ),
    ].join("\n");

    navigator.clipboard.writeText(text);
    alert("스크립트가 클립보드에 복사되었습니다.");
  }

  return (
    <main className="min-h-screen py-10 px-4">
      <div className="max-w-2xl mx-auto">
        {/* 헤더 */}
        <div className="mb-8 text-center">
          <h1 className="text-2xl font-bold text-black mb-1">보험 상담 스크립트 생성기</h1>
          <p className="text-sm text-black">발화 스크립트 엑셀을 업로드하면 모듈별로 랜덤 발화를 뽑아드립니다.</p>
        </div>

        {/* 고객 & 설계사 정보 */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-300 p-6 mb-4">
          <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">상담 정보 입력</h2>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-xs font-medium text-black mb-1">고객 이름</label>
              <input
                type="text"
                placeholder="홍길동"
                value={customerName}
                onChange={(e) => setCustomerName(e.target.value)}
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-black mb-1">설계사 이름</label>
              <input
                type="text"
                placeholder="내 이름"
                value={consultantName}
                onChange={(e) => setConsultantName(e.target.value)}
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-black mb-1">성별</label>
              <select
                value={gender}
                onChange={(e) => setGender(e.target.value)}
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white"
              >
                <option value="">선택 안함</option>
                <option value="남성">남성</option>
                <option value="여성">여성</option>
              </select>
            </div>
            <div>
              <label className="block text-xs font-medium text-black mb-1">나이</label>
              <input
                type="text"
                placeholder="예: 35세"
                value={age}
                onChange={(e) => setAge(e.target.value)}
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white"
              />
            </div>
          </div>
        </div>

        {/* 파일 업로드 */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-300 p-6 mb-6">
          <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">엑셀 파일 업로드</h2>

          <div
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors
              ${dragging ? "border-blue-400 bg-blue-50" : "border-gray-300 hover:border-gray-400 hover:bg-gray-50"}`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => { if (e.target.files?.[0]) handleFile(e.target.files[0]); }}
            />
            {fileName ? (
              <div>
                <p className="text-sm font-semibold text-black mb-1">{fileName}</p>
                <p className="text-xs text-black">{modules?.length}개 모듈 로드됨 — 클릭하면 다시 업로드</p>
              </div>
            ) : (
              <div>
                <p className="text-sm text-black mb-1">엑셀 파일을 드래그하거나 클릭해서 업로드</p>
                <p className="text-xs text-black">.xlsx, .xls 지원</p>
              </div>
            )}
          </div>

          {error && <p className="mt-3 text-sm text-red-500">{error}</p>}

          <button
            onClick={handleGenerate}
            disabled={!modules}
            className="mt-5 w-full py-2.5 rounded-lg text-sm font-semibold transition-colors
              bg-blue-600 text-white hover:bg-blue-700 disabled:bg-gray-200 disabled:text-black disabled:cursor-not-allowed"
          >
            {modules ? "스크립트 생성 (랜덤)" : "파일을 먼저 업로드해주세요"}
          </button>
        </div>

        {/* 스크립트 출력 */}
        {script && (
          <>
            <div className="flex gap-2 mb-4 print:hidden">
              <button
                onClick={handleCopy}
                className="flex-1 py-2 rounded-lg border border-gray-300 text-sm font-medium text-black bg-white hover:bg-gray-100 transition-colors"
              >
                복사
              </button>
              <button
                onClick={() => window.print()}
                className="flex-1 py-2 rounded-lg border border-gray-300 text-sm font-medium text-black bg-white hover:bg-gray-100 transition-colors"
              >
                인쇄
              </button>
              <button
                onClick={handleGenerate}
                className="flex-1 py-2 rounded-lg border border-blue-300 text-sm font-medium text-blue-700 bg-white hover:bg-blue-50 transition-colors"
              >
                다시 뽑기
              </button>
            </div>

            <div ref={scriptRef} className="space-y-4">
              {/* 고객 정보 요약 */}
              {(customerName || gender || age || consultantName) && (
                <div className="bg-white rounded-xl border border-gray-200 px-5 py-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-black">
                  {customerName && <span><span className="font-semibold">고객</span> {customerName}</span>}
                  {gender && <span><span className="font-semibold">성별</span> {gender}</span>}
                  {age && <span><span className="font-semibold">나이</span> {age}</span>}
                  {consultantName && <span><span className="font-semibold">설계사</span> {consultantName}</span>}
                </div>
              )}

              {script.map((stage, i) => (
                <div key={stage.cf} className="rounded-xl border-l-4 border-blue-400 bg-blue-50 p-5">
                  <h3 className="text-sm font-bold text-black mb-3">
                    {CF_LABELS[stage.cf] ?? stage.cf}
                  </h3>
                  <ul className="space-y-3">
                    {stage.lines.map((line, j) => (
                      <li key={j}>
                        <p className="text-xs font-semibold text-black mb-0.5">{line.모듈명}</p>
                        <div className="flex gap-2 text-sm text-black leading-relaxed">
                          <span className="mt-0.5 flex-shrink-0">▶</span>
                          <span className="whitespace-pre-wrap">{line.발화}</span>
                        </div>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}

              <p className="text-xs text-center text-black pt-2">
                ※ 버튼을 누를 때마다 각 모듈에서 다른 발화가 랜덤으로 뽑힙니다.
              </p>
            </div>
          </>
        )}
      </div>
    </main>
  );
}
