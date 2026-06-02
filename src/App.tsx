import { useState, useEffect, useRef } from "react";

// ─── Types ────────────────────────────────────────────────────────────────────

interface TrademarkResult {
  serialNumber: string;
  mark: string;
  status: string;
  correspondentName: string;
  correspondentAddress: string;
  phone: string;
  email: string;
  hasAttorney: boolean;
  attorneyName: string;
  confidence: string;
  needsReview: boolean;
  error?: string;
  pending?: boolean;
}

type TabType = "all" | "attorney" | "noAttorney" | "errors";

// ─── Google Apps Script code to copy ─────────────────────────────────────────

const APPS_SCRIPT_CODE = `function doGet(e) {
  var serialNumber = e.parameter.serial;

  if (!serialNumber) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: "Missing serial number parameter"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var result = lookupTrademark(serialNumber);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Fatal error for serial " + serialNumber + ": " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      serialNumber: serialNumber,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function lookupTrademark(serialNumber) {
  var data = {
    success: true,
    serialNumber: serialNumber,
    mark: "",
    status: "",
    correspondentName: "",
    correspondentAddress: "",
    phone: "",
    email: "",
    hasAttorney: false,
    attorneyName: "",
    confidence: "low",
    needsReview: false
  };

  // ── Method 1: XML (primary – richest data source) ──────────────────────────
  try {
    var xmlUrl = "https://tsdr.uspto.gov/documentparser/" + serialNumber + "/case.xml";
    var xmlResp = UrlFetchApp.fetch(xmlUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      deadline: 20
    });

    if (xmlResp.getResponseCode() === 200) {
      var x = xmlResp.getContentText();

      // Mark / verbal element
      var markMatch = x.match(/<(?:[a-z0-9]+:)?MarkVerbalElementText>([^<]+)<\\/(?:[a-z0-9]+:)?MarkVerbalElementText>/i)
                   || x.match(/<(?:[a-z0-9]+:)?wordMarkSpecification[^>]*>[\\s\\S]*?<(?:[a-z0-9]+:)?MarkVerbalElementText>([^<]+)<\\/(?:[a-z0-9]+:)?MarkVerbalElementText>/i)
                   || x.match(/<(?:[a-z0-9]+:)?markVerbalElementText>([^<]+)<\\/(?:[a-z0-9]+:)?markVerbalElementText>/i);
      if (markMatch && markMatch[1]) data.mark = markMatch[1].trim();

      // Status
      var statusMatch = x.match(/<(?:[a-z0-9]+:)?MarkCurrentStatusExternalDescriptionText>([^<]+)<\\/(?:[a-z0-9]+:)?MarkCurrentStatusExternalDescriptionText>/i)
                     || x.match(/<(?:[a-z0-9]+:)?statusDescriptionText>([^<]+)<\\/(?:[a-z0-9]+:)?statusDescriptionText>/i);
      if (statusMatch && statusMatch[1]) data.status = statusMatch[1].trim();

      // Correspondent Name
      var corrNameMatch = x.match(/<(?:[a-z0-9]+:)?corrAddr1>([^<]+)<\\/(?:[a-z0-9]+:)?corrAddr1>/i)
                       || x.match(/<(?:[a-z0-9]+:)?correspondentName>([^<]+)<\\/(?:[a-z0-9]+:)?correspondentName>/i)
                       || x.match(/<(?:[a-z0-9]+:)?CorrAddr1>([^<]+)<\\/(?:[a-z0-9]+:)?CorrAddr1>/i);
      if (corrNameMatch && corrNameMatch[1]) data.correspondentName = corrNameMatch[1].trim();

      // Correspondent Address (lines 2-5 concatenated)
      var addrParts = [];
      var addrPatterns = [
        /<(?:[a-z0-9]+:)?corrAddr2>([^<]+)<\\/(?:[a-z0-9]+:)?corrAddr2>/i,
        /<(?:[a-z0-9]+:)?corrAddr3>([^<]+)<\\/(?:[a-z0-9]+:)?corrAddr3>/i,
        /<(?:[a-z0-9]+:)?corrAddr4>([^<]+)<\\/(?:[a-z0-9]+:)?corrAddr4>/i,
        /<(?:[a-z0-9]+:)?corrAddr5>([^<]+)<\\/(?:[a-z0-9]+:)?corrAddr5>/i
      ];
      for (var a = 0; a < addrPatterns.length; a++) {
        var am = x.match(addrPatterns[a]);
        if (am && am[1] && am[1].trim()) addrParts.push(am[1].trim());
      }
      if (addrParts.length > 0) data.correspondentAddress = addrParts.join(", ");

      // Phone
      var phoneMatch = x.match(/<(?:[a-z0-9]+:)?corrPhone>([^<]+)<\\/(?:[a-z0-9]+:)?corrPhone>/i)
                    || x.match(/<(?:[a-z0-9]+:)?PhoneNumber>([^<]+)<\\/(?:[a-z0-9]+:)?PhoneNumber>/i)
                    || x.match(/<(?:[a-z0-9]+:)?phone>([^<]+)<\\/(?:[a-z0-9]+:)?phone>/i);
      if (phoneMatch && phoneMatch[1]) data.phone = phoneMatch[1].trim();

      // Email
      var emailMatch = x.match(/<(?:[a-z0-9]+:)?corrEmail>([^<]+)<\\/(?:[a-z0-9]+:)?corrEmail>/i)
                    || x.match(/<(?:[a-z0-9]+:)?attrneyPrimaryEmailAddr>([^<]+)<\\/(?:[a-z0-9]+:)?attrneyPrimaryEmailAddr>/i)
                    || x.match(/<(?:[a-z0-9]+:)?primaryEmailAddr>([^<]+)<\\/(?:[a-z0-9]+:)?primaryEmailAddr>/i);
      if (emailMatch && emailMatch[1]) data.email = emailMatch[1].trim();

      // Attorney
      var attyPatterns = [
        /<(?:[a-z0-9]+:)?attrneyNm>([^<]+)<\\/(?:[a-z0-9]+:)?attrneyNm>/i,
        /<(?:[a-z0-9]+:)?AttorneyName>([^<]+)<\\/(?:[a-z0-9]+:)?AttorneyName>/i,
        /<(?:[a-z0-9]+:)?attorneyName>([^<]+)<\\/(?:[a-z0-9]+:)?attorneyName>/i,
        /<(?:[a-z0-9]+:)?attrnyNm>([^<]+)<\\/(?:[a-z0-9]+:)?attrnyNm>/i,
        /<(?:[a-z0-9]+:)?PrimaryAttorney>([^<]+)<\\/(?:[a-z0-9]+:)?PrimaryAttorney>/i
      ];
      for (var i = 0; i < attyPatterns.length; i++) {
        var m = x.match(attyPatterns[i]);
        if (m && m[1] && m[1].trim() !== "") {
          data.hasAttorney = true;
          data.attorneyName = m[1].trim();
          data.confidence = "high";
          break;
        }
      }

      // Explicit no-attorney
      if (!data.hasAttorney) {
        if (x.match(/<(?:[a-z0-9]+:)?attrneyNm>\\s*<\\/(?:[a-z0-9]+:)?attrneyNm>/i)) {
          data.confidence = "high";
        }
      }
    }
  } catch (e1) {
    Logger.log("XML method failed: " + e1.toString());
  }

  // ── Method 2: Status JSON ──────────────────────────────────────────────────
  try {
    var jsonUrl = "https://tsdrapi.uspto.gov/ts/cd/casestatus/" + serialNumber + "/info.json";
    var jsonResp = UrlFetchApp.fetch(jsonUrl, {
      muteHttpExceptions: true,
      deadline: 15,
      headers: { "Accept": "application/json" }
    });

    if (jsonResp.getResponseCode() === 200) {
      var jt = jsonResp.getContentText();

      // Mark (fill if missing)
      if (!data.mark) {
        var jMark = jt.match(/"markVerbalElementText"\\s*:\\s*"([^"]+)"/i)
                 || jt.match(/"MarkVerbalElementText"\\s*:\\s*"([^"]+)"/i)
                 || jt.match(/"wordMark"\\s*:\\s*"([^"]+)"/i);
        if (jMark && jMark[1]) data.mark = jMark[1].trim();
      }

      // Status
      if (!data.status) {
        var jStatus = jt.match(/"markCurrentStatusExternalDescriptionText"\\s*:\\s*"([^"]+)"/i)
                   || jt.match(/"statusDescriptionText"\\s*:\\s*"([^"]+)"/i);
        if (jStatus && jStatus[1]) data.status = jStatus[1].trim();
      }

      // Correspondent Name
      if (!data.correspondentName) {
        var jCorrName = jt.match(/"corrAddr1"\\s*:\\s*"([^"]+)"/i)
                     || jt.match(/"correspondentName"\\s*:\\s*"([^"]+)"/i);
        if (jCorrName && jCorrName[1]) data.correspondentName = jCorrName[1].trim();
      }

      // Correspondent Address
      if (!data.correspondentAddress) {
        var jAddrParts = [];
        var jAddrPatterns = [
          /"corrAddr2"\\s*:\\s*"([^"]+)"/i,
          /"corrAddr3"\\s*:\\s*"([^"]+)"/i,
          /"corrAddr4"\\s*:\\s*"([^"]+)"/i,
          /"corrAddr5"\\s*:\\s*"([^"]+)"/i
        ];
        for (var b = 0; b < jAddrPatterns.length; b++) {
          var bm = jt.match(jAddrPatterns[b]);
          if (bm && bm[1] && bm[1].trim()) jAddrParts.push(bm[1].trim());
        }
        if (jAddrParts.length > 0) data.correspondentAddress = jAddrParts.join(", ");
      }

      // Phone
      if (!data.phone) {
        var jPhone = jt.match(/"corrPhone"\\s*:\\s*"([^"]+)"/i)
                  || jt.match(/"phone"\\s*:\\s*"([^"]+)"/i);
        if (jPhone && jPhone[1]) data.phone = jPhone[1].trim();
      }

      // Email
      if (!data.email) {
        var jEmail = jt.match(/"corrEmail"\\s*:\\s*"([^"]+)"/i)
                  || jt.match(/"attrneyPrimaryEmailAddr"\\s*:\\s*"([^"]+)"/i)
                  || jt.match(/"primaryEmailAddr"\\s*:\\s*"([^"]+)"/i);
        if (jEmail && jEmail[1] && jEmail[1].toLowerCase() !== "null") data.email = jEmail[1].trim();
      }

      // Attorney
      if (!data.hasAttorney) {
        var jAttyPatterns = [
          /"attrneyNm"\\s*:\\s*"([^"]+)"/,
          /"attorneyName"\\s*:\\s*"([^"]+)"/,
          /"AttorneyName"\\s*:\\s*"([^"]+)"/,
          /"attrnyNm"\\s*:\\s*"([^"]+)"/
        ];
        for (var j = 0; j < jAttyPatterns.length; j++) {
          var jm = jt.match(jAttyPatterns[j]);
          if (jm && jm[1] && jm[1].trim() !== "" && jm[1].trim().toLowerCase() !== "null") {
            data.hasAttorney = true;
            data.attorneyName = jm[1].trim();
            data.confidence = "high";
            break;
          }
        }
      }
    }
  } catch (e2) {
    Logger.log("JSON method failed: " + e2.toString());
  }

  // ── Method 3: HTML status page ─────────────────────────────────────────────
  try {
    var htmlUrl = "https://tsdr.uspto.gov/statusview/" + serialNumber;
    var htmlResp = UrlFetchApp.fetch(htmlUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      deadline: 15
    });

    if (htmlResp.getResponseCode() === 200) {
      var h = htmlResp.getContentText();

      // Mark
      if (!data.mark) {
        var hMark = h.match(/Mark:[\\s\\S]*?<[^>]+>([A-Z0-9 &'"\\.\\-]{2,60})<\\//)
                 || h.match(/Trademark:\\s*([\\w\\s&'".\\-]{2,60})/i);
        if (hMark && hMark[1]) data.mark = hMark[1].trim();
      }

      // Status
      if (!data.status) {
        var hStatus = h.match(/Status:\\s*([^<\\n]{5,120})/i);
        if (hStatus && hStatus[1]) data.status = hStatus[1].trim();
      }

      // Phone
      if (!data.phone) {
        var hPhone = h.match(/(?:Phone|Tel)[^:]*:\\s*([\\d\\(\\)\\-\\+\\. ]{7,20})/i);
        if (hPhone && hPhone[1]) data.phone = hPhone[1].trim();
      }

      // Email
      if (!data.email) {
        var hEmail = h.match(/[a-zA-Z0-9._%+\\-]+@[a-zA-Z0-9.\\-]+\\.[a-zA-Z]{2,}/);
        if (hEmail) data.email = hEmail[0];
      }

      // Attorney
      if (!data.hasAttorney) {
        if (h.indexOf("Attorney of Record - None") > -1) {
          data.confidence = "high"; // explicitly no attorney
        } else {
          var hAttyPatterns = [
            /Attorney(?:\\s+of\\s+Record)?(?:\\s+Name)?:\\s*([\\w\\xC0-\\xFF][^<\\n\\r]{2,})/i,
            /Correspondent(?:\\s*\\/\\s*Attorney)?:\\s*([\\w\\xC0-\\xFF][^<\\n\\r]{2,})/i
          ];
          for (var k = 0; k < hAttyPatterns.length; k++) {
            var hm = h.match(hAttyPatterns[k]);
            if (hm && hm[1] && hm[1].trim() !== "") {
              data.hasAttorney = true;
              if (!data.attorneyName) data.attorneyName = hm[1].trim();
              data.confidence = "high";
              break;
            }
          }
          if (!data.hasAttorney &&
              h.indexOf("Attorney of Record") > -1 &&
              h.indexOf("Attorney Primary Email") > -1) {
            data.hasAttorney = true;
            data.confidence = "medium";
          }
        }
      }
    }
  } catch (e3) {
    Logger.log("HTML method failed: " + e3.toString());
  }

  // Final confidence adjustment
  if (data.confidence === "low" && (data.mark || data.status)) {
    data.confidence = "medium";
  }
  if (!data.mark && !data.status && !data.correspondentName) {
    data.needsReview = true;
  }

  return data;
}

function doPost(e) { return doGet(e); }`;

// ─── Helper ───────────────────────────────────────────────────────────────────

const getTsdrLink = (serial: string) =>
  `https://tsdr.uspto.gov/#caseNumber=${serial}&caseSearchType=US_APPLICATION&caseType=DEFAULT&searchType=statusSearch`;

const parseSerials = (input: string): string[] =>
  input
    .split(/[\s,;\n]+/)
    .map((s) => s.trim().replace(/\D/g, ""))
    .filter((s) => s.length >= 7 && s.length <= 9);

// ─── Sub-components ───────────────────────────────────────────────────────────

function Badge({ value }: { value: string | boolean | null }) {
  if (value === true || value === "true")
    return (
      <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/20 px-2 py-0.5 text-xs font-semibold text-emerald-400 ring-1 ring-emerald-500/40">
        ✓ Yes
      </span>
    );
  if (value === false || value === "false")
    return (
      <span className="inline-flex items-center gap-1 rounded-full bg-amber-500/20 px-2 py-0.5 text-xs font-semibold text-amber-400 ring-1 ring-amber-500/40">
        ✗ No
      </span>
    );
  return (
    <span className="inline-flex items-center gap-1 rounded-full bg-slate-500/20 px-2 py-0.5 text-xs font-semibold text-slate-400 ring-1 ring-slate-500/40">
      ? Unknown
    </span>
  );
}

function ConfidencePill({ level }: { level: string }) {
  const map: Record<string, string> = {
    high: "bg-emerald-500/20 text-emerald-400 ring-emerald-500/40",
    medium: "bg-yellow-500/20 text-yellow-400 ring-yellow-500/40",
    low: "bg-red-500/20 text-red-400 ring-red-500/40",
  };
  return (
    <span
      className={`inline-flex rounded-full px-2 py-0.5 text-xs font-semibold ring-1 ${map[level] ?? map.low}`}
    >
      {level}
    </span>
  );
}

function CopyButton({ text, label }: { text: string; label: string }) {
  const [copied, setCopied] = useState(false);
  const copy = () => {
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };
  return (
    <button
      onClick={copy}
      className="flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white transition hover:bg-blue-500 active:scale-95"
    >
      {copied ? (
        <>
          <span>✅</span> Copied!
        </>
      ) : (
        <>
          <span>📋</span> {label}
        </>
      )}
    </button>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [step, setStep] = useState<1 | 2>(1);
  const [scriptUrl, setScriptUrl] = useState("");
  const [isTesting, setIsTesting] = useState(false);
  const [connectionError, setConnectionError] = useState("");
  const [serialInput, setSerialInput] = useState("");
  const [results, setResults] = useState<TrademarkResult[]>([]);
  const [isChecking, setIsChecking] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentSerial, setCurrentSerial] = useState("");
  const [activeTab, setActiveTab] = useState<TabType>("all");
  const [searchFilter, setSearchFilter] = useState("");
  const abortRef = useRef(false);

  useEffect(() => {
    const saved = localStorage.getItem("ladley_script_url");
    if (saved) {
      setScriptUrl(saved);
      setStep(2);
    }
  }, []);

  // ── Connection test ──
  const testConnection = async () => {
    if (!scriptUrl.trim()) {
      setConnectionError("Please enter your Google Apps Script Web App URL.");
      return;
    }
    setIsTesting(true);
    setConnectionError("");
    try {
      const url = scriptUrl.trim() + "?serial=97119270";
      const res = await fetch(url);
      const data = await res.json();
      if (data.success !== undefined || data.serialNumber !== undefined) {
        localStorage.setItem("ladley_script_url", scriptUrl.trim());
        setStep(2);
      } else {
        setConnectionError("Script responded but returned unexpected data.");
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setConnectionError("Connection failed: " + msg);
    }
    setIsTesting(false);
  };

  // ── Run lookup ──
  const runLookup = async () => {
    const serials = parseSerials(serialInput);
    if (!serials.length) {
      alert("Please enter at least one valid serial number (7-9 digits).");
      return;
    }
    abortRef.current = false;
    setIsChecking(true);
    setProgress(0);
    setResults([]);
    setActiveTab("all");

    // Pre-populate rows as pending
    const initial: TrademarkResult[] = serials.map((s) => ({
      serialNumber: s,
      mark: "",
      status: "",
      correspondentName: "",
      correspondentAddress: "",
      phone: "",
      email: "",
      hasAttorney: false,
      attorneyName: "",
      confidence: "low",
      needsReview: false,
      pending: true,
    }));
    setResults(initial);

    const updated = [...initial];

    for (let i = 0; i < serials.length; i++) {
      if (abortRef.current) break;
      const serial = serials[i];
      setCurrentSerial(serial);
      setProgress(Math.round(((i + 1) / serials.length) * 100));

      try {
        const res = await fetch(`${scriptUrl.trim()}?serial=${serial}`);
        const data = await res.json();

        updated[i] = {
          serialNumber: serial,
          mark: data.mark || "",
          status: data.status || "",
          correspondentName: data.correspondentName || "",
          correspondentAddress: data.correspondentAddress || "",
          phone: data.phone || "",
          email: data.email || "",
          hasAttorney: !!data.hasAttorney,
          attorneyName: data.attorneyName || "",
          confidence: data.confidence || "low",
          needsReview: !!data.needsReview,
          error: data.error || undefined,
          pending: false,
        };
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        updated[i] = {
          ...updated[i],
          error: msg,
          pending: false,
        };
      }

      setResults([...updated]);

      if (i < serials.length - 1) {
        await new Promise((r) => setTimeout(r, 600));
      }
    }

    setIsChecking(false);
    setCurrentSerial("");
  };

  const stopLookup = () => {
    abortRef.current = true;
    setIsChecking(false);
    setCurrentSerial("");
  };

  // ── CSV Export ──
  const exportCSV = (rows: TrademarkResult[], filename: string) => {
    if (!rows.length) { alert("No data to export."); return; }
    const headers = [
      "Serial Number",
      "Mark",
      "Status",
      "Has Attorney",
      "Attorney Name",
      "Correspondent Name",
      "Correspondent Address",
      "Phone",
      "Correspondent Email",
      "Confidence",
      "Needs Review",
      "TSDR Link",
    ];
    const escape = (v: string) => `"${v.replace(/"/g, '""')}"`;
    const rowLines = rows.map((r) =>
      [
        r.serialNumber,
        escape(r.mark),
        escape(r.status),
        r.hasAttorney ? "Yes" : "No",
        escape(r.attorneyName),
        escape(r.correspondentName),
        escape(r.correspondentAddress),
        escape(r.phone),
        escape(r.email),
        r.confidence,
        r.needsReview ? "Yes" : "No",
        getTsdrLink(r.serialNumber),
      ].join(",")
    );
    const csv = [headers.join(","), ...rowLines].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  };

  // ── Filtered views ──
  const withAttorney = results.filter((r) => !r.pending && !r.error && r.hasAttorney);
  const noAttorney = results.filter((r) => !r.pending && !r.error && !r.hasAttorney);
  const errors = results.filter((r) => !r.pending && r.error);
  const pending = results.filter((r) => r.pending);

  const filterText = searchFilter.toLowerCase();
  const filteredRows = (
    activeTab === "all"
      ? results.filter((r) => !r.pending && !r.error)
      : activeTab === "attorney"
      ? withAttorney
      : activeTab === "noAttorney"
      ? noAttorney
      : errors
  ).filter(
    (r) =>
      !filterText ||
      r.serialNumber.includes(filterText) ||
      r.mark.toLowerCase().includes(filterText) ||
      r.correspondentName.toLowerCase().includes(filterText) ||
      r.email.toLowerCase().includes(filterText) ||
      r.status.toLowerCase().includes(filterText)
  );

  const tabClass = (t: TabType) =>
    `px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
      activeTab === t
        ? "bg-slate-800 text-white border-b-2 border-blue-500"
        : "bg-slate-900 text-slate-400 hover:text-white"
    }`;

  // ─── Step 1: Setup ──────────────────────────────────────────────────────────
  if (step === 1) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-slate-950 text-white">
        <div className="mx-auto max-w-3xl px-4 py-10">
          {/* Header */}
          <div className="mb-10 text-center">
            <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-blue-500/30 bg-blue-500/10 px-4 py-1.5 text-sm text-blue-400">
              ⚖️ USPTO Trademark Intelligence
            </div>
            <h1 className="text-4xl font-extrabold tracking-tight text-white">
              LADLEY Trademark Lookup
            </h1>
            <p className="mt-2 text-slate-400">
              Bulk-check USPTO serial numbers for mark details, correspondent contact info &amp; attorney status.
            </p>
          </div>

          {/* Setup card */}
          <div className="rounded-2xl border border-slate-700 bg-slate-800/60 p-8 shadow-2xl backdrop-blur">
            <h2 className="mb-6 flex items-center gap-2 text-xl font-bold text-yellow-400">
              <span>⚙️</span> One-Time Setup
            </h2>

            <div className="space-y-6">
              {/* Step 1 */}
              <div className="rounded-xl border border-slate-600 bg-slate-900/60 p-5">
                <h3 className="mb-3 flex items-center gap-2 font-bold text-blue-300">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-blue-600 text-xs font-bold">1</span>
                  Create a Google Apps Script
                </h3>
                <ol className="mb-4 space-y-1.5 pl-4 text-sm text-slate-300 list-decimal">
                  <li>
                    Go to{" "}
                    <a
                      href="https://script.google.com"
                      target="_blank"
                      rel="noreferrer"
                      className="text-blue-400 underline hover:text-blue-300"
                    >
                      script.google.com
                    </a>{" "}
                    and click <strong>"New project"</strong>
                  </li>
                  <li>Delete all existing code in the editor</li>
                  <li>Paste the script below (click Copy Code)</li>
                  <li>Click <strong>Save</strong> (Ctrl+S / ⌘+S)</li>
                </ol>
                <CopyButton text={APPS_SCRIPT_CODE} label="Copy Apps Script Code" />
              </div>

              {/* Step 2 */}
              <div className="rounded-xl border border-slate-600 bg-slate-900/60 p-5">
                <h3 className="mb-3 flex items-center gap-2 font-bold text-blue-300">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-blue-600 text-xs font-bold">2</span>
                  Deploy as a Web App
                </h3>
                <ol className="space-y-1.5 pl-4 text-sm text-slate-300 list-decimal">
                  <li>Click <strong>Deploy → New deployment</strong></li>
                  <li>Click the ⚙️ gear icon → select <strong>Web app</strong></li>
                  <li>
                    Set <strong>"Who has access"</strong> to{" "}
                    <strong className="text-yellow-400">Anyone</strong>
                  </li>
                  <li>Click <strong>Deploy</strong> and authorize when prompted</li>
                  <li>Copy the <strong>Web App URL</strong> shown at the end</li>
                </ol>
              </div>

              {/* Step 3 */}
              <div className="rounded-xl border border-slate-600 bg-slate-900/60 p-5">
                <h3 className="mb-3 flex items-center gap-2 font-bold text-blue-300">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-blue-600 text-xs font-bold">3</span>
                  Paste URL &amp; Connect
                </h3>
                <input
                  type="text"
                  value={scriptUrl}
                  onChange={(e) => setScriptUrl(e.target.value)}
                  placeholder="https://script.google.com/macros/s/…/exec"
                  className="mb-3 w-full rounded-lg border border-slate-600 bg-slate-800 px-4 py-3 text-sm text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
                />
                {connectionError && (
                  <div className="mb-3 rounded-lg border border-red-500/40 bg-red-500/10 px-4 py-2 text-sm text-red-400">
                    ❌ {connectionError}
                  </div>
                )}
                <button
                  onClick={testConnection}
                  disabled={isTesting || !scriptUrl.trim()}
                  className="flex items-center gap-2 rounded-lg bg-emerald-600 px-6 py-2.5 text-sm font-bold text-white transition hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  {isTesting ? (
                    <>
                      <span className="inline-block h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                      Testing…
                    </>
                  ) : (
                    "🔗 Test Connection & Continue"
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ─── Step 2: Lookup Tool ────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-slate-950 text-white">
      <div className="mx-auto max-w-7xl px-4 py-8">

        {/* Header bar */}
        <div className="mb-8 flex flex-wrap items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl font-extrabold tracking-tight text-white">
              ⚖️ LADLEY Trademark Lookup
            </h1>
            <p className="text-sm text-slate-400">
              USPTO TSDR Bulk Intelligence Tool — Mark · Contact · Status · Attorney
            </p>
          </div>
          <div className="flex items-center gap-3">
            <span className="flex items-center gap-1.5 rounded-full border border-emerald-500/30 bg-emerald-500/10 px-3 py-1 text-xs text-emerald-400">
              <span className="h-2 w-2 animate-pulse rounded-full bg-emerald-400" />
              Connected
            </span>
            <button
              onClick={() => {
                localStorage.removeItem("ladley_script_url");
                setStep(1);
                setResults([]);
              }}
              className="rounded-lg border border-slate-600 px-3 py-1.5 text-xs text-slate-400 transition hover:border-slate-400 hover:text-white"
            >
              ⚙️ Change Script
            </button>
          </div>
        </div>

        {/* Input area */}
        <div className="mb-6 rounded-2xl border border-slate-700 bg-slate-800/60 p-6 shadow-xl backdrop-blur">
          <label className="mb-2 block text-sm font-semibold text-slate-300">
            USPTO Serial Numbers
          </label>
          <textarea
            value={serialInput}
            onChange={(e) => setSerialInput(e.target.value)}
            placeholder={"Paste serial numbers separated by commas, spaces, or new lines.\nExample:\n98765432\n97654321, 96543210\n95432109"}
            rows={5}
            className="w-full resize-none rounded-xl border border-slate-600 bg-slate-900 px-4 py-3 font-mono text-sm text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
          />
          <div className="mt-4 flex flex-wrap items-center gap-3">
            {!isChecking ? (
              <button
                onClick={runLookup}
                disabled={!serialInput.trim()}
                className="flex items-center gap-2 rounded-xl bg-blue-600 px-6 py-3 text-sm font-bold text-white transition hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-50"
              >
                🔍 Lookup Trademarks
              </button>
            ) : (
              <button
                onClick={stopLookup}
                className="flex items-center gap-2 rounded-xl bg-red-600 px-6 py-3 text-sm font-bold text-white transition hover:bg-red-500"
              >
                ⏹ Stop
              </button>
            )}
            {results.length > 0 && !isChecking && (
              <button
                onClick={() => { setResults([]); setSerialInput(""); setActiveTab("all"); }}
                className="rounded-xl border border-slate-600 px-4 py-3 text-sm text-slate-400 transition hover:border-slate-400 hover:text-white"
              >
                Clear Results
              </button>
            )}
            {serialInput.trim() && (
              <span className="text-sm text-slate-500">
                {parseSerials(serialInput).length} serial number(s) detected
              </span>
            )}
          </div>
        </div>

        {/* Progress bar */}
        {isChecking && (
          <div className="mb-6 rounded-2xl border border-slate-700 bg-slate-800/60 p-5">
            <div className="mb-2 flex justify-between text-sm">
              <span className="text-slate-300">
                Checking{" "}
                <span className="font-mono text-blue-400">{currentSerial}</span>
                …
              </span>
              <span className="text-slate-400">{progress}%</span>
            </div>
            <div className="h-2.5 w-full overflow-hidden rounded-full bg-slate-700">
              <div
                className="h-full rounded-full bg-gradient-to-r from-blue-500 to-violet-500 transition-all duration-300"
                style={{ width: `${progress}%` }}
              />
            </div>
            {pending.length > 0 && (
              <p className="mt-2 text-xs text-slate-500">
                {results.filter((r) => !r.pending).length} / {results.length} completed — {pending.length} queued
              </p>
            )}
          </div>
        )}

        {/* Summary cards */}
        {results.some((r) => !r.pending) && (
          <div className="mb-6 grid grid-cols-2 gap-4 sm:grid-cols-4">
            <div className="rounded-2xl border border-slate-700 bg-slate-800/60 p-4 text-center">
              <div className="text-3xl font-black text-white">{results.filter((r) => !r.pending && !r.error).length}</div>
              <div className="mt-1 text-xs text-slate-400">Total Checked</div>
            </div>
            <div className="rounded-2xl border border-emerald-500/30 bg-emerald-500/10 p-4 text-center">
              <div className="text-3xl font-black text-emerald-400">{withAttorney.length}</div>
              <div className="mt-1 text-xs text-emerald-400/70">Has Attorney</div>
            </div>
            <div className="rounded-2xl border border-amber-500/30 bg-amber-500/10 p-4 text-center">
              <div className="text-3xl font-black text-amber-400">{noAttorney.length}</div>
              <div className="mt-1 text-xs text-amber-400/70">No Attorney</div>
            </div>
            <div className="rounded-2xl border border-red-500/30 bg-red-500/10 p-4 text-center">
              <div className="text-3xl font-black text-red-400">{errors.length}</div>
              <div className="mt-1 text-xs text-red-400/70">Errors</div>
            </div>
          </div>
        )}

        {/* Results table */}
        {results.length > 0 && (
          <div className="rounded-2xl border border-slate-700 bg-slate-800/60 shadow-2xl backdrop-blur">

            {/* Tabs + actions bar */}
            <div className="flex flex-wrap items-end justify-between gap-2 border-b border-slate-700 px-4 pt-4">
              <div className="flex gap-1">
                <button className={tabClass("all")} onClick={() => setActiveTab("all")}>
                  All Results{" "}
                  <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">
                    {results.filter((r) => !r.pending && !r.error).length}
                  </span>
                </button>
                <button className={tabClass("attorney")} onClick={() => setActiveTab("attorney")}>
                  ✓ Has Attorney{" "}
                  <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{withAttorney.length}</span>
                </button>
                <button className={tabClass("noAttorney")} onClick={() => setActiveTab("noAttorney")}>
                  ✗ No Attorney{" "}
                  <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{noAttorney.length}</span>
                </button>
                {errors.length > 0 && (
                  <button className={tabClass("errors")} onClick={() => setActiveTab("errors")}>
                    ⚠ Errors{" "}
                    <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{errors.length}</span>
                  </button>
                )}
              </div>
              <div className="flex flex-wrap items-center gap-2 pb-1">
                <input
                  type="text"
                  value={searchFilter}
                  onChange={(e) => setSearchFilter(e.target.value)}
                  placeholder="Filter results…"
                  className="rounded-lg border border-slate-600 bg-slate-900 px-3 py-1.5 text-xs text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
                />
                <button
                  onClick={() =>
                    exportCSV(
                      activeTab === "attorney"
                        ? withAttorney
                        : activeTab === "noAttorney"
                        ? noAttorney
                        : results.filter((r) => !r.pending && !r.error),
                      activeTab === "noAttorney"
                        ? "no-attorney-prospects.csv"
                        : activeTab === "attorney"
                        ? "has-attorney.csv"
                        : "all-trademark-results.csv"
                    )
                  }
                  className="flex items-center gap-1.5 rounded-lg bg-slate-700 px-3 py-1.5 text-xs font-medium text-white transition hover:bg-slate-600"
                >
                  ⬇ Export CSV
                </button>
                <button
                  onClick={() => exportCSV(noAttorney, "no-attorney-prospects.csv")}
                  className="flex items-center gap-1.5 rounded-lg bg-amber-600 px-3 py-1.5 text-xs font-medium text-white transition hover:bg-amber-500"
                >
                  ⬇ No-Attorney CSV
                </button>
              </div>
            </div>

            {/* Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-700 text-left text-xs font-semibold uppercase tracking-wider text-slate-400">
                    <th className="px-4 py-3 whitespace-nowrap">Serial #</th>
                    <th className="px-4 py-3 whitespace-nowrap">Mark</th>
                    <th className="px-4 py-3 whitespace-nowrap">Status</th>
                    <th className="px-4 py-3 whitespace-nowrap">Correspondent Name</th>
                    <th className="px-4 py-3 whitespace-nowrap">Correspondent Address</th>
                    <th className="px-4 py-3 whitespace-nowrap">Phone</th>
                    <th className="px-4 py-3 whitespace-nowrap">Email</th>
                    <th className="px-4 py-3 whitespace-nowrap">Attorney?</th>
                    <th className="px-4 py-3 whitespace-nowrap">Attorney Name</th>
                    <th className="px-4 py-3 whitespace-nowrap">Confidence</th>
                    <th className="px-4 py-3 whitespace-nowrap">TSDR</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/50">
                  {/* Pending rows */}
                  {activeTab === "all" &&
                    results
                      .filter((r) => r.pending)
                      .map((r) => (
                        <tr key={r.serialNumber + "-pending"} className="animate-pulse bg-slate-800/30">
                          <td className="px-4 py-3 font-mono text-slate-500">{r.serialNumber}</td>
                          {Array.from({ length: 9 }).map((_, i) => (
                            <td key={i} className="px-4 py-3">
                              <div className="h-4 w-24 rounded bg-slate-700" />
                            </td>
                          ))}
                        </tr>
                      ))}

                  {/* Result rows */}
                  {filteredRows.map((r) => (
                    <tr
                      key={r.serialNumber}
                      className={`transition-colors hover:bg-slate-700/30 ${
                        r.error
                          ? "bg-red-500/5"
                          : r.hasAttorney
                          ? "bg-emerald-500/5"
                          : "bg-amber-500/5"
                      }`}
                    >
                      <td className="px-4 py-3 font-mono text-xs text-slate-300 whitespace-nowrap">
                        {r.serialNumber}
                        {r.needsReview && (
                          <span className="ml-1 text-yellow-500" title="Needs manual review">⚑</span>
                        )}
                      </td>
                      <td className="px-4 py-3 max-w-[160px]">
                        <span
                          className="font-semibold text-white"
                          title={r.mark}
                        >
                          {r.mark || <span className="italic text-slate-500">—</span>}
                        </span>
                      </td>
                      <td className="px-4 py-3 max-w-[200px]">
                        {r.error ? (
                          <span className="text-xs text-red-400">{r.error}</span>
                        ) : (
                          <span className="text-xs text-slate-300" title={r.status}>
                            {r.status
                              ? r.status.length > 60
                                ? r.status.slice(0, 57) + "…"
                                : r.status
                              : <span className="italic text-slate-500">—</span>}
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-3 max-w-[200px]">
                        <div className="text-xs text-slate-200 font-medium">
                          {r.correspondentName || <span className="italic text-slate-500">—</span>}
                        </div>
                      </td>
                      <td className="px-4 py-3 max-w-[220px]">
                        <div className="text-xs text-slate-400">
                          {r.correspondentAddress || <span className="italic text-slate-500">—</span>}
                        </div>
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-xs text-slate-300">
                        {r.phone || <span className="italic text-slate-500">—</span>}
                      </td>
                      <td className="px-4 py-3 max-w-[200px]">
                        {r.email ? (
                          <a
                            href={`mailto:${r.email}`}
                            className="text-xs text-blue-400 hover:text-blue-300 hover:underline break-all"
                          >
                            {r.email}
                          </a>
                        ) : (
                          <span className="italic text-xs text-slate-500">—</span>
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        <Badge value={r.error ? null : r.hasAttorney} />
                      </td>
                      <td className="px-4 py-3 max-w-[180px] text-xs text-slate-300">
                        {r.attorneyName || <span className="italic text-slate-500">—</span>}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        {r.error ? (
                          <span className="text-xs text-slate-500">—</span>
                        ) : (
                          <ConfidencePill level={r.confidence} />
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        <a
                          href={getTsdrLink(r.serialNumber)}
                          target="_blank"
                          rel="noreferrer"
                          className="text-xs text-blue-400 hover:text-blue-300 hover:underline"
                        >
                          View →
                        </a>
                      </td>
                    </tr>
                  ))}

                  {filteredRows.length === 0 && !isChecking && (
                    <tr>
                      <td colSpan={11} className="py-12 text-center text-slate-500">
                        {searchFilter ? "No results match your filter." : "No results in this category yet."}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Footer summary */}
            {filteredRows.length > 0 && (
              <div className="border-t border-slate-700 px-4 py-3 text-xs text-slate-500">
                Showing {filteredRows.length} row(s)
                {searchFilter && ` matching "${searchFilter}"`}
                {" · "}
                <button
                  onClick={() =>
                    exportCSV(filteredRows, "trademark-export.csv")
                  }
                  className="text-blue-400 hover:underline"
                >
                  Export this view as CSV
                </button>
              </div>
            )}
          </div>
        )}

        {/* Empty state */}
        {results.length === 0 && !isChecking && (
          <div className="rounded-2xl border border-dashed border-slate-700 p-16 text-center">
            <div className="mb-3 text-5xl">🔎</div>
            <p className="text-slate-400">
              Enter USPTO serial numbers above and click{" "}
              <strong className="text-white">Lookup Trademarks</strong> to begin.
            </p>
            <p className="mt-2 text-sm text-slate-500">
              Retrieves: Mark name · Status · Correspondent name &amp; address · Phone · Email · Attorney flag
            </p>
          </div>
        )}

        {/* Footer note */}
        <p className="mt-6 text-center text-xs text-slate-600">
          Data sourced from USPTO TSDR via your Google Apps Script proxy · For research &amp; outreach purposes only
        </p>
      </div>
    </div>
  );
}
