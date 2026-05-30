import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Term = {
  en: string;
  kr: string;
  comment: string;
};

function App() {
  const [query, setQuery] = useState("");
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [terms, setTerms] = useState<Term[]>([]);
  const [filteredResults, setFilteredResults] = useState<Term[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [fileName, setFileName] = useState("");

  const inputRef = useRef<HTMLInputElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const saved = localStorage.getItem("terms");
    if (saved) {
      const parsed = JSON.parse(saved);
      setTerms(parsed);
      setFilteredResults(parsed.slice(0, 20));
      setFileName("loaded");
    }
  }, []);

  useEffect(() => {
    const trimmedQuery = query.trim().toLowerCase();

    if (trimmedQuery === "") {
      setFilteredResults(terms.slice(0, 20));
      setSelectedIndex(0);
      return;
    }

    const ranked = [...terms]
      .filter((item) =>
        item.en.toLowerCase().includes(trimmedQuery) ||
        item.kr.toLowerCase().includes(trimmedQuery) ||
        item.comment.toLowerCase().includes(trimmedQuery)
      )
      .sort((a, b) => {
        const getRank = (item: Term) => {
          const en = item.en.toLowerCase();
          const kr = item.kr.toLowerCase();
          const comment = item.comment.toLowerCase();
          const q = trimmedQuery;
          const abbreviationPattern = new RegExp(`\\(${q}\\)`, "i");

          const acronym = item.en
            .split(" ")
            .map((word) => word[0])
            .join("")
            .toLowerCase();

          if (abbreviationPattern.test(item.en)) return 1;
          if (acronym === q) return 2;
          if (en === q) return 3;
          if (en.startsWith(q)) return 4;
          if (kr.startsWith(q)) return 5;
          if (en.includes(q)) return 6;
          if (kr.includes(q)) return 7;
          if (comment.includes(q)) return 8;
          return 9;
        };

        return getRank(a) - getRank(b) || a.en.localeCompare(b.en);
      });

    setFilteredResults(ranked.slice(0, 20));
    setSelectedIndex(0);
  }, [query, terms]);

  const processFile = async (file: File) => {
    setFileName(file.name);

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(sheet);

    const parsedTerms: Term[] = rows
      .map((row) => ({
        en: row.en || row.EN || row.English || row.english || "",
        kr: row.kr || row.KR || row.Korean || row.korean || "",
        comment: row.comment || row.Comment || row.comments || "",
      }))
      .filter((term) => term.en || term.kr || term.comment);

    localStorage.setItem("terms", JSON.stringify(parsedTerms));
    setTerms(parsedTerms);
    setFilteredResults(parsedTerms.slice(0, 20));
    setQuery("");
    setSelectedIndex(0);
    inputRef.current?.focus();
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    await processFile(file);
    e.target.value = "";
  };

  const clearGlossary = () => {
    localStorage.removeItem("terms");
    setTerms([]);
    setFilteredResults([]);
    setQuery("");
    setSelectedIndex(0);
    setFileName("");
    inputRef.current?.focus();
  };

  useEffect(() => {
    const handleKeyDown = async (e: KeyboardEvent) => {
      if (filteredResults.length === 0) return;

      if (e.key === "ArrowDown") {
        e.preventDefault();
        setSelectedIndex((prev) =>
          Math.min(prev + 1, filteredResults.length - 1)
        );
      }

      if (e.key === "ArrowUp") {
        e.preventDefault();
        setSelectedIndex((prev) => Math.max(prev - 1, 0));
      }

      if (e.key === "Enter") {
        e.preventDefault();
        const selected = filteredResults[selectedIndex];

        if (selected) {
          await navigator.clipboard.writeText(`${selected.en} → ${selected.kr}`);
        }
      }

      if (e.key === "Escape") {
        setQuery("");
        setSelectedIndex(0);
        inputRef.current?.focus();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [filteredResults, selectedIndex]);

 return (
  <div
    onDragOver={(e) => {
      e.preventDefault();
      setIsDragging(true);
    }}
    onDragLeave={() => setIsDragging(false)}
    onDrop={async (e) => {
      e.preventDefault();
      setIsDragging(false);

      const file = e.dataTransfer.files?.[0];
      if (!file) return;

      await processFile(file);
    }}
    style={{
      minHeight: "100vh",
      background: isDragging ? "#111827" : "#0a0a0a",
      color: "white",
      padding: "40px",
      fontFamily: "sans-serif",
      border: isDragging ? "2px dashed #3b82f6" : "2px solid transparent",
      transition: "background 0.15s ease, border 0.15s ease",
    }}
  >
    <div
      style={{
        maxWidth: "900px",
        margin: "0 auto",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          gap: "16px",
          marginBottom: "8px",
        }}
      >
        <h1
          style={{
            fontSize: "22px",
            fontWeight: "bold",
            margin: 0,
          }}
        >
          Terminology Search
        </h1>

        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "flex-end",
            gap: "4px",
          }}
        >
          <div
            style={{
              color: fileName ? "#60a5fa" : "#777",
              fontSize: "13px",
            }}
          >
            {fileName
              ? "Glossary loaded ✓"
              : "No glossary loaded"}
          </div>

          {terms.length > 0 && (
            <>
              <button
                onClick={clearGlossary}
                style={{
                  marginTop: "4px",
                  padding: "8px 12px",
                  background: "#171717",
                  border: "1px solid #333",
                  borderRadius: "8px",
                  color: "#aaa",
                  cursor: "pointer",
                  fontSize: "13px",
                }}
              >
                Clear
              </button>

              <div
                style={{
                  color: "#666",
                  fontSize: "12px",
                }}
              >
                {terms.length.toLocaleString()} terms
              </div>
            </>
          )}
        </div>
      </div>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "12px",
          marginBottom: "6px",
          flexWrap: "wrap",
        }}
      >
        <button
          type="button"
          onClick={() => fileInputRef.current?.click()}
          style={{
            display: "inline-block",
            padding: "10px 16px",
            background: "#171717",
            border: "1px solid #333",
            borderRadius: "10px",
            cursor: "pointer",
            fontSize: "14px",
            color: "white",
          }}
        >
          Upload Glossary
        </button>

        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.csv"
          onChange={handleFileUpload}
          style={{ display: "none" }}
        />

        <div
          style={{
            color: isDragging ? "#60a5fa" : "#666",
            fontSize: "13px",
          }}
        >
          Upload or drag & drop an Excel/CSV glossary.
        </div>
      </div>

      <input
        ref={inputRef}
        autoFocus
        type="text"
        placeholder="Search terminology..."
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        style={{
          width: "100%",
          marginTop: "12px",
          padding: "14px",
          fontSize: "16px",
          borderRadius: "12px",
          border: "1px solid #333",
          background: "#171717",
          color: "white",
          outline: "none",
        }}
      />

      <div style={{ marginTop: "24px" }}>
        {filteredResults.map((item, index) => {
          const isSelected = index === selectedIndex;

          return (
            <div
              key={index}
              style={{
                padding: "18px",
                marginBottom: "12px",
                borderRadius: "12px",
                background: isSelected ? "#262626" : "#171717",
                border: isSelected
                  ? "1px solid #3b82f6"
                  : "1px solid #333",
              }}
            >
              <div
                style={{
                  fontSize: "20px",
                  fontWeight: 600,
                }}
              >
                {item.en}
              </div>

              <div
                style={{
                  marginTop: "6px",
                  fontSize: "20px",
                  color: "#60a5fa",
                }}
              >
                {item.kr}
              </div>

              <div
                style={{
                  marginTop: "10px",
                  color: "#888",
                  fontSize: "14px",
                }}
              >
                {item.comment}
              </div>
            </div>
          );
        })}
      </div>
          </div>

      <img
        src="/favicon.png"
        alt="logo"
        style={{
          position: "fixed",
          bottom: "20px",
          right: "20px",
          width: "40px",
          height: "40px",
          opacity: 0.22,
          pointerEvents: "none",
          userSelect: "none",
        }}
      />
    </div>
);
}

export default App;