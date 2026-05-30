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

  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const saved = localStorage.getItem("terms");

    if (saved) {
      const parsed = JSON.parse(saved);

      setTerms(parsed);
      setFilteredResults(parsed);
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
      .filter((item) => {
        return (
          item.en.toLowerCase().includes(trimmedQuery) ||
          item.kr.toLowerCase().includes(trimmedQuery) ||
          item.comment.toLowerCase().includes(trimmedQuery)
        );
      })
      .sort((a, b) => {
        const aEnStarts = a.en.toLowerCase().startsWith(trimmedQuery);
        const bEnStarts = b.en.toLowerCase().startsWith(trimmedQuery);

        if (aEnStarts && !bEnStarts) return -1;
        if (!aEnStarts && bEnStarts) return 1;

        const aKrStarts = a.kr.toLowerCase().startsWith(trimmedQuery);
        const bKrStarts = b.kr.toLowerCase().startsWith(trimmedQuery);

        if (aKrStarts && !bKrStarts) return -1;
        if (!aKrStarts && bKrStarts) return 1;

        return a.en.localeCompare(b.en);
      });

    setFilteredResults(ranked.slice(0, 20));
    setSelectedIndex(0);
  }, [query, terms]);

  const handleFileUpload = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];

    if (!file) return;

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
          const textToCopy = `${selected.en} → ${selected.kr}`;

          await navigator.clipboard.writeText(textToCopy);
        }
      }

      if (e.key === "Escape") {
        setQuery("");
        setSelectedIndex(0);

        inputRef.current?.focus();
      }
    };

    window.addEventListener("keydown", handleKeyDown);

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [filteredResults, selectedIndex]);

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#0a0a0a",
        color: "white",
        padding: "40px",
        fontFamily: "sans-serif",
      }}
    >
      <div style={{ maxWidth: "900px", margin: "0 auto" }}>
        <h1
          style={{
            fontSize: "22px",
            fontWeight: "bold",
            marginBottom: "16px",
          }}
        >
          Terminology Search
        </h1>

        <input
          type="file"
          accept=".xlsx,.csv"
          onChange={handleFileUpload}
        />

        <input
          ref={inputRef}
          autoFocus
          type="text"
          placeholder="Search terminology..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          style={{
            width: "100%",
            marginTop: "20px",
            padding: "14px",
            fontSize: "16px",
            borderRadius: "12px",
            border: "1px solid #333",
            background: "#171717",
            color: "white",
            outline: "none",
          }}
        />

        <div style={{ marginTop: "16px", color: "#aaa" }}>
          Loaded terms: {terms.length}
        </div>

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
    </div>
  );
}

export default App;