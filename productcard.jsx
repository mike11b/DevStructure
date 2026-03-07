import { useState } from "react";

// ─── Mock data for demo ───────────────────────────────────────────
const DEMO_PRODUCTS = [
  {
    id: 1,
    slug: "excel-invoice-automator",
    title: "Invoice Automator Pro",
    tagline: "Genera 500 facturas en 8 segundos con un solo clic",
    language: "VBA",
    type: "Macro",
    level: "Intermedio",
    price: 349,
    originalPrice: 499,
    rating: 4.9,
    reviews: 128,
    sales: 847,
    lines: 1240,
    files: 3,
    tags: ["Excel", "VBA", "Facturas", "Automatización"],
    previewLines: [
      "Sub GenerateInvoices()",
      "  Dim ws As Worksheet",
      "  Dim lastRow As Long",
      "  Dim invoiceNum As Integer",
      "  ",
      "  Set ws = ThisWorkbook.Sheets(\"Data\")",
      "  lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row",
      "  invoiceNum = GetNextInvoiceNumber()",
    ],
    badge: "Best Seller",
    badgeColor: "#F59E0B",
    gradient: ["#00D4FF", "#0077AA"],
    icon: "⚡",
  },
  {
    id: 2,
    slug: "pandas-data-cleaner",
    title: "Pandas Data Cleaner",
    tagline: "Limpia datasets sucios en segundos con detección de anomalías",
    language: "Python",
    type: "Script",
    level: "Avanzado",
    price: 449,
    originalPrice: null,
    rating: 4.8,
    reviews: 94,
    sales: 612,
    lines: 890,
    files: 5,
    tags: ["Python", "Pandas", "Data Science", "ETL"],
    previewLines: [
      "import pandas as pd",
      "import numpy as np",
      "from typing import Optional, Dict",
      "",
      "class DataCleaner:",
      "  def __init__(self, df: pd.DataFrame,",
      "               config: Optional[Dict] = None):",
      "    self.df = df.copy()",
    ],
    badge: "Nuevo",
    badgeColor: "#10B981",
    gradient: ["#6EE7B7", "#059669"],
    icon: "🐍",
  },
  {
    id: 3,
    slug: "manual-vba-desde-cero",
    title: "Manual VBA Desde Cero",
    tagline: "De principiante a automatizador profesional en 12 horas",
    language: "VBA",
    type: "Manual",
    level: "Principiante",
    price: 199,
    originalPrice: 299,
    rating: 4.7,
    reviews: 203,
    sales: 1203,
    lines: null,
    files: 1,
    tags: ["VBA", "Excel", "Manual PDF", "148 páginas"],
    previewLines: [
      "# Capítulo 1: Introducción a VBA",
      "## ¿Por qué automatizar con VBA?",
      "",
      "VBA (Visual Basic for Applications) es el",
      "lenguaje de programación integrado en los",
      "productos de Microsoft Office que permite",
      "automatizar tareas repetitivas y crear",
      "herramientas personalizadas...",
    ],
    badge: "Top Rated",
    badgeColor: "#8B5CF6",
    gradient: ["#A78BFA", "#7C3AED"],
    icon: "📘",
  },
];

const LANG_COLORS = {
  VBA: { bg: "#1a1035", accent: "#A78BFA", dot: "#7C3AED" },
  Python: { bg: "#0d2818", accent: "#6EE7B7", dot: "#10B981" },
  JavaScript: { bg: "#1a1200", accent: "#FCD34D", dot: "#F59E0B" },
  SQL: { bg: "#0d1f35", accent: "#60A5FA", dot: "#3B82F6" },
};

// ─── ProductCard Component ─────────────────────────────────────────
function ProductCard({ product }) {
  const [hovering, setHovering] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const langTheme = LANG_COLORS[product.language] || LANG_COLORS["VBA"];
  const discount = product.originalPrice
    ? Math.round((1 - product.price / product.originalPrice) * 100)
    : null;

  return (
    <div
      onMouseEnter={() => setHovering(true)}
      onMouseLeave={() => setHovering(false)}
      style={{
        position: "relative",
        width: "340px",
        background: "#0D1117",
        border: `1px solid ${hovering ? langTheme.accent + "60" : "#1E2A38"}`,
        borderRadius: "12px",
        overflow: "hidden",
        transition: "all 0.3s cubic-bezier(0.4,0,0.2,1)",
        transform: hovering ? "translateY(-4px)" : "translateY(0)",
        boxShadow: hovering
          ? `0 20px 60px rgba(0,0,0,0.6), 0 0 0 1px ${langTheme.accent}30, inset 0 1px 0 ${langTheme.accent}20`
          : "0 4px 20px rgba(0,0,0,0.4)",
        cursor: "pointer",
        fontFamily: "'JetBrains Mono', 'Fira Code', 'Courier New', monospace",
      }}
    >
      {/* Gradient top bar */}
      <div style={{
        height: "3px",
        background: `linear-gradient(90deg, ${product.gradient[0]}, ${product.gradient[1]})`,
        opacity: hovering ? 1 : 0.6,
        transition: "opacity 0.3s",
      }} />

      {/* Header */}
      <div style={{ padding: "16px 18px 12px", background: "#0A0E1A" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          {/* Language badge + icon */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <span style={{ fontSize: "20px" }}>{product.icon}</span>
            <div style={{
              display: "flex", alignItems: "center", gap: "4px",
              background: langTheme.bg, border: `1px solid ${langTheme.accent}40`,
              borderRadius: "4px", padding: "2px 8px",
            }}>
              <span style={{
                width: "6px", height: "6px", borderRadius: "50%",
                background: langTheme.dot, display: "inline-block",
                boxShadow: `0 0 6px ${langTheme.dot}`,
              }} />
              <span style={{ color: langTheme.accent, fontSize: "11px", fontWeight: "600", letterSpacing: "0.05em" }}>
                {product.language}
              </span>
            </div>
            <span style={{
              background: "#1E2A38", color: "#94A3B8", fontSize: "10px",
              padding: "2px 7px", borderRadius: "4px", border: "1px solid #2D3748",
            }}>
              {product.type}
            </span>
          </div>

          {/* Badge */}
          {product.badge && (
            <span style={{
              background: product.badgeColor + "20",
              color: product.badgeColor,
              border: `1px solid ${product.badgeColor}40`,
              fontSize: "10px", fontWeight: "700",
              padding: "3px 8px", borderRadius: "4px",
              letterSpacing: "0.04em", textTransform: "uppercase",
            }}>
              {product.badge}
            </span>
          )}
        </div>

        {/* Title */}
        <div style={{ marginTop: "12px" }}>
          <h3 style={{
            margin: 0, color: "#F1F5F9", fontSize: "15px",
            fontWeight: "700", lineHeight: "1.3", letterSpacing: "-0.01em",
            fontFamily: "'Inter', 'Segoe UI', sans-serif",
          }}>
            {product.title}
          </h3>
          <p style={{
            margin: "4px 0 0", color: "#64748B", fontSize: "12px",
            lineHeight: "1.5", fontFamily: "'Inter', sans-serif",
          }}>
            {product.tagline}
          </p>
        </div>
      </div>

      {/* Code Preview */}
      <div style={{
        margin: "0 16px",
        background: "#0A0E1A",
        border: "1px solid #1E2A38",
        borderRadius: "8px",
        overflow: "hidden",
      }}>
        {/* Terminal header */}
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "7px 12px", background: "#111827", borderBottom: "1px solid #1E2A38",
        }}>
          <div style={{ display: "flex", gap: "5px" }}>
            {["#FF5F57", "#FEBC2E", "#28C840"].map((c, i) => (
              <span key={i} style={{ width: "9px", height: "9px", borderRadius: "50%", background: c, display: "block" }} />
            ))}
          </div>
          <span style={{ color: "#475569", fontSize: "10px" }}>
            {product.type === "Manual" ? "README.md" : `main.${product.language === "Python" ? "py" : product.language === "VBA" ? "bas" : "js"}`}
          </span>
          <button
            onClick={(e) => { e.stopPropagation(); setPreviewOpen(!previewOpen); }}
            style={{
              background: "transparent", border: `1px solid ${langTheme.accent}40`,
              color: langTheme.accent, fontSize: "9px", padding: "2px 7px",
              borderRadius: "3px", cursor: "pointer", letterSpacing: "0.05em",
            }}
          >
            {previewOpen ? "▲ CERRAR" : "▼ PREVIEW"}
          </button>
        </div>

        {/* Code lines */}
        <div style={{
          padding: "10px 0",
          maxHeight: previewOpen ? "200px" : "100px",
          overflow: "hidden",
          transition: "max-height 0.35s ease",
          position: "relative",
        }}>
          {product.previewLines.map((line, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center" }}>
              <span style={{
                color: "#2D3748", fontSize: "11px", width: "32px",
                textAlign: "right", paddingRight: "12px", userSelect: "none",
              }}>
                {i + 1}
              </span>
              <span style={{
                color: i === 0 ? langTheme.accent : i % 3 === 0 ? "#94A3B8" : "#CBD5E1",
                fontSize: "11px", lineHeight: "1.7",
                fontFamily: "'JetBrains Mono', monospace",
              }}>
                {line || "\u00A0"}
              </span>
            </div>
          ))}
          {/* Blur overlay cuando está cerrado */}
          {!previewOpen && (
            <div style={{
              position: "absolute", bottom: 0, left: 0, right: 0,
              height: "50px",
              background: "linear-gradient(transparent, #0A0E1A)",
            }} />
          )}
        </div>

        {/* Stats row */}
        {product.lines && (
          <div style={{
            display: "flex", gap: "12px", padding: "8px 12px",
            borderTop: "1px solid #1E2A38", background: "#080C14",
          }}>
            {[
              { label: "líneas", val: product.lines.toLocaleString() },
              { label: "archivos", val: product.files },
              { label: "nivel", val: product.level },
            ].map((s) => (
              <div key={s.label} style={{ display: "flex", gap: "4px", alignItems: "baseline" }}>
                <span style={{ color: langTheme.accent, fontSize: "11px", fontWeight: "700" }}>{s.val}</span>
                <span style={{ color: "#475569", fontSize: "10px" }}>{s.label}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Tags */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", padding: "12px 16px 8px" }}>
        {product.tags.map((tag) => (
          <span key={tag} style={{
            background: "#1E2A38", color: "#64748B",
            fontSize: "10px", padding: "2px 7px", borderRadius: "3px",
            border: "1px solid #2D3748",
          }}>
            {tag}
          </span>
        ))}
      </div>

      {/* Footer: rating + price + CTA */}
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "12px 16px 16px",
      }}>
        {/* Rating */}
        <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
            <span style={{ color: "#F59E0B", fontSize: "13px" }}>{"★".repeat(5)}</span>
            <span style={{ color: "#F1F5F9", fontSize: "12px", fontWeight: "700" }}>{product.rating}</span>
          </div>
          <span style={{ color: "#475569", fontSize: "10px" }}>
            {product.reviews} reviews · {product.sales.toLocaleString()} ventas
          </span>
        </div>

        {/* Price + CTA */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "6px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
            {discount && (
              <span style={{
                background: "#7C3AED20", color: "#A78BFA",
                fontSize: "10px", padding: "1px 5px", borderRadius: "3px",
                border: "1px solid #7C3AED40",
              }}>
                -{discount}%
              </span>
            )}
            {product.originalPrice && (
              <span style={{ color: "#374151", fontSize: "11px", textDecoration: "line-through" }}>
                ${product.originalPrice}
              </span>
            )}
            <span style={{
              color: "#F1F5F9", fontSize: "16px", fontWeight: "800",
              fontFamily: "'Inter', sans-serif",
            }}>
              ${product.price} <span style={{ fontSize: "10px", color: "#64748B" }}>MXN</span>
            </span>
          </div>
          <button style={{
            background: hovering
              ? `linear-gradient(135deg, ${product.gradient[0]}, ${product.gradient[1]})`
              : "#1E2A38",
            color: hovering ? "#0A0E1A" : langTheme.accent,
            border: `1px solid ${langTheme.accent}50`,
            borderRadius: "6px",
            padding: "7px 16px",
            fontSize: "11px", fontWeight: "700",
            cursor: "pointer",
            letterSpacing: "0.05em",
            transition: "all 0.25s ease",
            fontFamily: "'JetBrains Mono', monospace",
          }}>
            COMPRAR →
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Demo: ProductCard Grid ────────────────────────────────────────
export default function DVStructureShowcase() {
  const [filter, setFilter] = useState("Todos");
  const langs = ["Todos", "VBA", "Python", "JavaScript"];

  const filtered = filter === "Todos"
    ? DEMO_PRODUCTS
    : DEMO_PRODUCTS.filter((p) => p.language === filter);

  return (
    <div style={{
      minHeight: "100vh",
      background: "#060A10",
      padding: "40px 24px",
      fontFamily: "'Inter', 'Segoe UI', sans-serif",
    }}>
      {/* Header */}
      <div style={{ textAlign: "center", marginBottom: "40px" }}>
        <div style={{
          display: "inline-flex", alignItems: "center", gap: "8px",
          background: "#0A0E1A", border: "1px solid #00D4FF30",
          borderRadius: "6px", padding: "6px 14px", marginBottom: "20px",
        }}>
          <span style={{ width: "6px", height: "6px", borderRadius: "50%", background: "#00D4FF", boxShadow: "0 0 8px #00D4FF", display: "block" }} />
          <span style={{ color: "#00D4FF", fontSize: "12px", fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.1em" }}>
            MARKETPLACE v1.0
          </span>
        </div>
        <h1 style={{
          margin: 0, color: "#F1F5F9", fontSize: "36px", fontWeight: "800",
          letterSpacing: "-0.03em", lineHeight: "1.1",
        }}>
          DVStructure
          <span style={{ display: "block", background: "linear-gradient(90deg, #00D4FF, #6EE7B7)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
            ProductCard Component
          </span>
        </h1>
        <p style={{ color: "#64748B", fontSize: "14px", marginTop: "12px" }}>
          Hover sobre las cards para ver las interacciones • Click en "PREVIEW" para expandir el código
        </p>

        {/* Filter tabs */}
        <div style={{ display: "flex", gap: "8px", justifyContent: "center", marginTop: "24px" }}>
          {langs.map((l) => (
            <button
              key={l}
              onClick={() => setFilter(l)}
              style={{
                background: filter === l ? "#00D4FF15" : "transparent",
                color: filter === l ? "#00D4FF" : "#475569",
                border: `1px solid ${filter === l ? "#00D4FF40" : "#1E2A38"}`,
                borderRadius: "6px", padding: "6px 14px",
                fontSize: "12px", cursor: "pointer",
                fontFamily: "'JetBrains Mono', monospace",
                transition: "all 0.2s",
              }}
            >
              {l}
            </button>
          ))}
        </div>
      </div>

      {/* Cards grid */}
      <div style={{
        display: "flex", flexWrap: "wrap", gap: "24px", justifyContent: "center",
      }}>
        {filtered.map((product) => (
          <ProductCard key={product.id} product={product} />
        ))}
      </div>

      {/* Footer note */}
      <div style={{ textAlign: "center", marginTop: "48px", color: "#1E2A38", fontSize: "11px", fontFamily: "'JetBrains Mono', monospace" }}>
        // components/shop/ProductCard.jsx — DVStructure v1.0
      </div>
    </div>
  );
}
