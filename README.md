# Financial Valuation Tool — Render Deployment Guide

AI-powered DCF valuation tool for deep-tech medtech startups.  
Built at Hackathon 2026 · EPFL/ETH Medtech Comparable Database · 66 companies × 10 sub-sectors.

---

## 📁 Project Structure

```
valuation-tool/
├── public/
│   ├── index.html          # HTML shell
│   └── favicon.ico
├── src/
│   ├── index.jsx           # React entry point
│   └── App.jsx             # Main application (Valuation Tool v6)
├── package.json
├── webpack.config.js
├── .babelrc
├── .nvmrc                  # Node 18
├── render.yaml             # Render deployment config
└── README.md
```

---

## 🚀 Deploy to Render (step by step)

### Option A — Via render.yaml (recommended, 1 click)

1. Push this folder to a **GitHub repository** (public or private)
2. Go to [render.com](https://render.com) → **New** → **Web Service**
3. Connect your GitHub repo
4. Render auto-detects `render.yaml` and pre-fills all settings
5. Click **Deploy** — build takes ~2–3 minutes
6. Your app is live at `https://valuation-tool.onrender.com`

### Option B — Manual configuration on Render

| Setting | Value |
|---|---|
| **Environment** | Node |
| **Build Command** | `npm install && npm run build` |
| **Start Command** | `npx serve -s dist -l $PORT` |
| **Node Version** | 18 |
| **Plan** | Free (or Starter for always-on) |

---

## 🔑 API Key (Anthropic)

The tool calls the Anthropic API for:
- Document extraction (PDF/Excel → financial data)
- AI-generated investor narrative (PDF report)
- Valuation Advisor chatbox

**To enable AI features**, add your API key as an environment variable on Render:

```
ANTHROPIC_API_KEY = sk-ant-xxxxxxxxxxxxxxxx
```

> **Without the key**: The tool still works — WACC is computed from form inputs,  
> comparables are matched, and the Excel model is generated. Only the AI extraction  
> and PDF narrative are unavailable (graceful fallback messages shown).

---

## 🏗️ Local Development

```bash
# Install dependencies
npm install

# Start dev server (localhost:3000)
npm start

# Production build
npm run build

# Serve production build locally
npm run serve
```

---

## 📊 What the tool produces

| Output | Description |
|---|---|
| **Excel model** | Valuation_Template.xlsx pre-filled with WACC, comparables, scenarios, sensitivity |
| **PDF Report** | Investor-ready 6-section report (AI-generated narrative + WACC table + comparables) |
| **Shadow DCF** | Bear/Base/Bull indicative valuation range shown in UI |

### Excel sheets
- **Comparables** — 5 matched companies + selection methodology + per-company rationale
- **Cover** — company name, sector, date
- **Assumptions** — WACC inputs (Rf, ERP, Beta, Ke, WACC), deal structure, OPEX
- **Revenue** — deal 1/2/3 revenue projections
- **PnL_OPEX** — gross margin, EBITDA
- **FCF_DCF** — free cash flow, NPV, terminal value
- **Scenarios** — Bear/Base/Bull parameters
- **Sensitivity** — WACC × terminal growth sensitivity table
- **Dashboard** — summary KPIs

---

## 🗄️ Database

**EPFL/ETH Medtech Comparable Database** — 66 publicly listed medtech companies × 10 sub-sectors:

1. Neurovascular & Neurotech
2. Orthopedics & Musculoskeletal
3. Wearable & Drug Delivery
4. Cardiovascular & Interventional
5. Diagnostics & IVD
6. Surgical Robotics & MIS
7. Digital Health & SaMD
8. Ophthalmology & Photonics
9. Regenerative Medicine & Wound Care
10. Monitoring & Implantables

Data corrections applied (confirmed yfinance errors): VCEL D/E=0.0×, RaySearch D/E=0.01×, Creo Medical GM=46.6%.

---

## ⚠️ Disclaimer

This tool generates indicative DCF valuations for informational purposes only.  
Outputs do not constitute investment advice or certified valuations.  
Source: Damodaran January 2025 · EPFL/ETH Medtech DB · CAPM bottom-up WACC.
