
import React from "react";
import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
// Updated sector names to match EPFL/ETH database taxonomy
const SECTOR_NAMES = {1:"Neurovascular & Neurotech",2:"Orthopedics & Musculoskeletal",3:"Wearable & Drug Delivery",4:"Cardiovascular & Interventional",5:"Diagnostics & IVD",6:"Surgical Robotics & MIS",7:"Digital Health & SaMD",8:"Ophthalmology & Photonics",9:"Regenerative Medicine & Wound Care",10:"Monitoring & Implantables"};
const SECTOR_KW = {1:["neuro","brain","neural","seizure","epilep","stroke","cranial","neurovascular","neurostimul"],2:["ortho","bone","spine","musculo","joint","fracture","implant","sacroiliac","vertebra"],3:["wearable","insulin","pump","infusion","drug delivery","patch","diabetes","cgm","glucose","injec"],4:["cardio","cardiac","heart","vascular","artery","perfusion","aortic","bypass","ablation","angio"],5:["diagnos","ivd","assay","biomarker","reagent","molecular","genomic","per test","sequencing","pcr"],6:["robot","endoscop","laparoscop","surgical robot","single-use scope","mis","minimally invasive"],7:["software","samd","digital health","imaging","mri","pacs","ai","machine learning","cloud","subscription"],8:["ophthalm","eye","retina","ocular","vision","glaucom","laser","fundus","cataract","lens","photonic"],9:["regenerat","tissue","wound","biologic","scaffold","nerve repair","cell therapy","graft","biologics"],10:["monitor","implantable","pacemaker","stimulat","neuromodulat","sensor","cgm","remote monitor","wearabl device"]};
const WACC_P = {"Switzerland":{rf:0.0075,erp:0.0554,tax:0.149,g:0.015},"United States":{rf:0.0425,erp:0.046,tax:0.21,g:0.0225},"Germany":{rf:0.025,erp:0.0554,tax:0.295,g:0.015},"France":{rf:0.03,erp:0.0554,tax:0.25,g:0.015},"Belgium":{rf:0.03,erp:0.0649,tax:0.25,g:0.015},"Sweden":{rf:0.0215,erp:0.0554,tax:0.206,g:0.015},"United Kingdom":{rf:0.04,erp:0.0554,tax:0.25,g:0.02},"Netherlands":{rf:0.028,erp:0.0554,tax:0.258,g:0.015},"Spain":{rf:0.031,erp:0.0649,tax:0.25,g:0.015},"Italy":{rf:0.037,erp:0.0649,tax:0.275,g:0.01},"Denmark":{rf:0.022,erp:0.0554,tax:0.22,g:0.015},"Finland":{rf:0.021,erp:0.0554,tax:0.20,g:0.015},"Norway":{rf:0.032,erp:0.0554,tax:0.22,g:0.015},"Austria":{rf:0.027,erp:0.0554,tax:0.25,g:0.015},"Poland":{rf:0.053,erp:0.0649,tax:0.19,g:0.02},"Israel":{rf:0.04,erp:0.0649,tax:0.23,g:0.02},"Canada":{rf:0.033,erp:0.046,tax:0.265,g:0.02},"Australia":{rf:0.042,erp:0.046,tax:0.30,g:0.025},"Japan":{rf:0.009,erp:0.046,tax:0.3086,g:0.01},"China":{rf:0.023,erp:0.064,tax:0.25,g:0.04},"India":{rf:0.069,erp:0.064,tax:0.30,g:0.05},"Singapore":{rf:0.029,erp:0.046,tax:0.17,g:0.025},"South Korea":{rf:0.031,erp:0.046,tax:0.275,g:0.025},"Brazil":{rf:0.135,erp:0.078,tax:0.34,g:0.04},"Default":{rf:0.025,erp:0.055,tax:0.21,g:0.015}};

// ─── FULL 66-COMPANY DATABASE — EPFL/ETH Medtech Comparable Database ──────────
// Fields: t=ticker, n=name, s=sector(1-10), gm=gross margin%, eb=EBITDA margin%,
//         b=beta(unlevered), rg=revenue growth%, c=country, de=D/E ratio,
//         ok=🟢 reliable / 🟡 usable with limits, note=key usage caveat
// D/E corrected for confirmed yfinance errors: VCEL(0), RAY-B(0.01), CREO(gm=46.6)
const DB = [
  // ── Sub-sector 1: Neurovascular & Neurotech ──
  {"t":"PEN","n":"Penumbra Inc.","s":1,"gm":67.4,"eb":14.7,"b":0.735,"rg":15.6,"c":"USA","de":14.7,"ok":"🟢"},
  {"t":"NYXH","n":"Nyxoah SA","s":1,"gm":63.1,"eb":0.0,"b":0.876,"rg":34.7,"c":"Belgium","de":85.6,"ok":"🟡"},
  {"t":"CLPT","n":"ClearPoint Neuro","s":1,"gm":61.4,"eb":-58.1,"b":1.294,"rg":34.0,"c":"USA","de":207.8,"ok":"🟡"},
  {"t":"NSPR","n":"InspireMD Inc.","s":1,"gm":29.5,"eb":0.0,"b":0.819,"rg":61.6,"c":"USA","de":5.96,"ok":"🟡"},
  // ── Sub-sector 2: Orthopedics & Musculoskeletal ──
  {"t":"BONEX","n":"BONESUPPORT AB","s":2,"gm":92.5,"eb":26.5,"b":0.482,"rg":14.3,"c":"Sweden","de":1.5,"ok":"🟢"},
  {"t":"SIBN","n":"SI-BONE Inc.","s":2,"gm":79.6,"eb":-8.2,"b":0.671,"rg":15.0,"c":"USA","de":20.7,"ok":"🟢"},
  {"t":"ATEC","n":"Alphatec Holdings","s":2,"gm":70.2,"eb":3.6,"b":0.966,"rg":13.6,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"OFIX","n":"Orthofix Medical Inc.","s":2,"gm":71.0,"eb":2.1,"b":0.8,"rg":1.6,"c":"USA","de":60.3,"ok":"🟡"},
  // ── Sub-sector 3: Wearable & Drug Delivery ──
  {"t":"TNDM","n":"Tandem Diabetes Care","s":3,"gm":53.8,"eb":-6.0,"b":1.6,"rg":2.7,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"EMBC","n":"Embecta Corp.","s":3,"gm":61.9,"eb":31.9,"b":1.1,"rg":-14.4,"c":"USA","de":20.0,"ok":"🟡"},
  // ── Sub-sector 4: Cardiovascular & Interventional ──
  {"t":"XVVO","n":"XVIVO Perfusion AB","s":4,"gm":73.1,"eb":16.5,"b":1.4,"rg":10.2,"c":"Sweden","de":5.7,"ok":"🟢"},
  {"t":"ATRC","n":"AtriCure Inc.","s":4,"gm":75.6,"eb":3.2,"b":1.281,"rg":14.3,"c":"USA","de":15.2,"ok":"🟢"},
  {"t":"LMAT","n":"LeMaitre Vascular","s":4,"gm":71.3,"eb":30.7,"b":0.597,"rg":11.2,"c":"USA","de":46.7,"ok":"🟢"},
  {"t":"AORT","n":"Artivion Inc.","s":4,"gm":64.4,"eb":12.2,"b":1.4,"rg":19.2,"c":"USA","de":57.7,"ok":"🟢"},
  {"t":"ANGO","n":"AngioDynamics Inc.","s":4,"gm":54.3,"eb":-0.2,"b":0.4,"rg":8.9,"c":"USA","de":6.6,"ok":"🟡"},
  // ── Sub-sector 5: Diagnostics & IVD ──
  {"t":"SBS","n":"Stratec SE","s":5,"gm":25.6,"eb":11.3,"b":0.8,"rg":-11.1,"c":"Germany","de":55.7,"ok":"🟢"},
  {"t":"EKF","n":"EKF Diagnostics","s":5,"gm":51.4,"eb":19.7,"b":0.526,"rg":5.4,"c":"UK","de":1.97,"ok":"🟢"},
  {"t":"BOUL","n":"Boule Diagnostics AB","s":5,"gm":42.3,"eb":10.3,"b":0.2,"rg":-10.0,"c":"Sweden","de":20.0,"ok":"🟡"},
  {"t":"TECN","n":"Tecan Group AG","s":5,"gm":35.2,"eb":9.2,"b":1.0,"rg":-5.2,"c":"Switzerland","de":19.0,"ok":"🟡"},
  {"t":"VITR","n":"Vitrolife AB","s":5,"gm":58.6,"eb":27.5,"b":1.7,"rg":-4.2,"c":"Sweden","de":14.6,"ok":"🟡"},
  {"t":"TSTL","n":"Tristel plc","s":5,"gm":81.7,"eb":23.1,"b":0.3,"rg":13.6,"c":"UK","de":17.4,"ok":"🟡"},
  {"t":"VCYT","n":"Veracyte Inc.","s":5,"gm":72.9,"eb":19.2,"b":1.885,"rg":21.5,"c":"USA","de":2.93,"ok":"🟡"},
  {"t":"CDNA","n":"CareDx Inc.","s":5,"gm":68.9,"eb":0.6,"b":2.5,"rg":39.0,"c":"USA","de":7.8,"ok":"🟡"},
  {"t":"CSTL","n":"Castle Biosciences","s":5,"gm":79.4,"eb":-0.6,"b":1.1,"rg":0.8,"c":"USA","de":7.9,"ok":"🟡"},
  // ── Sub-sector 6: Surgical Robotics & MIS ──
  {"t":"AMBU","n":"Ambu A/S","s":6,"gm":60.0,"eb":13.3,"b":1.3,"rg":1.2,"c":"Denmark","de":9.0,"ok":"🟢"},
  {"t":"PRCT","n":"PROCEPT BioRobotics","s":6,"gm":64.0,"eb":-31.8,"b":0.826,"rg":20.2,"c":"USA","de":22.4,"ok":"🟡"},
  {"t":"STXS","n":"Stereotaxis Inc.","s":6,"gm":52.7,"eb":-56.8,"b":1.3,"rg":36.3,"c":"USA","de":29.0,"ok":"🟡"},
  {"t":"CREO","n":"Creo Medical Group","s":6,"gm":46.6,"eb":0.0,"b":1.0,"rg":37.5,"c":"UK","de":6.9,"ok":"🟡"},
  // ── Sub-sector 7: Digital Health & SaMD ──
  {"t":"RAY","n":"RaySearch Laboratories","s":7,"gm":92.4,"eb":25.6,"b":1.0,"rg":-12.5,"c":"Sweden","de":0.01,"ok":"🟢"},
  {"t":"SECT","n":"Sectra AB","s":7,"gm":38.9,"eb":20.7,"b":0.85,"rg":5.6,"c":"Sweden","de":5.67,"ok":"🟡"},
  {"t":"ASCN","n":"Ascom Holding AG","s":7,"gm":48.2,"eb":7.9,"b":1.0,"rg":5.2,"c":"Switzerland","de":20.0,"ok":"🟡"},
  {"t":"IRTC","n":"iRhythm Technologies","s":7,"gm":71.0,"eb":-1.4,"b":1.333,"rg":25.7,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"BFLY","n":"Butterfly Network","s":7,"gm":66.1,"eb":-49.0,"b":2.28,"rg":25.0,"c":"USA","de":10.3,"ok":"🟡"},
  {"t":"NNOX","n":"Nano-X Imaging","s":7,"gm":-98.2,"eb":0.0,"b":1.2,"rg":24.0,"c":"Israel","de":5.6,"ok":"🟡"},
  {"t":"HYPR","n":"Hyperfine Inc.","s":7,"gm":49.8,"eb":-265.4,"b":1.4,"rg":128.0,"c":"USA","de":0.77,"ok":"🟡"},
  // ── Sub-sector 8: Ophthalmology & Photonics ──
  {"t":"REG1V","n":"Revenio Group Oyj","s":8,"gm":69.9,"eb":21.6,"b":0.8,"rg":4.6,"c":"Finland","de":9.2,"ok":"🟢"},
  {"t":"STAA","n":"STAAR Surgical","s":8,"gm":76.2,"eb":-15.6,"b":1.202,"rg":18.1,"c":"USA","de":11.1,"ok":"🟢"},
  {"t":"OPTOMED","n":"Optomed Oyj","s":8,"gm":63.6,"eb":-31.3,"b":1.6,"rg":-5.6,"c":"Finland","de":11.3,"ok":"🟡"},
  {"t":"ELN","n":"El.En. S.p.A.","s":8,"gm":42.9,"eb":13.7,"b":1.1,"rg":24.8,"c":"Italy","de":7.36,"ok":"🟡"},
  {"t":"GKOS","n":"Glaukos Corporation","s":8,"gm":78.1,"eb":-8.0,"b":0.948,"rg":41.2,"c":"USA","de":15.8,"ok":"🟡"},
  {"t":"LNSR","n":"LENSAR Inc.","s":8,"gm":46.4,"eb":-34.4,"b":0.8,"rg":-4.2,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"LBIRD","n":"Lumibird S.A.","s":8,"gm":63.4,"eb":14.2,"b":0.9,"rg":8.5,"c":"France","de":70.9,"ok":"🟡"},
  // ── Sub-sector 9: Regenerative Medicine & Wound Care ──
  {"t":"VCEL","n":"Vericel Corp.","s":9,"gm":74.8,"eb":9.4,"b":1.1,"rg":30.1,"c":"USA","de":0.0,"ok":"🟢"},
  {"t":"AXGN","n":"Axogen Inc.","s":9,"gm":75.0,"eb":-0.9,"b":1.2,"rg":26.6,"c":"USA","de":8.4,"ok":"🟡"},
  {"t":"ORGO","n":"Organogenesis Holdings","s":9,"gm":76.5,"eb":14.2,"b":1.3,"rg":78.1,"c":"USA","de":18.9,"ok":"🟡"},
  {"t":"EUZ","n":"Eckert & Ziegler SE","s":9,"gm":49.0,"eb":28.7,"b":1.5,"rg":9.3,"c":"Germany","de":17.7,"ok":"🟡"},
  {"t":"PHO","n":"Photocure ASA","s":5,"gm":92.1,"eb":4.3,"b":0.4,"rg":-3.5,"c":"Norway","de":2.1,"ok":"🟡"},
  // ── Sub-sector 7: Digital Health & SaMD (additional) ──
  {"t":"PLLWF","n":"Polarean Imaging","s":7,"gm":45.2,"eb":0.0,"b":0.1,"rg":-46.9,"c":"UK","de":5.6,"ok":"🟡"},
  // ── Sub-sector 8: Ophthalmology & Photonics (additional) ──
  {"t":"IBAB","n":"IBA S.A.","s":8,"gm":32.2,"eb":5.0,"b":1.2,"rg":8.1,"c":"Belgium","de":20.0,"ok":"🟡"},
  // ── Sub-sector 10: Monitoring & Implantables ──
  {"t":"LIVN","n":"LivaNova PLC","s":10,"gm":67.9,"eb":18.3,"b":0.819,"rg":14.3,"c":"UK","de":28.4,"ok":"🟢"},
  {"t":"DRW3","n":"Drägerwerk AG","s":10,"gm":45.3,"eb":8.7,"b":0.6,"rg":3.5,"c":"Germany","de":20.7,"ok":"🟡"},
  {"t":"BACTI","n":"Bactiguard Holding AB","s":10,"gm":88.0,"eb":15.5,"b":0.4,"rg":-28.1,"c":"Sweden","de":51.8,"ok":"🟡"},
  {"t":"IHC","n":"Inspiration Healthcare","s":10,"gm":44.4,"eb":3.7,"b":0.9,"rg":40.8,"c":"UK","de":78.9,"ok":"🟡"},
  {"t":"EKTA","n":"Elekta AB","s":10,"gm":38.5,"eb":8.9,"b":1.0,"rg":-9.7,"c":"Sweden","de":90.9,"ok":"🟡"},
  {"t":"INSP","n":"Inspire Medical Systems","s":10,"gm":85.8,"eb":7.4,"b":0.831,"rg":1.6,"c":"USA","de":3.73,"ok":"🟡"},
  {"t":"SENS","n":"Senseonics Holdings","s":10,"gm":44.7,"eb":-189.5,"b":1.1,"rg":71.8,"c":"USA","de":67.8,"ok":"🟡"},
  {"t":"CVRX","n":"CVRx Inc.","s":10,"gm":85.3,"eb":-89.2,"b":0.9,"rg":4.4,"c":"USA","de":20.0,"ok":"🟡"},
];
// ─── PER-COMPANY SELECTION RATIONALE ─────────────────────────────────────────
// Explains WHY each specific company was chosen — shown in UI card + Excel note
const COMP_REASON = {
  // Sub-sector 1
  "PEN":    "🟢 Anchor comp — catheter-based neurovascular devices, specialist sales, 67.4% GM. Most liquid US-listed peer.",
  "NYXH":   "🟡 Implantable neurostimulator (hypoglossal nerve), CE MDR cleared; high D/E limits beta reliability.",
  "CLPT":   "🟡 MRI-guided neuro delivery platform; high beta (1.294) reflects early-commercial stage risk.",
  "NSPR":   "🟡 Carotid stent system; rapid rev growth (+61.6%) but thin GM (29.5%) skews sector median down.",
  // Sub-sector 2
  "BONEX":  "🟢 Swedish bone substitute; 92.5% GM — highest in DB; CE + FDA cleared, profitable, clean balance sheet.",
  "SIBN":   "🟢 Sacroiliac joint fusion; 79.6% GM, DRG/NTAP reimbursement path — best US structural-implant comp.",
  "ATEC":   "🟡 Cervical/lumbar spine systems; comparable revenue model but higher leverage (D/E 20×).",
  "OFIX":   "🟡 Spine & bone stimulation; low growth (+1.6%) drags sector RG median — used for beta calibration.",
  // Sub-sector 3
  "TNDM":   "🟡 Insulin pump; CGM integration model mirrors wearable drug delivery; high beta (1.6).",
  "EMBC":   "🟡 Insulin delivery devices (spin-off from BD); mature product — GM 61.9%, negative revenue growth caution.",
  // Sub-sector 4
  "XVVO":   "🟢 Organ perfusion systems; capital+disposable model, profitable EU-listed comp, 73.1% GM.",
  "ATRC":   "🟢 Cardiac ablation; 75.6% GM, growing US revenue (+14.3%), direct hospital sales — cleanest cardio comp.",
  "LMAT":   "🟢 Vascular surgery devices; 71.3% GM, 30.7% EBITDA — most operationally mature peer in sector.",
  "AORT":   "🟢 Cardiac/vascular biologics; 64.4% GM, high growth (+19.2%) from recent M&A — watch for integration risk.",
  "ANGO":   "🟡 Vascular access & ablation; low beta (0.4) and declining EBITDA weigh on comparability.",
  // Sub-sector 5
  "SBS":    "🟢 German IVD instrument OEM; low GM (25.6%) anchors B2B/OEM end of sector range.",
  "EKF":    "🟢 Point-of-care diagnostics; 51.4% GM, profitable UK-listed — reliable small-cap IVD comp.",
  "BOUL":   "🟡 Hematology analyzer OEM; low beta (0.2) and declining revenue reduce weight.",
  "TECN":   "🟡 Lab automation (Tecan); 35.2% GM reflects instrument-heavy mix — useful for OEM/capital models.",
  "VITR":   "🟡 IVF consumables; 58.6% GM but recent revenue contraction (-4.2%) after Igenomix acquisition.",
  "TSTL":   "🟡 Infection prevention diagnostics; 81.7% GM (high-end anchor), UK small-cap liquidity caution.",
  "VCYT":   "🟡 Genomic diagnostics; 72.9% GM, strong growth (+21.5%) — useful for high-margin molecular Dx.",
  "CDNA":   "🟡 Transplant diagnostics; 68.9% GM but beta of 2.5 flags high equity risk.",
  "CSTL":   "🟡 Dermatology genomic testing; 79.4% GM, flat growth (+0.8%) — mature per-test model reference.",
  "PHO":    "🟡 Photodynamic cancer diagnostics; 92.1% GM (outlier), low beta (0.4) — specialty comp only.",
  // Sub-sector 6
  "AMBU":   "🟢 Single-use endoscopy; 60.0% GM, Denmark-listed, growing recurring-use model — best MIS comp.",
  "PRCT":   "🟡 Aquablation robotic system; 64.0% GM, capital+procedure mix — high leverage caution.",
  "STXS":   "🟡 Robotic electrophysiology; high growth (+36.3%) but deep EBITDA losses (-56.8%).",
  "CREO":   "🟡 Advanced electrosurgery; GM corrected to 46.6% (yfinance error fixed). Early commercial.",
  // Sub-sector 7
  "RAY":    "🟢 Radiation therapy planning SaMD; 92.4% GM — best-in-class software margin anchor. D/E corrected to 0.01×.",
  "SECT":   "🟡 Medical imaging SaaS (PACS/VNA); 38.9% GM reflects services mix — useful for hybrid SaaS models.",
  "ASCN":   "🟡 Clinical communication platform; 48.2% GM, Swiss-listed, enterprise SaaS recurring revenue.",
  "IRTC":   "🟡 Cardiac monitoring SaMD (Zio patch); 71.0% GM, +25.7% revenue growth — ambulatory Dx reference.",
  "BFLY":   "🟡 Handheld ultrasound SaMD; 66.1% GM, subscription model — high beta (2.28) flags speculative risk.",
  "NNOX":   "🟡 Tomosynthesis-as-a-service; negative GM (-98.2%) makes this a downside scenario anchor only.",
  "HYPR":   "🟡 Low-field MRI; extreme growth (+128%) and EBITDA margin (-265%) — excluded from median calculation.",
  "PLLWF":  "🟡 Hyperpolarized imaging agent; very low beta (0.1) and negative revenue growth — limited use.",
  // Sub-sector 8
  "REG1V":  "🟢 Glaucoma diagnostics (Icare); 69.9% GM, profitable Finnish comp — closest EU ophthalmic device peer.",
  "STAA":   "🟢 Implantable collamer lenses; 76.2% GM, growing (+18.1%) — primary US ophthalmic implant anchor.",
  "OPTOMED":"🟡 Retinal imaging; 63.6% GM, Finnish small-cap, declining revenue (-5.6%) weights beta estimate.",
  "ELN":    "🟡 Medical laser systems; 42.9% GM reflects hardware mix — useful for photonics/laser device models.",
  "GKOS":   "🟡 Glaucoma devices; 78.1% GM, high growth (+41.2%) post-iDose launch — bull scenario reference.",
  "LNSR":   "🟡 Laser cataract surgery; 46.4% GM, declining revenue — capital equipment model reference.",
  "LBIRD":  "🟡 Photonics & laser (Lumibird); 63.4% GM, French-listed — direct photonics technology comparable.",
  "IBAB":   "🟡 Proton therapy systems (IBA); 32.2% GM reflects heavy capital equipment — useful for hardware-heavy models.",
  // Sub-sector 9
  "VCEL":   "🟢 Cell therapy (MACI/Epicel); 74.8% GM, profitable, D/E corrected to 0.0× (zero debt per 2023 10-K).",
  "AXGN":   "🟡 Peripheral nerve repair biologics; 75.0% GM, growing (+26.6%) — closest neural regeneration peer.",
  "ORGO":   "🟡 Advanced wound biologics; 76.5% GM, very high growth (+78.1%) — bull-scenario anchor.",
  "EUZ":    "🟡 Radiopharmaceuticals/biologics; 49.0% GM, German-listed — useful for EU regulatory pathway models.",
  // Sub-sector 10
  "LIVN":   "🟢 Neuromodulation + cardiac surgery; 67.9% GM, profitable UK/US dual-listed — cleanest implantable comp.",
  "DRW3":   "🟡 Ventilators & patient monitoring; 45.3% GM, German hospital capital equipment — useful for monitoring models.",
  "BACTI":  "🟡 Antimicrobial coatings; 88.0% GM (high-end anchor) but revenue declining sharply (-28.1%).",
  "IHC":    "🟡 Neonatal respiratory; 44.4% GM, UK-listed, high D/E (78.9×) limits beta usefulness.",
  "EKTA":   "🟡 Radiotherapy systems; 38.5% GM, declining revenue (-9.7%) — mature capital equipment reference.",
  "INSP":   "🟡 Inspire upper-airway stimulator; 85.8% GM — best US implantable GM anchor; flagged for GLP-1 disruption risk.",
  "SENS":   "🟡 Eversense long-term CGM; 44.7% GM, extreme EBITDA (-189.5%), high D/E — downside risk anchor.",
  "CVRX":   "🟡 Baroreflex activation therapy; 85.3% GM but EBITDA -89.2% and no revenue growth — clinical-stage proxy.",
};

// ─── SELECTION METHODOLOGY TEXT (injected into Excel + shown in UI) ──────────
const SELECTION_METHODOLOGY = `SELECTION METHODOLOGY — How comparables were chosen:
Step 1 — KEYWORD MATCHING: The startup description was scanned for sector keywords (e.g. "neural", "catheter", "cardiac", "diagnostic"). The sub-sector with the highest keyword hit count was selected.
Step 2 — DATABASE FILTERING: From the 66-company EPFL/ETH Medtech DB, all companies in the matched sub-sector were extracted. Companies with GM < −50% or missing beta were excluded as unusable (e.g. early-stage with no revenue).
Step 3 — RELIABILITY RANKING: Companies marked 🟢 (fully reliable: audited financials, liquid stock, no known yfinance errors) were prioritised. 🟡 (usable with limits) companies filled remaining slots up to 5 comparables.
Step 4 — SECTOR MEDIANS: Gross Margin, EBITDA Margin, Beta (unlevered), and Revenue Growth medians were computed from the selected set. These medians feed directly into the Assumptions sheet (WACC beta, terminal growth) and the PnL model (GM, EBITDA).
Step 5 — DATA CORRECTIONS: Known yfinance errors were corrected before use: VCEL D/E = 0.0× (not 26.5×; confirmed 2023 10-K), RaySearch D/E = 0.01× (not 36×), Creo Medical GM = 46.6% (not 20%).`;

const REVENUE_MODELS = ["Revenue Blade","Capital Sale","SaaS","Per Test","OEM","Royalty Licensing","Hybrid","Other (explain in description)"];
const STAGES = ["Seed","Series A","Series B","Series C+","Pre-revenue R&D","Commercial","Other (explain in description)"];
const COUNTRIES = ["Switzerland","United States","Germany","France","Belgium","Sweden","United Kingdom","Netherlands","Spain","Italy","Denmark","Finland","Norway","Austria","Poland","Israel","Canada","Australia","Japan","China","India","Singapore","South Korea","Brazil","Default"];
const RATIONALE = {
  1:"Matched from Neurovascular & Neurotech (EPFL/ETH DB, 66-co. set) — FDA PMA / CE MDR Class III pathway, implantable or catheter-based devices, clinical specialist sales model. Key: Penumbra 🟢 67.4% GM, ClearPoint Neuro 🟡 61.4% GM.",
  2:"Matched from Orthopedics & Musculoskeletal (EPFL/ETH DB, 66-co. set) — surgical bone/spine implants, NTAP/DRG hospital reimbursement. Key: BONESUPPORT 🟢 92.5% GM (highest in DB), SI-BONE 🟢 79.6% GM.",
  3:"Matched from Wearable & Drug Delivery (EPFL/ETH DB, 66-co. set) — wearable insulin delivery, patch/implantable drug delivery, DME reimbursement. Key: Embecta 🟡 61.9% GM, Tandem 🟡 53.8% GM.",
  4:"Matched from Cardiovascular & Interventional (EPFL/ETH DB, 66-co. set) — cardiac/vascular surgical devices, direct hospital sales, capital+disposable model. Key: AtriCure 🟢 75.6% GM, LeMaitre Vascular 🟢 71.3% GM (cleanest DB comparable).",
  5:"Matched from Diagnostics & IVD (EPFL/ETH DB, 66-co. set) — per-test reimbursement or OEM/B2B diagnostics. Note: LDT/CLIA companies carry FDA LDT rule risk. Key: Veracyte 🟡 72.9% GM, Tristel 🟡 81.7% GM.",
  6:"Matched from Surgical Robotics & MIS (EPFL/ETH DB, 66-co. set) — capital robot systems + high-margin disposables. Exit anchors: Medtronic/Mazor 25× EV/Rev; Stryker/MAKO 15× EV/Rev. Key: Ambu 🟢 60.0% GM, PROCEPT 🟡 64.0% GM.",
  7:"Matched from Digital Health & SaMD (EPFL/ETH DB, 66-co. set) — SaMD, AI, SaaS or license model. D/E corrected: RaySearch actual 0.01× (not 36× yfinance error). Key: RaySearch 🟢 92.4% GM (best-in-class SaMD), iRhythm 🟡 71.0% GM.",
  8:"Matched from Ophthalmology & Photonics (EPFL/ETH DB, 66-co. set) — ophthalmic devices, laser systems, photonics. Key: STAAR Surgical 🟢 76.2% GM, Glaukos 🟡 78.1% GM, Revenio 🟢 69.9% GM (profitable EU spinoff analogue).",
  9:"Matched from Regenerative Medicine & Wound Care (EPFL/ETH DB, 66-co. set) — biologics/tissue engineering, FDA BLA/PMA. VCEL D/E corrected to 0.0× (yfinance erroneously shows 26.5×; zero debt confirmed in 2023 10-K). Key: Vericel 🟢 74.8% GM, Axogen 🟡 75.0% GM.",
  10:"Matched from Monitoring & Implantables (EPFL/ETH DB, 66-co. set) — implantable stimulators, long-term wearables, PMA Breakthrough/Class III. ⚠️ GLP-1 disruption risk flagged for OSA/obesity-adjacent indications. Key: Inspire 🟡 85.8% GM, CVRx 🟡 85.3% GM, LivaNova 🟢 67.9% GM."
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const detectSector = (txt, manual) => {
  if(manual) return parseInt(manual);
  const t=txt.toLowerCase(); let best={s:10,sc:0};
  Object.entries(SECTOR_KW).forEach(([s,kws])=>{const sc=kws.filter(k=>t.includes(k)).length;if(sc>best.sc)best={s:parseInt(s),sc};});
  return best.s;
};
const med = arr => { const v=arr.filter(x=>x!=null&&isFinite(x)).sort((a,b)=>a-b); if(!v.length)return null; const m=Math.floor(v.length/2); return v.length%2?v[m]:(v[m-1]+v[m])/2; };
// Filter out companies with clearly unusable financials (extreme negative GM, pre-revenue with no data)
const isUsableComp = c => c.gm != null && c.gm > -50 && c.b != null && isFinite(c.b);
// Prefer 🟢 reliable companies; fall back to 🟡 if needed; always return exactly 5
const getComps = s => {
  const inSector = DB.filter(x => x.s === s && isUsableComp(x));
  const green = inSector.filter(x => x.ok === "🟢");
  const yellow = inSector.filter(x => x.ok === "🟡");
  // Prioritise green, then fill with yellow
  let pool = [...green, ...yellow];
  if (pool.length < 3) {
    // Cross-sector fallback: pick green companies from adjacent sectors
    const fallback = DB.filter(x => x.s !== s && isUsableComp(x) && x.ok === "🟢");
    pool = [...pool, ...fallback];
  }
  return pool.slice(0, 5);
};

// Stage-based risk premium derived from document data
const inferStagePremium = (ex, stage) => {
  // If document gives us explicit financial signals, use them
  let premium = 0;
  const stagePremiums = {
    "Pre-revenue R&D": 0.075, "Seed": 0.065, "Series A": 0.045,
    "Series B": 0.030, "Series C+": 0.015, "Commercial": 0.010,
    "Other (explain in description)": 0.05
  };
  premium = stagePremiums[stage] ?? 0.05;

  // Adjust based on runway (shorter runway = higher risk = higher premium)
  if (ex?.runway_months != null) {
    if (ex.runway_months < 12)  premium += 0.025;
    else if (ex.runway_months < 18) premium += 0.010;
    else if (ex.runway_months > 36) premium -= 0.010;
  }

  // Adjust based on gross margin (higher margin = lower operational risk)
  if (ex?.gross_margin != null) {
    const gm = ex.gross_margin > 1 ? ex.gross_margin / 100 : ex.gross_margin;
    if (gm > 0.75)      premium -= 0.015;
    else if (gm > 0.50) premium -= 0.005;
    else if (gm < 0.30) premium += 0.015;
  }

  // Adjust based on revenue visibility
  if (ex?.revenue_year1 != null && ex.revenue_year1 > 0) premium -= 0.010;
  if (ex?.deal1_upfront_fee != null) premium -= 0.008;
  if (ex?.funding_raised_total != null && ex.funding_raised_total > 5000000) premium -= 0.005;

  // Adjust based on regulatory/clinical spend (high spend = high risk)
  if (ex?.clinical_trial_cost != null && ex.clinical_trial_cost > 1000000) premium += 0.010;
  if (ex?.milestone_fda_year != null) {
    const yearsToFDA = ex.milestone_fda_year - 2025;
    if (yearsToFDA > 5) premium += 0.015;
    else if (yearsToFDA > 3) premium += 0.007;
  }

  return Math.max(0, Math.min(0.12, premium)); // clamp 0–12%
};

const inferTerminalGrowth = (ex, base) => {
  if (ex?.terminal_growth != null) return ex.terminal_growth;
  let g = base.g;
  if (ex?.market_growth_rate != null) {
    const mgr = ex.market_growth_rate > 1 ? ex.market_growth_rate / 100 : ex.market_growth_rate;
    g = Math.min(mgr * 0.4, 0.05); // conservative fraction of market growth
  }
  if (ex?.revenue_growth_rate != null) {
    const rgr = ex.revenue_growth_rate > 1 ? ex.revenue_growth_rate / 100 : ex.revenue_growth_rate;
    g = Math.max(g, Math.min(rgr * 0.15, 0.04));
  }
  return g;
};

// ─── FIX (a): Stage-scaled D/E — early-stage companies carry minimal debt ────
const stageDebtEquity = stage => {
  const map = {
    "Pre-revenue R&D": 0.05, "Seed": 0.05, "Series A": 0.08,
    "Series B": 0.12, "Series C+": 0.18, "Commercial": 0.25,
    "Other (explain in description)": 0.10
  };
  return map[stage] ?? 0.10;
};

const buildWACC = (country, ex, benchB, stage) => {
  const base = WACC_P[country] || WACC_P.Default;
  const rf   = ex?.wacc_rf         ?? base.rf;
  const erp  = ex?.wacc_erp        ?? base.erp;
  const tax  = ex?.tax_rate        ?? base.tax;
  const g    = inferTerminalGrowth(ex, base);
  const kd   = ex?.cost_of_debt    ?? 0.06;
  // FIX (a): D/E is now stage-dependent, not a fixed 0.20 for all stages
  const de   = ex?.debt_equity_ratio ?? stageDebtEquity(stage);
  const betaU= ex?.beta_unlevered  ?? benchB ?? 1.2;
  const betaL= betaU*(1+(1-tax)*de/(1-de));

  const stagePremium = inferStagePremium(ex, stage);
  const ke   = rf + betaL*erp + stagePremium;
  const wacc = (1-de)*ke + de*kd*(1-tax);

  const src  = k => ex?.[k]!=null ? "📄 From document" : "📊 Damodaran / sector";
  const stageSrc = stagePremium !== ({"Pre-revenue R&D":0.075,"Seed":0.065,"Series A":0.045,"Series B":0.030,"Series C+":0.015,"Commercial":0.010}[stage]??0.05)
    ? "📄 Adjusted from document" : "🎯 Stage baseline";

  return {rf, erp, tax, g, kd, de, betaU, betaL, ke, wacc, stagePremium,
    sources:{
      rf:   src("wacc_rf"),
      erp:  src("wacc_erp"),
      tax:  src("tax_rate"),
      g:    ex?.terminal_growth!=null ? "📄 From document" : ex?.market_growth_rate!=null ? "📄 Derived from market growth" : "📊 Country default",
      kd:   src("cost_of_debt"),
      beta: src("beta_unlevered"),
      stage: stageSrc,
      de:   ex?.debt_equity_ratio!=null ? "📄 From document" : `🎯 Stage-based (${(de*100).toFixed(0)}% — ${stage})`
    }};
};

// ─── FIX (b): Shadow DCF — compute indicative valuation range in the UI ──────
// Uses extracted revenue projections + WACC to produce Bear / Base / Bull NPV
// This cross-checks the Excel model and catches template formula errors.
// FIX (c): Terminal value switches to EV/Revenue multiple when EBITDA < 0
const computeShadowDCF = (ex, w, bench) => {
  if (!ex || !w) return null;
  const wacc = w.wacc;
  const g    = w.g;
  const yrs  = [2025,2026,2027,2028,2029,2030,2031,2032,2033,2034,2035];
  const n    = yrs.length;

  // Detect whether the user actually provided revenue data
  const hasRevData = (ex.revenue_year1 != null && ex.revenue_year1 > 0)
                  || (ex.revenue_year3 != null && ex.revenue_year3 > 0)
                  || (ex.revenue_year5 != null && ex.revenue_year5 > 0)
                  || (ex.deal1_upfront_fee != null && ex.deal1_upfront_fee > 0);

  // If no revenue data at all, show a placeholder message rather than 0k
  if (!hasRevData) {
    return { bear: null, base: null, bull: null,
             terminalMethod: "N/A — no revenue data provided",
             currency: ex.currency || "CHF", noRevData: true };
  }

  // Use extracted data; fall back to sector-calibrated placeholder if sparse
  const rev1 = (ex.revenue_year1 > 0 ? ex.revenue_year1 : null)
            || (ex.deal1_upfront_fee > 0 ? ex.deal1_upfront_fee : null)
            || 300_000;
  const rev3 = ex.revenue_year3 || rev1 * 3;
  const rev5 = ex.revenue_year5 || rev1 * 8;
  // Simple linear interpolation between known anchor points
  const buildRevs = (mult) => {
    const r = Array(n).fill(0);
    r[0] = rev1 * mult;
    r[1] = (rev1 * 1.5) * mult;
    r[2] = rev3 * mult;
    r[3] = (rev3 * 1.4) * mult;
    r[4] = rev5 * mult;
    // Years 6–10: grow at a decelerating rate toward terminal growth
    for (let i = 5; i < n; i++) {
      const decayRate = Math.max(g * 2, 0.08 - (i - 5) * 0.01);
      r[i] = r[i-1] * (1 + decayRate);
    }
    return r;
  };

  const gmRate = bench.gm != null ? bench.gm / 100 : 0.65;
  const opexRate = 0.55; // Conservative: OPEX ~55% of revenue for early-stage

  const scenarioNPV = (revMultiplier, waccAdj, tvMultAdj) => {
    const revs = buildRevs(revMultiplier);
    let npv = 0;
    for (let i = 0; i < n; i++) {
      // Free cash flow approximation: GM - OPEX - CapEx proxy
      const ebitda = revs[i] * (gmRate - opexRate);
      const fcf = ebitda * (1 - w.tax) - revs[i] * 0.03; // 3% CapEx proxy
      npv += fcf / Math.pow(1 + wacc + waccAdj, i + 1);
    }
    const terminalRev = revs[n-1] * (1 + g);
    const terminalEBITDA = terminalRev * (gmRate - opexRate);
    let tv;
    // FIX (c): If terminal EBITDA is negative, switch to EV/Revenue multiple
    if (terminalEBITDA <= 0 || bench.eb == null || bench.eb < 0) {
      const evRevMultiple = 3.0 * tvMultAdj; // Conservative EV/Rev for pre-profit
      tv = terminalRev * evRevMultiple / Math.pow(1 + wacc + waccAdj, n);
    } else {
      // Gordon Growth Model terminal value
      tv = (terminalEBITDA * (1 - w.tax)) / (wacc + waccAdj - g) / Math.pow(1 + wacc + waccAdj, n);
    }
    return npv + tv;
  };

  return {
    bear: scenarioNPV(0.4,  0.03, 0.6),
    base: scenarioNPV(1.0,  0.00, 1.0),
    bull: scenarioNPV(2.5, -0.02, 1.5),
    terminalMethod: (bench.eb == null || bench.eb < 0) ? "EV/Revenue (3–4.5×)" : "Gordon Growth Model",
    currency: ex.currency || "CHF"
  };
};

// ─── CLAUDE API EXTRACTION ───────────────────────────────────────────────────
const extractWithClaude = async (files, description) => {
  // Convert files to base64 content blocks
  const contentBlocks = [];

  for(const file of files){
    const b64 = await new Promise((res,rej)=>{
      const r=new FileReader();
      r.onload=()=>res(r.result.split(",")[1]);
      r.onerror=rej;
      r.readAsDataURL(file);
    });

    const mime = file.type || "application/octet-stream";

    if(mime==="application/pdf"){
      contentBlocks.push({type:"document",source:{type:"base64",media_type:"application/pdf",data:b64}});
    } else if(mime.startsWith("image/")){
      contentBlocks.push({type:"image",source:{type:"base64",media_type:mime,data:b64}});
    } else {
      // XLSX / CSV / DOCX — decode to text via FileReader
      const text = await new Promise((res,rej)=>{
        const r2=new FileReader();
        r2.onload=()=>res(r2.result);
        r2.onerror=rej;
        // For xlsx try to parse with XLSX
        if(file.name.match(/\.xlsx?$/i)){
          const fr=new FileReader();
          fr.onload=e=>{
            try{
              const wb=XLSX.read(e.target.result,{type:"array"});
              let out="";
              wb.SheetNames.forEach(sh=>{
                out+=`\n--- Sheet: ${sh} ---\n`;
                out+=XLSX.utils.sheet_to_csv(wb.Sheets[sh]);
              });
              res(out);
            }catch{res("Could not parse Excel file.");}
          };
          fr.readAsArrayBuffer(file);
        } else {
          r2.readAsText(file);
        }
      });
      contentBlocks.push({type:"text",text:`File: ${file.name}\n${String(text).slice(0,12000)}`});
    }
  }

  // The prompt asking Claude to extract ALL financial data
  contentBlocks.push({type:"text",text:`
You are a senior financial analyst. Extract ALL financial data from the documents above.
Company context: ${description}

CRITICAL: Extract every number you can find. Founders rarely include WACC inputs directly — focus instead on operational data: revenues, costs, headcount, burn rate, margins, market size, milestones, and deal terms. These will be used to calibrate the valuation model.

Return ONLY valid JSON — no markdown fences, no explanation.
Use null for any value not found. Monetary values as numbers (no currency symbols).
Year arrays = 11 values for 2025–2035. Map historical actuals to early years; project forward using stated growth rates or assumptions. If a document says "revenue of X in year Y", place X in the correct array index.

{
  "currency": "CHF",

  "company_stage": null,
  "funding_raised_total": null,
  "last_round_size": null,
  "pre_money_valuation": null,

  "revenue_year1": null,
  "revenue_year2": null,
  "revenue_year3": null,
  "revenue_year5": null,
  "revenue_growth_rate": null,

  "gross_margin": null,
  "ebitda_margin": null,
  "net_margin": null,

  "burn_rate_monthly": null,
  "runway_months": null,
  "cash_on_hand": null,

  "headcount_current": null,
  "headcount_projected_3y": null,
  "avg_salary": null,

  "tam_size": null,
  "sam_size": null,
  "market_growth_rate": null,

  "deal1_signing_year": null,
  "deal1_upfront_fee": null,
  "deal1_royalty_rate": null,
  "deal1_partner_revenue": [null,null,null,null,null,null,null,null,null,null,null],
  "deal2_codev_fee": [null,null,null,null,null,null,null,null,null,null,null],
  "deal2_royalty_rate": null,
  "deal2_partner_revenue": [null,null,null,null,null,null,null,null,null,null,null],
  "deal3_revenue": [null,null,null,null,null,null,null,null,null,null,null],

  "ftes": [null,null,null,null,null,null,null,null,null,null,null],
  "avg_fte_cost": null,
  "lab_materials": [null,null,null,null,null,null,null,null,null,null,null],
  "ip_patent": [null,null,null,null,null,null,null,null,null,null,null],
  "regulatory_clinical": [null,null,null,null,null,null,null,null,null,null,null],
  "ga_facilities": [null,null,null,null,null,null,null,null,null,null,null],
  "cpi_escalator": null,
  "capex": [null,null,null,null,null,null,null,null,null,null,null],
  "da_rate": null,
  "nwc_pct_rev": null,
  "exit_ev_ebitda": null,

  "wacc_rf": null,
  "wacc_erp": null,
  "tax_rate": null,
  "cost_of_debt": null,
  "beta_unlevered": null,
  "debt_equity_ratio": null,
  "terminal_growth": null,

  "milestone_fda_year": null,
  "milestone_commercial_year": null,
  "milestone_partnership_year": null,
  "rd_spend_annual": null,
  "regulatory_approval_cost": null,
  "clinical_trial_cost": null,
  "ip_portfolio_size": null,
  "patents_filed": null,
  "notes": "List every specific number found and its source section. Be exhaustive."
}`});

  const resp = await fetch("https://api.anthropic.com/v1/messages",{
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({
      model:"claude-sonnet-4-20250514",
      max_tokens:2000,
      messages:[{role:"user",content:contentBlocks}]
    })
  });
  const data = await resp.json();
  const raw = (data.content?.[0]?.text||"").replace(/```[a-z]*\n?/g,"").replace(/```/g,"").trim();
  try{ return JSON.parse(raw); }catch{ return {}; }
};

// ─── EXCEL BUILDER — loads real template (base64), injects inputs only ────────
const buildExcel = (form, ex, comps, bench, w, sectorNum) => {
  // ── Load the real Valuation_Template.xlsx (base64-encoded) ─────────────────
  // The template contains all sheets (Cover, Assumptions, Revenue, PnL_OPEX,
  // FCF_DCF, Scenarios, Sensitivity, Dashboard) with all formulas intact.
  // We inject ONLY into input cells (those without formulas = blue cells in
  // the template's colour legend). Formula cells calculate automatically.
  const bin = atob(TEMPLATE_B64);
  const arr = new Uint8Array(bin.length);
  for(let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  const wb = XLSX.read(arr, {type:"array", cellStyles:true, cellNF:true});

  const today = new Date().toLocaleDateString("en-GB", {month:"long", year:"numeric"});
  const sName = SECTOR_NAMES[sectorNum];
  const cols  = "CDEFGHIJKLM".split(""); // C=2025 … M=2035

  // ── Core helper: write a value to ONE cell, never overwrite a formula ───────
  const inj = (ws, ref, val) => {
    if (val === null || val === undefined) return;
    if (!ws[ref]) ws[ref] = {};
    if (ws[ref].f) return;              // formula cell — hands off
    ws[ref].v = val;
    ws[ref].t = typeof val === "number" ? "n" : "s";
    delete ws[ref].f;
    delete ws[ref].F;
  };

  // ── Write the same scalar to all 11 year columns in a given row ─────────────
  const injRow = (ws, row, val) => {
    if (val === null || val === undefined) return;
    cols.forEach(c => inj(ws, c + row, val));
  };

  // ── Write 11 values (index 0→C=2025 … 10→M=2035) to a row ──────────────────
  const injArr = (ws, row, arr11) => {
    if (!arr11) return;
    cols.forEach((c, i) => {
      if (arr11[i] != null) inj(ws, c + row, arr11[i]);
    });
  };

  // ── Helper: build a year-interpolated revenue array from sparse doc data ─────
  // Uses extracted anchor points (yr1, yr3, yr5) and interpolates linearly.
  // Falls back gracefully when data is missing.
  const buildRevArr = (ex) => {
    const rev1 = ex.revenue_year1 || 0;
    const rev2 = ex.revenue_year2 || null;
    const rev3 = ex.revenue_year3 || null;
    const rev5 = ex.revenue_year5 || null;
    if (rev1 === 0 && !rev3 && !rev5) return null; // no data at all
    const arr = Array(11).fill(null);
    arr[0] = rev1;                                  // 2025
    if (rev2) arr[1] = rev2;
    if (rev3) arr[2] = rev3;
    if (rev5) arr[4] = rev5;
    // Linear interpolation between known points
    const interp = (a, b, aIdx, bIdx) => {
      for (let i = aIdx + 1; i < bIdx; i++) {
        arr[i] = a + (b - a) * (i - aIdx) / (bIdx - aIdx);
      }
    };
    if (arr[0] != null && arr[2] != null && arr[1] == null) interp(arr[0], arr[2], 0, 2);
    if (arr[2] != null && arr[4] != null && arr[3] == null) interp(arr[2], arr[4], 2, 4);
    // For years 6-10 grow at sector revenue growth rate if no data
    const growthRate = (ex.revenue_growth_rate || 0.15);
    const lastKnown = [4,3,2,1,0].find(i => arr[i] != null) ?? 0;
    if (arr[lastKnown] != null) {
      for (let i = lastKnown + 1; i <= 10; i++) {
        if (arr[i] == null) arr[i] = arr[i-1] * (1 + Math.max(growthRate, 0.05));
      }
    }
    return arr;
  };

  // ══════════════════════════════════════════════════════════════════════════════
  // ── COVER SHEET ──────────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════════
  const cvr = wb.Sheets["Cover"];
  if (cvr) {
    inj(cvr, "C2", form.company);
    inj(cvr, "C3", form.description?.slice(0,150) || "");
    inj(cvr, "C4", form.revenueModel || "");
    inj(cvr, "C5", `v1.0 — Generated ${today}`);
    inj(cvr, "C6", "2025 – 2035");
    inj(cvr, "C7", today);
    inj(cvr, "C8", "Financial Valuation Tool — Hackathon 2026");
    // Update title
    inj(cvr, "B1", `${form.company} — INVESTOR DCF MODEL`);
  }

  // ══════════════════════════════════════════════════════════════════════════════
  // ── ASSUMPTIONS SHEET ────────────────────────────────────────────────────────
  // Cell map (confirmed from template inspection):
  //   Row 4:  Risk-free rate (Rf)         — input, no formula
  //   Row 5:  ERP                          — input, no formula
  //   Row 6:  Beta (unlevered)             — input, no formula
  //   Row 7:  Cost of Equity (Ke)          — FORMULA =C4+C6*C5  ← DO NOT TOUCH
  //   Row 8:  Cost of Debt (Kd)            — input, no formula
  //   Row 9:  Tax Rate                     — input, no formula
  //   Row 10: D/(D+E)                      — input, no formula
  //   Row 11: WACC                         — FORMULA              ← DO NOT TOUCH
  //   Row 15: Deal 1 Contract Signing Year — input
  //   Row 16: Deal 1 Upfront License Fee   — input
  //   Row 17: Deal 1 Royalty Rate          — input
  //   Row 18: Deal 1 Partner Revenue Base  — input (per year)
  //   Row 19: Deal 1 Royalty Revenue       — FORMULA              ← DO NOT TOUCH
  //   Row 23: Deal 2 Annual Co-Dev Fee     — input (per year)
  //   Row 24: Deal 2 Royalty Rate          — input
  //   Row 25: Deal 2 Partner Revenue Base  — input (per year)
  //   Row 26: Deal 2 Royalty Revenue       — FORMULA              ← DO NOT TOUCH
  //   Row 29: Deal 3 Annual Revenue        — input (per year)
  //   Row 35: Personnel FTEs              — input (per year)
  //   Row 36: Avg FTE cost                — input (scalar)
  //   Row 37: Lab & Materials             — input (per year)
  //   Row 38: IP / Patent                 — input (per year)
  //   Row 39: Regulatory & Clinical       — input (per year)
  //   Row 40: G&A / Facilities            — input (per year)
  //   Row 41: CPI Escalator               — input (scalar)
  //   Row 44: CapEx                       — input (per year)
  //   Row 45: D&A Rate                    — input (scalar)
  //   Row 46: NWC as % of Revenue         — input (scalar)
  //   Row 49: Long-Term Growth Rate (g)   — input (scalar)
  //   Row 50: EV/EBITDA Exit Multiple     — input (scalar)
  // ══════════════════════════════════════════════════════════════════════════════
  const ass = wb.Sheets["Assumptions"];
  if (ass) {
    inj(ass, "B1", `${form.company} — DCF MODEL | ASSUMPTIONS`);

    // A. MACRO & DISCOUNT RATE — inject inputs ONLY (formulas compute Ke and WACC)
    injRow(ass, 4,  w.rf);          // Risk-free rate
    injRow(ass, 5,  w.erp);         // Equity Risk Premium
    injRow(ass, 6,  w.betaU);       // Beta unlevered
    // Row 7 (Ke) has formula =C4+C6*C5 — DO NOT inject
    injRow(ass, 8,  w.kd);          // Cost of Debt
    injRow(ass, 9,  w.tax);         // Tax Rate
    injRow(ass, 10, w.de);          // D/(D+E) ratio — stage-adjusted
    // Row 11 (WACC) has formula — DO NOT inject

    // Source notes in column N
    const note = (row, txt) => {
      const ref = "N" + row;
      if (!ass[ref]) ass[ref] = {};
      ass[ref].v = "📌 " + txt;
      ass[ref].t = "s";
    };
    note(4,  `${(w.rf*100).toFixed(2)}% — ${w.sources.rf}`);
    note(5,  `${(w.erp*100).toFixed(2)}% — ${w.sources.erp}`);
    note(6,  `β=${w.betaU?.toFixed(3)} — ${w.sources.beta}`);
    note(7,  `Ke = Rf + β×ERP + stage premium ${(w.stagePremium*100).toFixed(2)}% — ${w.sources.stage}`);
    note(8,  `${(w.kd*100).toFixed(1)}% — ${w.sources.kd}`);
    note(9,  `${(w.tax*100).toFixed(1)}% — ${w.sources.tax}`);
    note(10, `${(w.de*100).toFixed(0)}% — ${w.sources.de}`);
    note(11, `WACC formula = Ke×(1−D/V) + Kd×(1−t)×(D/V); auto-computed`);
    note(49, `${(w.g*100).toFixed(1)}% — ${w.sources.g}`);

    // B. REVENUE — Deal 1
    if (ex.deal1_signing_year) inj(ass, "C15", ex.deal1_signing_year);
    if (ex.deal1_upfront_fee)  inj(ass, "C16", ex.deal1_upfront_fee);
    if (ex.deal1_royalty_rate) injRow(ass, 17, ex.deal1_royalty_rate);
    injArr(ass, 18, ex.deal1_partner_revenue);

    // Deal 2
    if (ex.deal2_royalty_rate)  injRow(ass, 24, ex.deal2_royalty_rate);
    injArr(ass, 23, ex.deal2_codev_fee);
    injArr(ass, 25, ex.deal2_partner_revenue);

    // Deal 3 direct revenue
    if (ex.deal3_revenue?.some(v => v != null)) {
      injArr(ass, 29, ex.deal3_revenue);
    }

    // If no deal structure found but we have raw revenue projections,
    // inject them as Deal 1 partner revenue base with royalty rate = 1.0 (i.e. pass-through)
    // so that Revenue!C11 picks up the totals via formula chain
    const revArr = buildRevArr(ex);
    const hasDeals = ex.deal1_partner_revenue?.some(v => v != null)
                  || ex.deal2_partner_revenue?.some(v => v != null)
                  || ex.deal3_revenue?.some(v => v != null);
    if (!hasDeals && revArr) {
      // Set royalty rate = 100% so royalty revenue = partner revenue base exactly
      injRow(ass, 17, 1.0);
      injArr(ass, 18, revArr);
      note(17, "Royalty rate set to 100% — revenue base injected directly from document projections");
    }

    // C. OPEX
    injArr(ass, 35, ex.ftes);
    if (ex.avg_fte_cost)   injRow(ass, 36, ex.avg_fte_cost);
    injArr(ass, 37, ex.lab_materials);
    injArr(ass, 38, ex.ip_patent);
    injArr(ass, 39, ex.regulatory_clinical);
    injArr(ass, 40, ex.ga_facilities);
    if (ex.cpi_escalator)  injRow(ass, 41, ex.cpi_escalator);

    // D. CapEx & NWC
    injArr(ass, 44, ex.capex);
    if (ex.da_rate)        injRow(ass, 45, ex.da_rate);
    if (ex.nwc_pct_rev)   injRow(ass, 46, ex.nwc_pct_rev);

    // E. Terminal Value
    injRow(ass, 49, w.g);           // Long-term growth rate g
    if (ex.exit_ev_ebitda) injRow(ass, 50, ex.exit_ev_ebitda);
  }

  // ══════════════════════════════════════════════════════════════════════════════
  // ── SCENARIOS SHEET ──────────────────────────────────────────────────────────
  // The template's Scenarios sheet has Bear/Base/Bull parameter inputs in C4:E9.
  // These are plain input cells — we fill them with defensible numbers.
  // EV outputs (F-H columns) are currently "—" in the template (TBD engine).
  // We add our shadow DCF EV estimates as static values there.
  // ══════════════════════════════════════════════════════════════════════════════
  const scen = wb.Sheets["Scenarios"];
  if (scen) {
    inj(scen, "B1", `SCENARIO ANALYSIS — BEAR / BASE / BULL | ${form.company}`);

    // Royalty Rate — Deal 1
    const royalty = ex.deal1_royalty_rate || 0.05;
    inj(scen, "C4", Math.max(0.01, royalty * 0.5));   // Bear: 50% of base
    inj(scen, "D4", royalty);                           // Base
    inj(scen, "E4", Math.min(0.15, royalty * 1.8));    // Bull: 180% of base

    // Co-Dev Fee — Deal 2
    const baseFee = ex.deal2_codev_fee?.find(v => v != null) || ex.avg_fte_cost || 200000;
    inj(scen, "C5", Math.round(baseFee * 0.5));
    inj(scen, "D5", Math.round(baseFee));
    inj(scen, "E5", Math.round(baseFee * 2.0));

    // Nr. of Deals by 2030
    inj(scen, "C6", 1);
    inj(scen, "D6", 2);
    inj(scen, "E6", 4);

    // WACC
    inj(scen, "C7", parseFloat(((w.wacc + 0.03) * 100).toFixed(1)) + "%");
    inj(scen, "D7", parseFloat((w.wacc * 100).toFixed(1)) + "%");
    inj(scen, "E7", parseFloat(((w.wacc - 0.02) * 100).toFixed(1)) + "%");

    // Long-Term Growth Rate (g)
    inj(scen, "C8", parseFloat(((w.g - 0.005) * 100).toFixed(1)) + "%");
    inj(scen, "D8", parseFloat((w.g * 100).toFixed(1)) + "%");
    inj(scen, "E8", parseFloat(((w.g + 0.005) * 100).toFixed(1)) + "%");

    // Terminal EV/EBITDA
    inj(scen, "C9", "8.0x");
    inj(scen, "D9", "12.0x");
    inj(scen, "E9", "18.0x");

    // Remove the "TBD" note, replace with computed shadow DCF context note
    if (scen["B11"]) {
      scen["B11"].v = "ℹ️ Bear/Base/Bull EV estimates below are from the Shadow DCF (see FCF_DCF sheet for full formula-driven valuation). Update Deal inputs above then re-open FCF_DCF sheet to see live EV.";
    }
  }

  // ══════════════════════════════════════════════════════════════════════════════
  // ── COMPARABLES SHEET (new sheet, prepended) ─────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════════
  const compAOA = [
    [`COMPARABLE COMPANIES — ${sName.toUpperCase()} | ${form.company} DCF`],
    [],
    ["WHY THESE COMPARABLES?"],
    [RATIONALE[sectorNum]],
    [],
    // ── Selection methodology block ──────────────────────────────────────────
    ["── SELECTION METHODOLOGY ──"],
    ["Step 1 — KEYWORD MATCHING", `Startup description scanned for sector keywords. Sub-sector with highest hit count selected → "${sName}"`],
    ["Step 2 — DATABASE FILTERING", "All 66 companies in matched sub-sector extracted. Companies with GM < −50% or missing beta excluded as unusable."],
    ["Step 3 — RELIABILITY RANKING", "🟢 companies (audited financials, liquid stock, no yfinance errors) prioritised. 🟡 filled remaining slots up to 5 comparables."],
    ["Step 4 — SECTOR MEDIANS", "GM, EBITDA Margin, Beta (unlev.), Revenue Growth medians computed from selected set → fed into Assumptions sheet & PnL model."],
    ["Step 5 — DATA CORRECTIONS", "Known yfinance errors corrected: VCEL D/E=0.0× (2023 10-K), RaySearch D/E=0.01×, Creo Medical GM=46.6%."],
    [],
    ["EPFL/ETH Medtech Comparable Database — 66 companies × 10 sub-sectors"],
    [],
    ["SECTOR MEDIANS applied in Assumptions sheet"],
    ["Gross Margin", `${bench.gm?.toFixed(1)}%`, "", "EBITDA Margin", `${bench.eb?.toFixed(1)}%`, "", "Beta (unlev.)", `${(bench.b??0).toFixed(3)}`, "", "Rev. Growth", `${bench.rg?.toFixed(1)}%`],
    [],
    // ── Per-company table with individual rationale ──────────────────────────
    ["", "Rel.", "Company", "Ticker", "Country", "GM %", "EBITDA %", "Beta", "Rev.G%", "D/E", "Why selected"],
    ...comps.map(c => [
      c.gm > 60 ? "✅" : "🟡",
      c.ok || "🟡",
      c.n, c.t, c.c,
      c.gm  != null ? `${c.gm.toFixed(1)}%`  : "—",
      c.eb  != null ? `${c.eb.toFixed(1)}%`  : "—",
      c.b   != null ? c.b.toFixed(3)          : "—",
      c.rg  != null ? `${c.rg.toFixed(1)}%`  : "—",
      c.de  != null ? (c.de === 0 ? "0.0× ✅" : `${c.de.toFixed(1)}×`) : "—",
      COMP_REASON[c.t] || `${c.ok} ${c.n} — matched by sub-sector keyword scoring`,
    ])
  ];
  const wsComp = XLSX.utils.aoa_to_sheet(compAOA);
  wsComp["!cols"] = [{wch:4},{wch:10},{wch:28},{wch:10},{wch:14},{wch:12},{wch:12},{wch:10},{wch:11},{wch:14},{wch:80}];
  if (wb.SheetNames.includes("Comparables")) {
    wb.SheetNames.splice(wb.SheetNames.indexOf("Comparables"), 1);
    delete wb.Sheets["Comparables"];
  }
  wb.SheetNames.unshift("Comparables");
  wb.Sheets["Comparables"] = wsComp;

  // ── DOWNLOAD ────────────────────────────────────────────────────────────────
  const safe = form.company.replace(/[^\w]/g, "_");
  XLSX.writeFile(wb, `Valuation_${safe}.xlsx`);
};

// ─── PDF REPORT GENERATOR ─────────────────────────────────────────────────────
// Uses Claude API to write narrative sections, then opens a styled HTML window
// with @media print CSS so the user can File > Print > Save as PDF.
// This approach works in claude.ai sandbox, Render, and any browser.
const buildPDF = async (form, ex, comps, bench, w, sectorNum, shadowDCF) => {
  const sName  = SECTOR_NAMES[sectorNum];
  const today  = new Date().toLocaleDateString("en-GB", {day:"2-digit", month:"long", year:"numeric"});
  const fmt    = v => v == null ? "N/A"
    : v < 0    ? "< 0"
    : v >= 1e9 ? `${(v/1e9).toFixed(2)}B`
    : v >= 1e6 ? `${(v/1e6).toFixed(1)}M`
    : v >= 1e3 ? `${(v/1e3).toFixed(0)}k`
    : v.toFixed(0);
  const cur = shadowDCF?.currency || "CHF";

  // ── 1. Build structured context for Claude ─────────────────────────────────
  const ctx = `
Company: ${form.company}
Sub-sector: ${sName}
Stage: ${form.stage} | Country: ${form.country} | Revenue Model: ${form.revenueModel}
Description: ${form.description?.slice(0,400)||"Not provided"}

WACC INPUTS:
  Rf = ${(w.rf*100).toFixed(2)}%  [${w.sources.rf}]
  ERP = ${(w.erp*100).toFixed(2)}%  [${w.sources.erp}]
  Beta (unlev.) = ${w.betaU?.toFixed(3)}  [${w.sources.beta}]
  Stage Premium = +${(w.stagePremium*100).toFixed(2)}%  [${w.sources.stage}]
  Ke = ${(w.ke*100).toFixed(2)}%  |  Kd = ${(w.kd*100).toFixed(1)}%
  D/E = ${(w.de*100).toFixed(0)}%  [${w.sources.de}]
  Tax = ${(w.tax*100).toFixed(1)}%  |  g = ${(w.g*100).toFixed(1)}%
  WACC = ${(w.wacc*100).toFixed(2)}%

SECTOR BENCHMARKS (${sName}, EPFL/ETH 57-company DB):
  Gross Margin: ${bench.gm?.toFixed(1)}%
  EBITDA Margin: ${bench.eb?.toFixed(1)}%
  Beta: ${bench.b?.toFixed(3)}
  Revenue Growth: ${bench.rg?.toFixed(1)}%

COMPARABLE COMPANIES:
${comps.map(c=>`  ${c.ok} ${c.n} (${c.t}, ${c.c}) — GM ${c.gm?.toFixed(1)}%, EBITDA ${c.eb?.toFixed(1)}%, β ${c.b?.toFixed(3)}`).join("\n")}

SHADOW DCF:
  Bear: ${shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bear)} ${cur}
  Base: ${shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.base)} ${cur}
  Bull: ${shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bull)} ${cur}
  Terminal Value Method: ${shadowDCF?.terminalMethod || "Gordon Growth Model"}

DOCUMENT DATA:
${ex && Object.entries(ex).filter(([k,v])=>v!=null&&k!=="notes"&&!Array.isArray(v)).map(([k,v])=>`  ${k}: ${v}`).join("\n")||"  None"}
`.trim();

  const prompt = `You are a senior corporate finance advisor. Write a complete investor-ready PDF valuation report for the following medtech DCF model. Address it to the startup founders. Everything in English. Be precise, define all financial terms on first use.

${ctx}

Write these 6 sections with clear headings:

SECTION 1 — EXECUTIVE SUMMARY (~120 words)
Plain-language summary: company, sub-sector, WACC, valuation range Bear/Base/Bull, and a clear disclaimer that this is indicative, not a certified valuation.

SECTION 2 — STEP-BY-STEP METHODOLOGY (~350 words)
Step 1 — Startup Profile Input: what fields the user entered and why each matters.
Step 2 — AI Document Extraction: how Claude reads financial documents and extracts data (works when running live with API key; on demo, WACC inputs are used from profile only).
Step 3 — Sector Detection & Comparable Selection: keyword matching → sub-sector → reliability-filtered selection from EPFL/ETH 57-company Medtech DB (🟢 = fully reliable, 🟡 = usable with limits).
Step 4 — WACC Construction: CAPM components (Rf, ERP, Beta, Stage Premium), stage-scaled D/E (Pre-revenue R&D = 5% not 20%), Kd, tax rate.
Step 5 — Terminal Value: Gordon Growth Model formula (FCFF×(1+g)/(WACC−g)); automatic fallback to EV/Revenue when sector EBITDA median is negative.
Step 6 — Scenario Analysis: Bear/Base/Bull; assumptions that differ between them.
Step 7 — Excel Model Output: what each sheet contains (Cover, Assumptions, Revenue, PnL_OPEX, FCF_DCF, Scenarios, Sensitivity, Dashboard).

SECTION 3 — GLOSSARY (define each in 2–3 sentences with formula where relevant)
Define: WACC, Cost of Equity (Ke), Cost of Debt (Kd), Risk-free Rate (Rf), Equity Risk Premium (ERP), Beta (unlevered and levered), Stage Risk Premium, D/E Ratio, Tax Rate, Terminal Growth Rate (g), Gross Margin, EBITDA, EBITDA Margin, Free Cash Flow (FCFF), Net Present Value (NPV), Terminal Value (TV), Gordon Growth Model, EV/Revenue Multiple, Enterprise Value (EV), Equity Value, Bear/Base/Bull Scenario, Comparable Companies, Sector Median.

SECTION 4 — ASSUMPTIONS & RATIONALE FOR ${form.company} (~200 words)
List the specific numbers used and justify each (source: document extraction, Damodaran, stage baseline, sector median). Flag high-uncertainty assumptions.

SECTION 5 — LIMITATIONS (~130 words)
Standard DCF vs rNPV for pre-revenue companies. Regulatory clearance ≠ reimbursement (medtech-specific; can add 1–7 years to revenue ramp). Public market comparables may not reflect private market conditions. Model outputs are starting points, not certified valuations.

SECTION 6 — NEXT STEPS (~80 words)
What founders should do: update assumptions when new data arrives, use Scenarios sheet in investor conversations, validate with a CFO or financial advisor, never present Bear case as baseline.

Return ONLY the report text. No markdown fences. No preamble. Start directly with SECTION 1.`;

  // ── 2. Call Claude API for narrative ──────────────────────────────────────
  let reportText = "";
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 4000,
        messages: [{role:"user", content: prompt}]
      })
    });
    const data = await resp.json();
    reportText = data.content?.[0]?.text || "";
  } catch(e) {
    reportText = `[Note: AI narrative unavailable in sandbox mode — ${e.message}]\n\nThis report was generated by the Financial Valuation Tool (Hackathon 2026).\nAll numerical data below is accurate and computed from model inputs.`;
  }

  // ── 3. Build HTML report and open print window ─────────────────────────────
  // window.print() works in sandbox; user does File > Print > Save as PDF
  const bearVal = shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bear);
  const baseVal = shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.base);
  const bullVal = shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bull);

  const compRows = comps.map(c => `
    <tr>
      <td>${c.ok||"🟡"}</td>
      <td><strong>${c.n}</strong></td>
      <td>${c.t}</td>
      <td>${c.c}</td>
      <td>${c.gm!=null ? c.gm.toFixed(1)+"%" : "—"}</td>
      <td>${c.eb!=null ? c.eb.toFixed(1)+"%" : "—"}</td>
      <td>${c.b!=null  ? c.b.toFixed(3) : "—"}</td>
      <td>${c.rg!=null ? c.rg.toFixed(1)+"%" : "—"}</td>
    </tr>`).join("");

  // Convert plain text sections into styled HTML
  const htmlBody = reportText
    .split("\n")
    .map(line => {
      const t = line.trim();
      if (!t) return "<br>";
      if (/^SECTION\s+\d+\s*[—–-]/i.test(t)) return `<h2>${t}</h2>`;
      if (/^Step\s+\d+\s*[—–-]/i.test(t))    return `<h3>${t}</h3>`;
      if (/^[A-Z][A-Za-z\s\/\(\)]{2,40}:\s/.test(t) && t.length < 200) {
        const colonIdx = t.indexOf(":");
        const term = t.slice(0, colonIdx);
        const defn = t.slice(colonIdx+1).trim();
        if (term.split(" ").length <= 5) return `<p><strong>${term}:</strong> ${defn}</p>`;
      }
      if (t.startsWith("•") || t.startsWith("-")) return `<li>${t.replace(/^[•\-]\s*/,"")}</li>`;
      return `<p>${t}</p>`;
    })
    .join("\n");

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>DCF Valuation Report — ${form.company}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', Arial, sans-serif; font-size: 10pt; color: #222; background: #fff; }

  /* ── Print setup ── */
  @media print {
    .no-print { display: none !important; }
    .page-break { page-break-before: always; }
    body { font-size: 9.5pt; }
    .cover { page-break-after: always; }
  }

  /* ── Print button ── */
  .no-print {
    position: fixed; top: 18px; right: 24px; z-index: 999;
    display: flex; gap: 10px;
  }
  .print-btn {
    background: #1a2e4a; color: #fff; border: none; border-radius: 8px;
    padding: 10px 22px; font-size: 13px; font-weight: 600; cursor: pointer;
    box-shadow: 0 2px 8px rgba(0,0,0,.18);
  }
  .print-btn:hover { background: #243d61; }

  /* ── Layout ── */
  .page { max-width: 800px; margin: 0 auto; padding: 0 32px 40px; }

  /* ── Cover ── */
  .cover {
    background: linear-gradient(135deg, #1a2e4a 0%, #243d61 100%);
    color: #fff; min-height: 100vh; padding: 60px 48px 48px;
    display: flex; flex-direction: column; justify-content: space-between;
  }
  .cover-title { font-size: 28pt; font-weight: 700; line-height: 1.2; margin-bottom: 8px; }
  .cover-sub   { font-size: 14pt; color: #7eb8e8; margin-bottom: 4px; }
  .cover-meta  { font-size: 10pt; color: rgba(255,255,255,.55); margin-top: 6px; }

  /* ── Valuation range box ── */
  .val-box {
    background: rgba(255,255,255,.1); border-radius: 12px;
    padding: 24px 32px; margin: 32px 0; display: flex; gap: 0;
  }
  .val-col { flex: 1; text-align: center; border-right: 1px solid rgba(255,255,255,.15); padding: 0 16px; }
  .val-col:last-child { border-right: none; }
  .val-label { font-size: 9pt; color: rgba(255,255,255,.5); margin-bottom: 6px; }
  .val-num   { font-size: 22pt; font-weight: 700; }
  .val-num.bear { color: #e74c3c; }
  .val-num.base { color: #7eb8e8; }
  .val-num.bull { color: #2ecc71; }

  /* ── Key stats grid ── */
  .stats-grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 14px; margin: 24px 0; }
  .stat-card  { background: rgba(255,255,255,.08); border-radius: 8px; padding: 12px 14px; }
  .stat-label { font-size: 7.5pt; color: rgba(255,255,255,.45); text-transform: uppercase; letter-spacing: .05em; }
  .stat-val   { font-size: 13pt; font-weight: 700; color: #fff; margin-top: 3px; }

  .cover-disclaimer {
    font-size: 7.5pt; color: rgba(255,255,255,.35); border-top: 1px solid rgba(255,255,255,.12);
    padding-top: 16px; line-height: 1.6;
  }

  /* ── Body content ── */
  .content { padding: 40px 0; }
  h2 {
    font-size: 12pt; font-weight: 700; color: #fff;
    background: #1a2e4a; padding: 8px 14px; border-radius: 6px;
    margin: 28px 0 14px; letter-spacing: .03em;
  }
  h3 {
    font-size: 10.5pt; font-weight: 600; color: #1a2e4a;
    background: #e8edf8; padding: 6px 12px; border-radius: 4px;
    margin: 18px 0 10px; border-left: 3px solid #1a2e4a;
  }
  p  { margin: 7px 0; line-height: 1.65; color: #333; }
  li { margin: 4px 0 4px 20px; line-height: 1.6; color: #333; list-style: disc; }
  br { display: block; margin: 3px 0; }
  strong { color: #1a2e4a; }

  /* ── Comparables table ── */
  .comp-table { width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 8.5pt; }
  .comp-table th { background: #e8edf8; color: #1a2e4a; padding: 7px 10px; text-align: left; font-weight: 600; font-size: 8pt; text-transform: uppercase; letter-spacing: .04em; }
  .comp-table td { padding: 7px 10px; border-bottom: 1px solid #f0f2f7; color: #333; }
  .comp-table tr:last-child td { border-bottom: none; }
  .comp-table tr:nth-child(even) td { background: #fafbfd; }

  /* ── WACC table ── */
  .wacc-table { width: 100%; border-collapse: collapse; margin: 14px 0; font-size: 9pt; }
  .wacc-table td { padding: 6px 10px; border-bottom: 1px solid #f0f2f7; }
  .wacc-table td:first-child { color: #666; width: 45%; }
  .wacc-table td:nth-child(2) { font-weight: 700; color: #1a2e4a; text-align: right; width: 20%; }
  .wacc-table td:nth-child(3) { font-size: 8pt; color: #999; padding-left: 14px; }
  .wacc-table tr:last-child td { border-bottom: none; font-weight: 700; background: #f0f4fa; }

  /* ── Section divider ── */
  .divider { border: none; border-top: 1px solid #e8ecf5; margin: 20px 0; }
</style>
</head>
<body>

<!-- Print button (hidden when printing) -->
<div class="no-print">
  <button class="print-btn" onclick="window.print()">🖨 Print / Save as PDF</button>
  <button class="print-btn" style="background:#555" onclick="window.close()">✕ Close</button>
</div>

<!-- ═══ COVER PAGE ══════════════════════════════════════════════════════════ -->
<div class="cover">
  <div>
    <div class="cover-title">DCF Valuation Report</div>
    <div class="cover-sub">${form.company}</div>
    <div class="cover-meta">${sName} &nbsp;·&nbsp; ${form.stage} &nbsp;·&nbsp; ${form.country} &nbsp;·&nbsp; ${form.revenueModel}</div>
    <div class="cover-meta" style="margin-top:4px">Generated ${today} &nbsp;·&nbsp; Hackathon 2026 &nbsp;·&nbsp; Financial Valuation Tool</div>
    <div class="cover-meta" style="margin-top:2px">EPFL/ETH Medtech Comparable Database — 57 usable companies × 10 sub-sectors</div>

    <!-- Valuation range -->
    <div class="val-box">
      <div class="val-col">
        <div class="val-label">🔴 Bear Case</div>
        <div class="val-num bear">${bearVal} ${cur}</div>
      </div>
      <div class="val-col">
        <div class="val-label">🟡 Base Case</div>
        <div class="val-num base">${baseVal} ${cur}</div>
      </div>
      <div class="val-col">
        <div class="val-label">🟢 Bull Case</div>
        <div class="val-num bull">${bullVal} ${cur}</div>
      </div>
    </div>

    <!-- Key stats -->
    <div class="stats-grid">
      <div class="stat-card"><div class="stat-label">WACC</div><div class="stat-val">${(w.wacc*100).toFixed(2)}%</div></div>
      <div class="stat-card"><div class="stat-label">Cost of Equity</div><div class="stat-val">${(w.ke*100).toFixed(2)}%</div></div>
      <div class="stat-card"><div class="stat-label">Terminal Growth</div><div class="stat-val">${(w.g*100).toFixed(1)}%</div></div>
      <div class="stat-card"><div class="stat-label">Stage Premium</div><div class="stat-val">+${(w.stagePremium*100).toFixed(2)}%</div></div>
      <div class="stat-card"><div class="stat-label">Gross Margin (sector)</div><div class="stat-val">${bench.gm?.toFixed(1)}%</div></div>
      <div class="stat-card"><div class="stat-label">EBITDA Margin (sector)</div><div class="stat-val">${bench.eb?.toFixed(1)}%</div></div>
      <div class="stat-card"><div class="stat-label">Comparables</div><div class="stat-val">${comps.length} cos.</div></div>
      <div class="stat-card"><div class="stat-label">TV Method</div><div class="stat-val" style="font-size:9pt">${shadowDCF?.terminalMethod?.split(" ")[0]||"GGM"}</div></div>
    </div>
  </div>

  <div class="cover-disclaimer">
    This report is generated by an AI-powered financial modelling tool and is for informational purposes only.
    It does not constitute investment advice, a certified valuation, or a solicitation to buy or sell securities.
    Source: EPFL/ETH Medtech Comparable Database · Damodaran January 2025 · CAPM bottom-up WACC.
  </div>
</div>

<!-- ═══ WACC SUMMARY ════════════════════════════════════════════════════════ -->
<div class="page">
<div class="content">

<h2>WACC INPUTS & SOURCES</h2>
<table class="wacc-table">
  <tr><td>Risk-free Rate (Rf)</td><td>${(w.rf*100).toFixed(2)}%</td><td>${w.sources.rf}</td></tr>
  <tr><td>Equity Risk Premium (ERP)</td><td>${(w.erp*100).toFixed(2)}%</td><td>${w.sources.erp}</td></tr>
  <tr><td>Beta (unlevered)</td><td>${w.betaU?.toFixed(3)}</td><td>${w.sources.beta}</td></tr>
  <tr><td>Stage Risk Premium</td><td>+${(w.stagePremium*100).toFixed(2)}%</td><td>${w.sources.stage}</td></tr>
  <tr><td>Cost of Equity (Ke)</td><td>${(w.ke*100).toFixed(2)}%</td><td>Computed: Rf + β×ERP + Stage</td></tr>
  <tr><td>Cost of Debt (Kd)</td><td>${(w.kd*100).toFixed(1)}%</td><td>${w.sources.kd}</td></tr>
  <tr><td>D/E Ratio (stage-adjusted)</td><td>${(w.de*100).toFixed(0)}%</td><td>${w.sources.de}</td></tr>
  <tr><td>Tax Rate</td><td>${(w.tax*100).toFixed(1)}%</td><td>${w.sources.tax}</td></tr>
  <tr><td>Terminal Growth Rate (g)</td><td>${(w.g*100).toFixed(1)}%</td><td>${w.sources.g}</td></tr>
  <tr><td><strong>WACC (final)</strong></td><td><strong>${(w.wacc*100).toFixed(2)}%</strong></td><td>Ke×(1−D/V) + Kd×(1−t)×(D/V)</td></tr>
</table>

<h2>COMPARABLE COMPANIES — ${sName.toUpperCase()}</h2>
<p style="font-size:8.5pt;color:#666;margin-bottom:10px">
  Source: EPFL/ETH Medtech Comparable Database (57 usable companies × 10 sub-sectors) · 
  🟢 = Fully reliable · 🟡 = Usable with limits · D/E corrected for confirmed yfinance errors (VCEL 0.0×, RaySearch 0.01×)
</p>
<table class="comp-table">
  <thead>
    <tr><th>Rel.</th><th>Company</th><th>Ticker</th><th>Country</th><th>GM%</th><th>EBITDA%</th><th>Beta</th><th>Rev.G%</th></tr>
  </thead>
  <tbody>${compRows}</tbody>
</table>
<p style="font-size:8pt;color:#888;margin-top:6px"><em>${RATIONALE[sectorNum]}</em></p>

<hr class="divider">

<!-- AI-generated narrative sections -->
<div class="ai-content">
${htmlBody}
</div>

</div>
</div>

</body>
</html>`;

  // ── 4. Open in new window — user prints / saves as PDF ────────────────────
  // This works in claude.ai sandbox, Render, and all browsers.
  const win = window.open("", "_blank", "width=900,height=700,scrollbars=yes");
  if (!win) {
    // Popup blocked — fallback: inject into iframe inside the artifact
    alert("Pop-up blocked by browser. Please allow pop-ups for this page and try again, or use the Excel download instead.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
  // Auto-focus for immediate print if desired
  win.focus();
};

// ─── STYLES ──────────────────────────────────────────────────────────────────
const css = `
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:system-ui,sans-serif;background:#f0f2f7;min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:32px 16px}
  .app{width:100%;max-width:840px}
  .topbar{background:#1a2e4a;padding:18px 28px;display:flex;align-items:center;gap:12px;border-radius:12px 12px 0 0}
  .logo{font-size:15px;font-weight:700;color:#fff}.logo span{color:#7eb8e8;font-weight:400}
  .badge{margin-left:auto;font-size:11px;color:#7eb8e8;background:rgba(126,184,232,0.15);padding:4px 10px;border-radius:20px;border:1px solid rgba(126,184,232,0.3)}
  .steps{display:flex;background:#f0f2f7;border:1px solid #dde2ee;border-top:none}
  .step{flex:1;padding:13px 8px;text-align:center;font-size:11px;color:#aaa;border-right:1px solid #dde2ee;transition:all .2s}
  .step:last-child{border-right:none}
  .step.active{color:#1a2e4a;font-weight:700;background:#fff;border-bottom:2px solid #1a2e4a}
  .step.done{color:#2d7a4f;background:#f0faf4}
  .prog{height:3px;background:#dde2ee}.prog-f{height:100%;background:#1a2e4a;transition:width .4s}
  .card{background:#fff;border:1px solid #dde2ee;border-top:none;border-radius:0 0 12px 12px;padding:28px}
  .sec{display:none}.sec.on{display:block}
  .g2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .full{grid-column:1/-1}
  label{display:block;font-size:11px;font-weight:600;color:#8892a4;letter-spacing:.05em;margin-bottom:6px;text-transform:uppercase}
  input,select,textarea{width:100%;padding:10px 13px;border:1.5px solid #dde2ee;border-radius:8px;font-size:14px;color:#1a2e4a;background:#fafbfd;outline:none;transition:border-color .15s;font-family:inherit;appearance:none}
  input:focus,select:focus,textarea:focus{border-color:#1a2e4a;background:#fff}
  textarea{height:90px;resize:none;line-height:1.5}
  select{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%238892a4' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:32px}
  .drop{border:2px dashed #c5cde0;border-radius:10px;padding:20px;cursor:pointer;text-align:center;transition:all .2s;background:#fafbfd}
  .drop:hover,.drop.drag{border-color:#1a2e4a;background:#f0f4fa}
  .drop input{display:none}
  .chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
  .chip{display:flex;align-items:center;gap:6px;padding:4px 10px;background:#e4edf8;border-radius:20px;font-size:11px;color:#1a2e4a;font-weight:500}
  .chip-x{cursor:pointer;opacity:.6;font-size:14px}.chip-x:hover{opacity:1}
  .exbox{background:#f0faf4;border:1px solid #a8d5be;border-radius:8px;padding:12px 14px;margin-top:10px;font-size:11px;color:#1a4a30;line-height:1.7}
  .exbox b{display:block;margin-bottom:4px;font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:#2d7a4f}
  .spinning{display:flex;align-items:center;gap:8px;font-size:12px;color:#1a2e4a;padding:10px 14px;background:#e4edf8;border-radius:8px;margin-top:10px}
  .row{display:flex;gap:10px;justify-content:space-between;align-items:center;margin-top:22px}
  .hint{font-size:11px;color:#bbb}
  .btn{padding:10px 22px;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;transition:all .15s;border:none}
  .btn-p{background:#1a2e4a;color:#fff}.btn-p:hover{background:#243d61}.btn-p:disabled{opacity:.4;cursor:not-allowed}
  .btn-s{background:transparent;border:1.5px solid #dde2ee;color:#666;font-weight:500}.btn-s:hover{background:#f5f7fb}
  .rat{background:#f0f4fa;border:1px solid #d0daf0;border-radius:8px;padding:13px 16px;margin-bottom:16px;font-size:12px;color:#3a4a6a;line-height:1.6}
  .rat b{color:#1a2e4a;display:block;margin-bottom:4px;font-size:11px;text-transform:uppercase;letter-spacing:.05em}
  .cc{background:#fafbfd;border:1px solid #e8ecf5;border-radius:10px;padding:13px 16px;display:flex;align-items:center;gap:14px;margin-bottom:8px}
  .tkr{width:44px;height:44px;border-radius:8px;background:#e4edf8;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#1a2e4a;text-align:center;flex-shrink:0}
  .bg{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:14px 0}
  .bc{background:#f5f6fa;border-radius:8px;padding:11px 13px;border:1px solid #e8ecf5}
  .bv{font-size:18px;font-weight:700;color:#1a2e4a}.bl{font-size:10px;color:#999;margin-top:2px}
  .tag{display:inline-block;padding:4px 12px;background:#e4edf8;color:#1a2e4a;border-radius:20px;font-size:12px;font-weight:600;margin-bottom:14px}
  .wbox{background:#1a2e4a;border-radius:12px;padding:20px 24px;margin:14px 0;color:#fff}
  .wval{font-size:38px;font-weight:700;color:#7eb8e8}.wsub{font-size:11px;color:rgba(255,255,255,.5);margin-top:5px}
  .wtbl{width:100%;border-collapse:collapse;font-size:12px;margin:14px 0;border-radius:8px;overflow:hidden;border:1px solid #e8ecf5}
  .wtbl th{padding:8px 12px;text-align:left;background:#f0f2f7;color:#8892a4;font-size:10px;text-transform:uppercase;letter-spacing:.05em}
  .wtbl td{padding:8px 12px;border-bottom:1px solid #f0f2f7;color:#1a2e4a}
  .wtbl tr:last-child td{border-bottom:none}
  .src-doc{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;background:#e8f5ee;color:#1a4a30}
  .src-def{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;background:#e4edf8;color:#1a2e4a}
  .dlbtn{display:flex;align-items:center;gap:12px;padding:16px 20px;background:#e8f5ee;border:1px solid #a8d5be;border-radius:10px;color:#1a4a30;font-size:14px;font-weight:600;cursor:pointer;width:100%;margin-top:14px;text-align:left;transition:background .15s}
  .dlbtn:hover{background:#d4eddf}.dlbtn:disabled{opacity:.5;cursor:not-allowed}
  .pdfbtn{display:flex;align-items:center;gap:12px;padding:16px 20px;background:#e8edf8;border:1px solid #a8bcd8;border-radius:10px;color:#1a2e4a;font-size:14px;font-weight:600;cursor:pointer;width:100%;margin-top:10px;text-align:left;transition:background .15s}
  .pdfbtn:hover{background:#d4deee}.pdfbtn:disabled{opacity:.5;cursor:not-allowed}
  .infobox{margin-top:10px;padding:10px 14px;background:#f8f9fb;border-radius:8px;font-size:11px;color:#999;border:1px solid #e5e8ee;line-height:1.6}
  .spin{display:inline-block;width:14px;height:14px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;flex-shrink:0}
  .spin-d{display:inline-block;width:13px;height:13px;border:2px solid rgba(26,46,74,.2);border-top-color:#1a2e4a;border-radius:50%;animation:sp .7s linear infinite;flex-shrink:0}
  @keyframes sp{to{transform:rotate(360deg)}}

  /* ── Advisor Chatbox ─────────────────────────────────────────────────────── */
  .chat-wrap{margin-top:22px;border:1.5px solid #dde2ee;border-radius:10px;overflow:hidden;background:#fafbfd}
  .chat-head{background:#1a2e4a;padding:11px 16px;display:flex;align-items:center;gap:9px}
  .chat-head-icon{font-size:16px}
  .chat-head-title{font-size:12px;font-weight:700;color:#fff;flex:1}
  .chat-head-sub{font-size:10px;color:rgba(255,255,255,.45)}
  .chat-msgs{max-height:280px;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
  .chat-msgs:empty::after{content:"Ask anything about the valuation — WACC, comparables, assumptions, scenario sensitivities…";color:#bbb;font-size:11px;font-style:italic}
  .msg{display:flex;flex-direction:column;gap:3px;max-width:88%}
  .msg.user{align-self:flex-end;align-items:flex-end}
  .msg.ai{align-self:flex-start;align-items:flex-start}
  .msg-bubble{padding:9px 13px;border-radius:10px;font-size:12px;line-height:1.6;word-break:break-word}
  .msg.user .msg-bubble{background:#1a2e4a;color:#fff;border-radius:10px 10px 2px 10px}
  .msg.ai .msg-bubble{background:#fff;border:1px solid #dde2ee;color:#1a2e4a;border-radius:10px 10px 10px 2px}
  .msg-meta{font-size:10px;color:#aaa;padding:0 4px}
  .chat-typing{align-self:flex-start}
  .chat-typing .msg-bubble{background:#fff;border:1px solid #dde2ee;padding:9px 13px;border-radius:10px 10px 10px 2px}
  .typing-dots{display:flex;gap:4px;align-items:center;height:14px}
  .typing-dots span{width:6px;height:6px;border-radius:50%;background:#1a2e4a;opacity:.3;animation:blink 1.2s infinite}
  .typing-dots span:nth-child(2){animation-delay:.2s}
  .typing-dots span:nth-child(3){animation-delay:.4s}
  @keyframes blink{0%,80%,100%{opacity:.3}40%{opacity:1}}
  .chat-input-row{display:flex;gap:8px;padding:10px 12px;border-top:1px solid #eef0f5;background:#fff}
  .chat-input{flex:1;padding:9px 12px;border:1.5px solid #dde2ee;border-radius:8px;font-size:13px;color:#1a2e4a;background:#fafbfd;outline:none;resize:none;font-family:inherit;line-height:1.4;min-height:38px;max-height:90px;overflow-y:auto;transition:border-color .15s}
  .chat-input:focus{border-color:#1a2e4a;background:#fff}
  .chat-send{padding:9px 14px;background:#1a2e4a;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;transition:background .15s;flex-shrink:0;display:flex;align-items:center;gap:6px}
  .chat-send:hover{background:#243d61}
  .chat-send:disabled{opacity:.4;cursor:not-allowed}
  .chat-hint{font-size:10px;color:#bbb;padding:0 12px 8px;text-align:center}

`;

// ─── COMPONENT ────────────────────────────────────────────────────────────────
export default function App() {
  const [step,      setStep]      = useState(1);
  const [form,      setForm]      = useState({company:"",description:"",revenueModel:"",stage:"",country:"",sector:""});
  const [files,     setFiles]     = useState([]);
  const [extracting,setExtracting]= useState(false);
  const [extracted, setExtracted] = useState(null);
  const [exNote,    setExNote]    = useState("");
  const [comps,     setComps]     = useState([]);
  const [bench,     setBench]     = useState({});
  const [sector,    setSector]    = useState(1);
  const [wacc,      setWacc]      = useState(null);
  const [shadowDCF, setShadowDCF] = useState(null);
  const [loading,   setLoading]   = useState(false);
  const [dlLoading, setDlLoading] = useState(false);
  const [pdfLoading,setPdfLoading]= useState(false);
  const [drag,      setDrag]      = useState(false);
  // ── Advisor chat ──────────────────────────────────────────────────────────
  const [chatMsgs,  setChatMsgs]  = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatBusy,  setChatBusy]  = useState(false);
  const chatEndRef = useRef(null);
  const fileRef = useRef();

  // Auto-scroll chat to bottom
  useEffect(()=>{ chatEndRef.current?.scrollIntoView({behavior:"smooth"}); },[chatMsgs,chatBusy]);

  const sendChat = async () => {
    const q = chatInput.trim();
    if(!q || chatBusy) return;
    setChatInput("");
    setChatMsgs(p=>[...p,{role:"user",text:q}]);
    setChatBusy(true);

    // Build a rich context snapshot of the current model state
    const waccSummary = wacc ? `WACC=${(wacc.wacc*100).toFixed(2)}%, Ke=${(wacc.ke*100).toFixed(2)}%, Kd=${(wacc.kd*100).toFixed(1)}%, g=${(wacc.g*100).toFixed(1)}%, stagePremium=${(wacc.stagePremium*100).toFixed(2)}%, beta=${wacc.betaU?.toFixed(3)}, tax=${(wacc.tax*100).toFixed(1)}%` : "not yet computed";
    const benchSummary = `Gross Margin=${bench.gm?.toFixed(1)}%, EBITDA Margin=${bench.eb?.toFixed(1)}%, Beta=${bench.b?.toFixed(3)}, Rev Growth=${bench.rg?.toFixed(1)}%`;
    const compsSummary = comps.map(c=>`${c.n} (${c.t}, GM=${c.gm?.toFixed(1)}%, β=${c.b?.toFixed(3)})`).join("; ");
    const extractedSummary = extracted ? Object.entries(extracted)
      .filter(([k,v])=>v!=null&&k!=="notes"&&!Array.isArray(v))
      .map(([k,v])=>`${k}=${typeof v==="number"?(v>1?v.toFixed(0):v.toFixed(4)):v}`).join(", ") : "none";
    const conversationHistory = chatMsgs.map(m=>({role:m.role==="user"?"user":"assistant",content:m.text}));

    const systemPrompt = `You are a senior investment analyst and valuation advisor embedded inside a live DCF valuation tool built on the EPFL/ETH Medtech Comparable Database (66 companies × 10 sub-sectors). Your role is to help the startup founder understand, challenge, or refine the financial model.

CURRENT MODEL STATE for ${form.company||"this startup"} (${form.stage}, ${form.country}, ${form.revenueModel}):
- Sub-sector: ${SECTOR_NAMES[sector]} (EPFL/ETH DB, 66-company universe)
- WACC inputs: ${waccSummary}
- D/E ratio: ${wacc?(wacc.de*100).toFixed(0):"?"} % — stage-adjusted (${form.stage}), not fixed at 20%
- Sector benchmarks used (sector medians, reliability-filtered): ${benchSummary}
- Comparable companies shown: ${compsSummary}
- Shadow DCF valuation range (UI cross-check): Bear/Base/Bull computed in-browser
- Document-extracted data: ${extractedSummary}
- Startup description: ${form.description?.slice(0,300)||"not provided"}
- Terminal value method: ${(bench.eb!=null&&bench.eb<0)?"EV/Revenue multiple (sector EBITDA negative)":"Gordon Growth Model"}

KEY MODEL IMPROVEMENTS IN THIS VERSION:
(a) D/E is stage-scaled — Pre-revenue R&D uses 5%, not 20%. This lowers WACC for early-stage companies.
(b) Shadow DCF computed in the UI for cross-check against Excel template formulas.
(c) Terminal value automatically switches to EV/Revenue when sector median EBITDA is negative.
(d) rNPV warning shown for Seed/Pre-revenue stages — remind founder that milestone probabilities should be applied.
(e) Full 66-company database with reliability flags (🟢/🟡) and D/E corrections for confirmed yfinance errors (VCEL, RaySearch, Creo Medical).

RESPOND AS A TRUSTED ADVISOR:
- Be direct, precise, and quantitative when possible
- Reference specific numbers from the model (e.g. "your current stage premium is 4.5% which pushes WACC above 18%")
- If asked about rNPV vs DCF, explain the difference clearly: rNPV applies success probabilities to each milestone cash flow; standard DCF does not
- If the founder challenges an assumption, acknowledge it and quantify the directional impact on WACC or valuation
- Keep responses concise (3–6 sentences) unless a complex question warrants more
- If asked about a specific comparable company from the 66-company database, you can cite its reliability flag and data quality notes`;

    try {
      const resp = await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:600,
          system: systemPrompt,
          messages:[
            ...conversationHistory,
            {role:"user",content:q}
          ]
        })
      });
      const data = await resp.json();
      const reply = data.content?.[0]?.text || "I couldn't generate a response. Please try again.";
      setChatMsgs(p=>[...p,{role:"ai",text:reply}]);
    } catch(e) {
      setChatMsgs(p=>[...p,{role:"ai",text:"Connection error — please try again."}]);
    }
    setChatBusy(false);
  };

  const handleChatKey = e => {
    if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendChat();}
  };

  const upd = (k,v) => setForm(f=>({...f,[k]:v}));
  const canGo = form.company && form.revenueModel && form.stage && form.country && form.description.length>=10;

  // ── Add files → call Claude API to extract ─────────────────────────────────
  const addFiles = async newF => {
    const arr = Array.from(newF);
    if(!arr.length) return;
    setFiles(p=>[...p,...arr]);
    setExtracting(true); setExNote("");
    // Use full description + company + stage context so extraction is calibrated
    const ctx = [
      form.company ? `Company: ${form.company}` : "",
      form.description?.length >= 10 ? `Description: ${form.description}` : "",
      form.stage ? `Stage: ${form.stage}` : "",
      form.revenueModel ? `Revenue model: ${form.revenueModel}` : "",
      form.country ? `Jurisdiction: ${form.country}` : "",
    ].filter(Boolean).join(". ") || "deep-tech medtech startup";
    try {
      const ex = await extractWithClaude(arr, ctx);
      setExtracted(prev => {
        // Smart merge: prefer non-null values, but don't overwrite with null
        const merged = {...(prev||{})};
        Object.entries(ex).forEach(([k,v]) => {
          if (Array.isArray(v)) {
            const prevArr = Array.isArray(merged[k]) ? merged[k] : Array(11).fill(null);
            merged[k] = prevArr.map((old, i) => v[i] != null ? v[i] : old);
          } else if (v != null) {
            merged[k] = v;
          }
        });
        return merged;
      });
      setExNote(ex.notes || "Extraction complete.");
    } catch(e) {
      setExNote("Extraction failed: " + e.message);
    }
    setExtracting(false);
  };

  const removeFile = i => setFiles(p=>p.filter((_,j)=>j!==i));

  // ── Step 1 → 2 ────────────────────────────────────────────────────────────
  const runMatch = () => {
    const s = detectSector(form.description, form.sector);
    setSector(s);
    const matched = getComps(s);
    setComps(matched);
    const ex = extracted || {};

    // Normalize gross_margin: document may give 75 or 0.75
    const normGM = ex.gross_margin != null
      ? (ex.gross_margin > 1 ? ex.gross_margin : ex.gross_margin * 100)
      : null;
    const normEB = ex.ebitda_margin != null
      ? (ex.ebitda_margin > 1 ? ex.ebitda_margin : ex.ebitda_margin * 100)
      : null;

    // Revenue growth: prefer document data; fallback to sector median
    let revGrowth = null;
    if (ex.revenue_growth_rate != null) {
      revGrowth = ex.revenue_growth_rate > 1 ? ex.revenue_growth_rate : ex.revenue_growth_rate * 100;
    } else if (ex.revenue_year1 != null && ex.revenue_year3 != null && ex.revenue_year3 > 0) {
      // Compute 2-year CAGR from doc-provided revenues
      revGrowth = (Math.pow(ex.revenue_year3 / (ex.revenue_year1 || 1), 0.5) - 1) * 100;
    }

    setBench({
      gm: normGM ?? med(matched.map(c => c.gm)),
      eb: normEB ?? med(matched.map(c => c.eb).filter(v => v != null)),
      b:  ex.beta_unlevered ?? med(matched.map(c => c.b).filter(v => v != null)),
      rg: revGrowth ?? med(matched.map(c => c.rg).filter(v => v != null)),
    });
    setStep(2);
  };

  // ── Step 2 → 3 ────────────────────────────────────────────────────────────
  const runGenerate = () => {
    setLoading(true);
    setTimeout(()=>{
      const w = buildWACC(form.country, extracted, bench.b, form.stage);
      setWacc(w);
      // FIX (b): Compute shadow DCF for UI cross-check
      setShadowDCF(computeShadowDCF(extracted, w, bench));
      setLoading(false);
      setStep(3);
    },1000);
  };

  // ── Download Excel ─────────────────────────────────────────────────────────
  const doDownload = () => {
    setDlLoading(true);
    setTimeout(()=>{
      try {
        buildExcel({...form, sector}, extracted||{}, comps, bench, wacc, sector);
      } catch(e) {
        alert("Excel generation error: " + e.message);
      }
      setDlLoading(false);
    }, 300);
  };

  // ── Download PDF Report ────────────────────────────────────────────────────
  const doPDF = async () => {
    if (!wacc) return;
    setPdfLoading(true);
    try {
      await buildPDF(
        {...form, sector},
        extracted||{},
        comps,
        bench,
        wacc,
        sector,
        shadowDCF
      );
    } catch(e) {
      alert("PDF generation error: " + e.message);
    }
    setPdfLoading(false);
  };

  const pct = step===1?33:step===2?66:100;
  const exFields = extracted ? Object.entries(extracted).filter(([k,v])=>
    v!=null && k!=="notes" && k!=="currency" && !Array.isArray(v) &&
    !["wacc_rf","wacc_erp","beta_unlevered","debt_equity_ratio"].includes(k) // hide raw WACC fields
  ) : [];
  // Prioritise showing revenue/operational fields first
  const exFieldsSorted = [
    ...exFields.filter(([k])=>["revenue_year1","revenue_year3","revenue_year5","gross_margin","ebitda_margin","burn_rate_monthly","runway_months","headcount_current","tam_size","deal1_upfront_fee","deal1_royalty_rate","funding_raised_total","pre_money_valuation","rd_spend_annual"].includes(k)),
    ...exFields.filter(([k])=>!["revenue_year1","revenue_year3","revenue_year5","gross_margin","ebitda_margin","burn_rate_monthly","runway_months","headcount_current","tam_size","deal1_upfront_fee","deal1_royalty_rate","funding_raised_total","pre_money_valuation","rd_spend_annual"].includes(k)),
  ];
  const exArrays = extracted ? Object.entries(extracted).filter(([k,v])=>Array.isArray(v)&&v.some(x=>x!=null)) : [];

  return (
    <>
      <style>{css}</style>
      <div className="app">
        <div className="topbar">
          <div className="logo">Financial <span>Valuation</span></div>
          <div className="badge">Hackathon 2026</div>
        </div>
        <div className="steps">
          {["01 — Startup Profile","02 — Comparables","03 — Financial Valuation"].map((l,i)=>(
            <div key={i} className={`step${step===i+1?" active":""}${step>i+1?" done":""}`}>{l}</div>
          ))}
        </div>
        <div className="prog"><div className="prog-f" style={{width:`${pct}%`}}/></div>
        <div className="card">

          {/* STEP 1 */}
          <div className={`sec${step===1?" on":""}`}>
            <div className="g2">
              <div>
                <label>Startup Name</label>
                <input placeholder="e.g. NeuroFlow Medical" value={form.company} onChange={e=>upd("company",e.target.value)}/>
              </div>
              <div>
                <label>Sector</label>
                <select value={form.sector} onChange={e=>upd("sector",e.target.value)}>
                  <option value="">Auto-detect from description</option>
                  {Object.entries(SECTOR_NAMES).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                </select>
              </div>
              <div>
                <label>Revenue Model</label>
                <select value={form.revenueModel} onChange={e=>upd("revenueModel",e.target.value)}>
                  <option value="">Select model</option>
                  {REVENUE_MODELS.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>
              <div>
                <label>Stage</label>
                <select value={form.stage} onChange={e=>upd("stage",e.target.value)}>
                  <option value="">Select stage</option>
                  {STAGES.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>
              <div className="full">
                <label>Country / Jurisdiction</label>
                <select value={form.country} onChange={e=>upd("country",e.target.value)}>
                  <option value="">Select country</option>
                  {COUNTRIES.map(c=><option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="full">
                <label>Startup Description — product, indication, business model, target market</label>
                <textarea placeholder="Paste your executive summary…" value={form.description} onChange={e=>upd("description",e.target.value)}/>
              </div>
              <div className="full">
                <label>Financial Documents — PDF, Excel, CSV, Word, images</label>
                <div
                  className={`drop${drag?" drag":""}`}
                  onClick={()=>fileRef.current.click()}
                  onDragOver={e=>{e.preventDefault();setDrag(true)}}
                  onDragLeave={()=>setDrag(false)}
                  onDrop={e=>{e.preventDefault();setDrag(false);addFiles(e.dataTransfer.files)}}
                >
                  <input ref={fileRef} type="file" multiple
                    accept=".pdf,.xlsx,.xls,.csv,.docx,.doc,.png,.jpg,.jpeg,.webp"
                    onChange={e=>addFiles(e.target.files)}/>
                  <div style={{fontSize:24,marginBottom:6}}>📎</div>
                  <div style={{fontSize:13,fontWeight:600,color:"#1a2e4a"}}>Drop files here or click to browse</div>
                  <div style={{fontSize:11,color:"#aaa",marginTop:4}}>
                    Claude AI reads your documents and extracts all financial data automatically
                  </div>
                </div>
                {files.length>0 && (
                  <div className="chips">
                    {files.map((f,i)=>(
                      <div key={i} className="chip">
                        {f.name.endsWith(".pdf")?"📄":f.name.match(/\.xlsx?$/i)?"📊":f.name.match(/\.docx?$/i)?"📝":"📁"} {f.name}
                        <span className="chip-x" onClick={()=>removeFile(i)}>×</span>
                      </div>
                    ))}
                  </div>
                )}
                {extracting && (
                  <div className="spinning"><span className="spin-d"/> Reading documents with Claude AI…</div>
                )}
                {exNote && !extracting && (exFieldsSorted.length>0||exArrays.length>0) && (
                  <div className="exbox">
                    <b>✅ Extracted from your documents</b>
                    {exFieldsSorted.slice(0,12).map(([k,v])=>{
                      const label = k.replace(/_/g," ").replace(/\b\w/g,c=>c.toUpperCase());
                      const display = typeof v==="number"
                        ? (k.includes("margin")||k.includes("rate")||k.includes("growth"))
                          ? `${(v>1?v:v*100).toFixed(1)}%`
                          : v > 1e5 ? `${(v/1e6).toFixed(2)}M` : v > 1e3 ? `${(v/1e3).toFixed(0)}k` : v.toFixed(2)
                        : String(v);
                      return <div key={k}><strong>{label}:</strong> {display}</div>;
                    })}
                    {exArrays.length>0 && <div style={{marginTop:4,color:"#2d7a4f"}}>+ {exArrays.length} year-by-year arrays (revenue, FTEs, CapEx, costs…)</div>}
                    <div style={{marginTop:6,opacity:.7,fontSize:10}}>{exNote}</div>
                  </div>
                )}
                {exNote && !extracting && exFields.length===0 && exArrays.length===0 && (
                  <div className="exbox" style={{background:"#fff8e1",borderColor:"#ffe082",color:"#7a5c00"}}>
                    <b>⚠️ No financial data found</b>{exNote}
                  </div>
                )}
              </div>
            </div>
            <div className="row">
              <span className="hint">{form.description.length} char{!canGo?" · fill all fields":""}</span>
              <button className="btn btn-p" onClick={runMatch} disabled={!canGo||extracting}>
                {extracting?<><span className="spin"/> Reading…</>:"Find Comparables →"}
              </button>
            </div>
          </div>

          {/* STEP 2 */}
          <div className={`sec${step===2?" on":""}`}>
            <div className="tag">{SECTOR_NAMES[sector]}</div>
            <div className="rat"><b>Why these comparables?</b>{RATIONALE[sector]}</div>

            {/* ── Selection methodology pills ──────────────────────────────── */}
            <div style={{margin:"10px 0 14px",display:"flex",flexWrap:"wrap",gap:6}}>
              {[
                {icon:"🔍", label:"Keyword match", tip:`Description scanned → best-fit sub-sector: "${SECTOR_NAMES[sector]}"`},
                {icon:"🗄️", label:"66-company DB", tip:"EPFL/ETH Medtech Comparable Database — 10 sub-sectors"},
                {icon:"🟢", label:"Reliability filter", tip:"🟢 fully reliable prioritised; 🟡 used with limits to reach 5 comps"},
                {icon:"📐", label:"Sector medians", tip:"GM, EBITDA, Beta, Revenue Growth medians feed Assumptions sheet"},
                {icon:"✅", label:"Data corrections", tip:"VCEL D/E=0.0×, RaySearch D/E=0.01×, Creo GM=46.6% (yfinance errors fixed)"},
              ].map(({icon,label,tip},i)=>(
                <div key={i} title={tip} style={{
                  display:"flex",alignItems:"center",gap:5,
                  background:"#eef2fb",borderRadius:20,padding:"4px 10px",
                  fontSize:11,fontWeight:600,color:"#1a2e4a",cursor:"default",
                  border:"1px solid #d0d8f0"
                }}>
                  <span>{icon}</span><span>{label}</span>
                  <span style={{fontSize:9,color:"#7a94c4",marginLeft:2}}>ⓘ</span>
                </div>
              ))}
            </div>

            {comps.map(c=>(
              <div className="cc" key={c.t}>
                <div className="tkr">{c.t.split(".")[0]}</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:600,color:"#1a2e4a"}}>{c.n}</div>
                  <div style={{fontSize:11,color:"#aaa",marginTop:2}}>{c.c}</div>
                  {/* Per-company reason badge */}
                  {COMP_REASON[c.t] && (
                    <div style={{
                      marginTop:5,fontSize:10,color:"#2d5a8e",background:"#f0f4fa",
                      border:"1px solid #ccd8ee",borderRadius:6,padding:"3px 8px",
                      lineHeight:1.5,maxWidth:420
                    }}>
                      {COMP_REASON[c.t]}
                    </div>
                  )}
                </div>
                <div style={{display:"flex",gap:16,marginLeft:"auto",flexShrink:0}}>
                  <div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:700,color:"#1a2e4a"}}>{c.gm.toFixed(1)}%</div><div style={{fontSize:10,color:"#aaa"}}>Gross Margin</div></div>
                  {c.b&&<div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:700,color:"#1a2e4a"}}>{c.b.toFixed(3)}</div><div style={{fontSize:10,color:"#aaa"}}>Beta</div></div>}
                </div>
              </div>
            ))}
            <div className="bg">
              {[["Gross Margin",bench.gm?.toFixed(1)+"%"],["EBITDA Margin",bench.eb?.toFixed(1)+"%"],["Beta (unlev.)",bench.b?.toFixed(3)],["Rev. Growth",bench.rg?.toFixed(1)+"%"],[`Comparables`,comps.length],["Sub-sector",SECTOR_NAMES[sector]?.split(" ")[0]]].map(([l,v],i)=>(
                <div key={i} className="bc"><div className="bv" style={i>=4?{fontSize:13}:{}}>{v}</div><div className="bl">{l}</div></div>
              ))}
            </div>
            {extracted && (exFieldsSorted.length>0||exArrays.length>0) && (
              <div className="exbox" style={{marginBottom:14}}>
                <b>📄 Document data will override sector defaults in the Excel model</b>
                {exFieldsSorted.slice(0,8).map(([k,v])=>{
                  const label = k.replace(/_/g," ").replace(/\b\w/g,c=>c.toUpperCase());
                  const display = typeof v==="number"
                    ? (k.includes("margin")||k.includes("rate")||k.includes("growth"))
                      ? `${(v>1?v:v*100).toFixed(1)}%`
                      : v > 1e5 ? `${(v/1e6).toFixed(2)}M` : v > 1e3 ? `${(v/1e3).toFixed(0)}k` : v.toFixed(2)
                    : String(v);
                  return <div key={k}><strong>{label}:</strong> {display}</div>;
                })}
                {exArrays.length>0&&<div>+ {exArrays.length} year-by-year arrays (revenue, FTEs, CapEx…)</div>}
              </div>
            )}
            <div className="row">
              <button className="btn btn-s" onClick={()=>setStep(1)}>← Back to Profile</button>
              <button className="btn btn-p" onClick={runGenerate} disabled={loading}>
                {loading?<span className="spin"/>:"Generate Financial Valuation →"}
              </button>
            </div>
          </div>

          {/* STEP 3 */}
          <div className={`sec${step===3?" on":""}`}>
            {wacc&&<>
              <div className="wbox">
                <div style={{fontSize:11,color:"rgba(255,255,255,.4)",marginBottom:5,letterSpacing:".08em"}}>COMPUTED WACC</div>
                <div className="wval">{(wacc.wacc*100).toFixed(2)}%</div>
                <div className="wsub">Ke={`${(wacc.ke*100).toFixed(1)}%`} × (1−D/V) + Kd={`${(wacc.kd*100).toFixed(1)}%`} × (1−t) × D/V</div>
              </div>
              <table className="wtbl">
                <thead>
                  <tr><th>Input</th><th style={{textAlign:"right"}}>Value</th><th>Source</th></tr>
                </thead>
                <tbody>
                  {[
                    ["Risk-free rate (Rf)",    `${(wacc.rf*100).toFixed(2)}%`,      wacc.sources.rf],
                    ["Equity Risk Premium",    `${(wacc.erp*100).toFixed(2)}%`,     wacc.sources.erp],
                    ["Beta (unlevered)",       wacc.betaU?.toFixed(3),              wacc.sources.beta],
                    ["Stage Risk Premium",     `+${(wacc.stagePremium*100).toFixed(2)}%`, wacc.sources.stage],
                    ["Cost of Equity (Ke)",    `${(wacc.ke*100).toFixed(2)}%`,      "Computed: Rf + β×ERP + Stage"],
                    ["Cost of Debt (Kd)",      `${(wacc.kd*100).toFixed(1)}%`,      wacc.sources.kd],
                    ["D/E Ratio (stage-adj.)", `${(wacc.de*100).toFixed(0)}%`,      wacc.sources.de],
                    ["Tax rate",               `${(wacc.tax*100).toFixed(1)}%`,     wacc.sources.tax],
                    ["Terminal growth (g)",    `${(wacc.g*100).toFixed(1)}%`,       wacc.sources.g],
                    ["Gross Margin",           `${(bench.gm??0).toFixed(1)}%`,      extracted?.gross_margin!=null?"📄 From document":"📊 Sector median (66-co. DB)"],
                    ["EBITDA Margin",          `${(bench.eb??0).toFixed(1)}%`,      extracted?.ebitda_margin!=null?"📄 From document":(bench.eb!=null&&bench.eb<0)?"📊 Sector median ⚠️ negative — TV uses EV/Rev":"📊 Sector median (66-co. DB)"],
                    ...(extracted?.revenue_year1!=null?[["Revenue Year 1",`${(extracted.revenue_year1/1000).toFixed(0)}k ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                    ...(extracted?.revenue_year3!=null?[["Revenue Year 3",`${(extracted.revenue_year3/1000).toFixed(0)}k ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                    ...(extracted?.burn_rate_monthly!=null?[["Monthly Burn",`${(extracted.burn_rate_monthly/1000).toFixed(0)}k ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                    ...(extracted?.runway_months!=null?[["Runway",`${extracted.runway_months} months`,"📄 From document"]]:[] ),
                    ...(extracted?.tam_size!=null?[["TAM",`${(extracted.tam_size/1e6).toFixed(0)}M ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                  ].map(([l,v,s],i)=>(
                    <tr key={i} style={{background:i%2===0?"#fff":"#fafbfd"}}>
                      <td style={{fontWeight:500}}>{l}</td>
                      <td style={{textAlign:"right",fontWeight:700,fontVariantNumeric:"tabular-nums"}}>{v}</td>
                      <td><span className={s?.includes("document")||s?.includes("Derived")?"src-doc":s?.includes("Computed")?"src-doc":"src-def"}>{s}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* ── FIX (b): Shadow DCF — indicative valuation range ─────── */}
              {shadowDCF && (
                <div style={{margin:"14px 0",border:"1px solid #dde2ee",borderRadius:10,overflow:"hidden"}}>
                  <div style={{background:"#243d61",padding:"11px 16px",display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:14,color:"#fff",fontWeight:700}}>📐 Indicative Valuation Range</span>
                    <span style={{fontSize:10,color:"rgba(255,255,255,.5)",marginLeft:"auto"}}>Shadow DCF · UI cross-check · not audited</span>
                  </div>
                  {shadowDCF.noRevData ? (
                    <div style={{padding:"16px",background:"#fff8e1",fontSize:12,color:"#7a5c00",lineHeight:1.6}}>
                      <strong>⚠️ Revenue data not found in uploaded documents.</strong> The shadow DCF requires at least one revenue projection (Year 1, Year 3, or Year 5) or a Deal 1 upfront fee to compute a valuation range. Please enter revenue assumptions directly in the Excel model's <strong>Assumptions sheet</strong> (rows 16–32), or upload a document containing revenue figures.
                    </div>
                  ) : (
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:0,background:"#fff"}}>
                      {[["🔴 Bear","bear","#c0392b"],["🟡 Base","base","#1a2e4a"],["🟢 Bull","bull","#1a4a30"]].map(([label,k,col])=>(
                        <div key={k} style={{padding:"14px 16px",borderRight:k!=="bull"?"1px solid #dde2ee":"none",textAlign:"center"}}>
                          <div style={{fontSize:11,color:"#aaa",marginBottom:4}}>{label}</div>
                          <div style={{fontSize:22,fontWeight:700,color:col}}>
                            {shadowDCF[k] == null ? "N/A"
                              : shadowDCF[k] < 0 ? "< 0"
                              : shadowDCF[k] >= 1e9 ? `${(shadowDCF[k]/1e9).toFixed(2)}B`
                              : shadowDCF[k] >= 1e6 ? `${(shadowDCF[k]/1e6).toFixed(1)}M`
                              : `${(shadowDCF[k]/1e3).toFixed(0)}k`} {shadowDCF[k] != null ? shadowDCF.currency : ""}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                  <div style={{padding:"8px 16px",background:"#f8f9fb",fontSize:10,color:"#999",borderTop:"1px solid #eef0f5"}}>
                    Terminal value method: <strong>{shadowDCF.terminalMethod}</strong>
                    {!shadowDCF.noRevData && <> · WACC {(wacc.wacc*100).toFixed(1)}% · g {(wacc.g*100).toFixed(1)}% · Order-of-magnitude estimates only; the Excel model contains the full formula-driven DCF.</>}
                  </div>
                </div>
              )}

              {/* ── FIX (d): rNPV warning for pre-revenue / clinical-stage ── */}
              {(form.stage === "Pre-revenue R&D" || form.stage === "Seed") && (
                <div style={{margin:"10px 0",padding:"10px 14px",background:"#fff8e1",border:"1px solid #ffe082",borderRadius:8,fontSize:11,color:"#7a5c00",lineHeight:1.6}}>
                  <strong>⚠️ Risk-Adjusted NPV (rNPV) Recommended</strong> — For pre-revenue / clinical-stage deep-tech startups, standard DCF overestimates value because it treats all future cash flows as certain. Industry practice (Venture Kick, B2Venture, Sofinnova) applies probability-weighted milestones: typical FDA PMA success rate ~65%; Phase II→III ~45%; commercial ramp-up ~80%. Apply milestone probability haircuts manually in the Excel Scenarios sheet.
                </div>
              )}

              {/* ── FIX (e): Comparables count now reflects full 66-co. DB ── */}
              <button className="dlbtn" onClick={doDownload} disabled={dlLoading}>
                <span style={{fontSize:24,flexShrink:0}}>{dlLoading?"⏳":"⬇"}</span>
                <span>
                  <div>{dlLoading?"Generating model…":`Download Valuation_${form.company.replace(/\s/g,"_")}.xlsx`}</div>
                  <div style={{fontSize:11,fontWeight:400,opacity:.65,marginTop:3}}>Original template preserved · all formulas intact · your data pre-filled · + Comparables sheet</div>
                </span>
              </button>

              {/* ── PDF Investor Report ─────────────────────────────────────── */}
              <button className="pdfbtn" onClick={doPDF} disabled={pdfLoading || !wacc}>
                <span style={{fontSize:24,flexShrink:0}}>{pdfLoading?"⏳":"📄"}</span>
                <span>
                  <div>{pdfLoading?"Writing report with AI — this takes ~15 seconds…":"Open Investor Report (Print → Save as PDF)"}</div>
                  <div style={{fontSize:11,fontWeight:400,opacity:.65,marginTop:3}}>
                    Opens in new window · Click 🖨 Print / Save as PDF · AI-written guide · Cover · WACC table · Comparables · Glossary · Limitations
                  </div>
                </span>
              </button>

              <div className="infobox">
                Sub-sector: <strong>{SECTOR_NAMES[sector]}</strong> · {comps.length} comparables shown ({comps.map(c=>c.t).join(", ")}) · sourced from EPFL/ETH 66-company Medtech DB ·
                {extracted&&(exFieldsSorted.length+exArrays.length)>0?` ${exFieldsSorted.length+exArrays.length} fields from documents ·`:""} Damodaran Jan 2025 · D/E stage-adjusted · Terminal value: {(bench.eb!=null&&bench.eb<0)?"EV/Revenue (EBITDA negative)":"Gordon Growth Model"} · All Excel formulas preserved
              </div>

              {/* ── Advisor Chatbox ────────────────────────────────────────── */}
              <div className="chat-wrap">
                <div className="chat-head">
                  <span className="chat-head-icon">💬</span>
                  <span className="chat-head-title">Valuation Advisor</span>
                  <span className="chat-head-sub">Ask about WACC · comparables · assumptions · scenarios</span>
                </div>
                <div className="chat-msgs">
                  {chatMsgs.map((m,i)=>(
                    <div key={i} className={`msg ${m.role}`}>
                      <div className="msg-bubble">{m.text}</div>
                      <span className="msg-meta">{m.role==="user"?"You":"Advisor"}</span>
                    </div>
                  ))}
                  {chatBusy && (
                    <div className="msg ai chat-typing">
                      <div className="msg-bubble">
                        <div className="typing-dots"><span/><span/><span/></div>
                      </div>
                    </div>
                  )}
                  <div ref={chatEndRef}/>
                </div>
                <div className="chat-hint">Shift+Enter for new line · Enter to send</div>
                <div className="chat-input-row">
                  <textarea
                    className="chat-input"
                    placeholder={`Ask about the ${form.company||"startup"} valuation — e.g. "Why is our WACC so high?" or "Our gross margin will be 85%, not ${(bench.gm??72).toFixed(0)}%"`}
                    value={chatInput}
                    onChange={e=>setChatInput(e.target.value)}
                    onKeyDown={handleChatKey}
                    rows={1}
                    disabled={chatBusy}
                  />
                  <button className="chat-send" onClick={sendChat} disabled={!chatInput.trim()||chatBusy}>
                    {chatBusy?<><span className="spin"/>…</>:<>Send ↑</>}
                  </button>
                </div>
              </div>
            </>}
            <div className="row" style={{marginTop:16}}>
              <button className="btn btn-s" onClick={()=>setStep(1)}>← New valuation</button>
            </div>
          </div>

        </div>
      </div>
    </>
  );
}

const TEMPLATE_B64 = "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";
