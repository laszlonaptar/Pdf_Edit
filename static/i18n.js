// i18n.js — nyelvkezelés (URL ?lang=, localStorage, default=de)
(function () {
  const KEY = "pdfedit.lang";

  function parseQuery() {
    const q = {};
    const s = window.location.search.substring(1);
    if (!s) return q;
    for (const part of s.split("&")) {
      const [k, v] = part.split("=");
      if (!k) continue;
      q[decodeURIComponent(k)] = decodeURIComponent(v || "");
    }
    return q;
  }

  function getLang() {
    const qs = parseQuery();
    if (qs.lang && ["de", "hr", "en"].includes(qs.lang)) {
      localStorage.setItem(KEY, qs.lang);
      return qs.lang;
    }
    const saved = localStorage.getItem(KEY);
    if (saved && ["de", "hr", "en"].includes(saved)) return saved;
    return "de";
  }

  const L = {
    de: {
      title: "Leistungsnachweis",
      logout: "Abmelden",
      date: "Datum",
      site: "Bau / Ausführungsort",
      bf: "Bauleiter/Fachbauleiter",
      break: "Pausen (Min)",
      desc: "Beschreibung / Was wurde gemacht?",
      translate_to_de: "Auf Deutsch übersetzen",
      workers: "Mitarbeiter",
      th_last: "Name",
      th_first: "Vorname",
      th_id: "Ausweis",
      th_start: "Beginn",
      th_end: "Ende",
      th_vh: "Vorhaltung",
      btn_excel: "Excel generieren",
      btn_pdf: "PDF Vorschau",
      t_busy: "Übersetzen...",
      t_done: "Übersetzung fertig.",
      t_error: "Übersetzung fehlgeschlagen."
    },
    hr: {
      title: "Evidencija rada",
      logout: "Odjava",
      date: "Datum",
      site: "Gradilište / Mjesto izvođenja",
      bf: "Voditelj gradilišta",
      break: "Pauze (min)",
      desc: "Opis / Što je rađeno?",
      translate_to_de: "Prevedi na njemački",
      workers: "Radnici",
      th_last: "Prezime",
      th_first: "Ime",
      th_id: "Iskaznica",
      th_start: "Početak",
      th_end: "Kraj",
      th_vh: "Oprema",
      btn_excel: "Generiraj Excel",
      btn_pdf: "PDF pregled",
      t_busy: "Prevodim...",
      t_done: "Prijevod gotov.",
      t_error: "Greška pri prijevodu."
    },
    en: {
      title: "Work report",
      logout: "Logout",
      date: "Date",
      site: "Site / Location",
      bf: "Site manager",
      break: "Breaks (min)",
      desc: "Description / What was done?",
      translate_to_de: "Translate to German",
      workers: "Workers",
      th_last: "Last name",
      th_first: "First name",
      th_id: "ID",
      th_start: "Start",
      th_end: "End",
      th_vh: "Equipment",
      btn_excel: "Generate Excel",
      btn_pdf: "PDF preview",
      t_busy: "Translating...",
      t_done: "Translation done.",
      t_error: "Translation failed."
    }
  };

  function applyI18n(lang) {
    for (const el of document.querySelectorAll("[data-i18n]")) {
      const key = el.getAttribute("data-i18n");
      const val = L[lang][key];
      if (typeof val === "string") el.textContent = val;
    }
    const select = document.getElementById("lang-select");
    if (select) select.value = lang;

    // Fordítás gomb csak HR nyelvnél
    const btn = document.getElementById("btn-translate-hr-de");
    if (btn) btn.style.display = (lang === "hr" ? "" : "none");
  }

  function initLangSelect() {
    const select = document.getElementById("lang-select");
    if (!select) return;
    select.addEventListener("change", () => {
      const lang = select.value;
      localStorage.setItem(KEY, lang);
      const url = new URL(window.location.href);
      url.searchParams.set("lang", lang);
      window.location.replace(url.toString());
    });
  }

  window.__i18n__ = { getLang, applyI18n };

  document.addEventListener("DOMContentLoaded", () => {
    const lang = getLang();
    applyI18n(lang);
    initLangSelect();
  });
})();
