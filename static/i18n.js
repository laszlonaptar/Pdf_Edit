// i18n.js — DE és HR; HR nézetben fordítás gomb
(function () {
  const KEY = "pdfedit.lang";

  const L = {
    de: {
      title: "Leistungsnachweis",
      logout: "Abmelden",
      date: "Datum der Leistungsausführung",
      site: "Bau und Ausführungsort",
      bf: "Bauleiter/Fachbauleiter",
      break: "Pausen (Min)",
      half_break: "Halbstündige Pause (30 Min)",
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
      t_busy: "Übersetzen…",
      t_done: "Übersetzung fertig.",
      t_error: "Übersetzung fehlgeschlagen."
    },
    hr: {
      title: "Evidencija rada",
      logout: "Odjava",
      date: "Datum izvođenja radova",
      site: "Gradilište / Mjesto izvođenja",
      bf: "Voditelj gradilišta",
      break: "Pauze (min)",
      half_break: "Pola sata pauze (30 min)",
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
      t_busy: "Prevodim…",
      t_done: "Prijevod gotov.",
      t_error: "Greška pri prijevodu."
    }
  };

  function parseQuery() {
    const q = {};
    const s = window.location.search.substring(1);
    if (!s) return q;
    for (const part of s.split("&")) {
      const [k, v] = part.split("=");
      q[decodeURIComponent(k)] = decodeURIComponent(v || "");
    }
    return q;
  }

  function getLang() {
    const qs = parseQuery();
    if (qs.lang && ["de", "hr"].includes(qs.lang)) {
      localStorage.setItem(KEY, qs.lang);
      return qs.lang;
    }
    const saved = localStorage.getItem(KEY);
    if (saved && ["de", "hr"].includes(saved)) return saved;
    return "de";
  }

  function applyI18n(lang) {
    for (const el of document.querySelectorAll("[data-i18n]")) {
      const key = el.getAttribute("data-i18n");
      const val = L[lang][key];
      if (typeof val === "string") el.textContent = val;
    }
    const select = document.getElementById("lang-select");
    if (select) select.value = lang;

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
  window.__i18nDict = L;

  document.addEventListener("DOMContentLoaded", () => {
    const lang = getLang();
    applyI18n(lang);
    initLangSelect();
  });
})();
