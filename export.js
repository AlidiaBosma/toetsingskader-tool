/* export-word.js
   Universele Word-export voor het hele Toetsingskader
   - Werkt op elke pagina met .question-card elementen (geen per-pagina scripts meer nodig).
   - Ondersteunt:
       1) Export van de huidige pagina (knop .export-word-button).
       2) Export van ALLE pagina’s in het project (knop .export-word-all-button) door index.html te crawlen.
   - Verzamelt per vraag:
       • Vraagnummer, categorie, onderzoeksvraag
       • Risico & beheersmaatregel (uit .maatregel-block)
       • Inspectiestappen (Opzet & Bestaan)
       • Gebruikersinvoer uit de formulieren:
         interpretation, citation, assessment, motivation, sources
   - Integratie:
       1) Plaats dit bestand in je project, bv. export-word.js
       2) Voeg op elke pagina <script src="export-word.js" defer></script> toe
       3) Zorg dat er een knop met class="export-word-button" staat voor export van de huidige pagina
          (optioneel ook .export-word-all-button voor alles exporteren)
*/

(function () {
  // Optionele whitelist (als je niet wilt crawlen via index.html):
  const PAGES_WHITELIST = [
    // "index.html",
    // "menselijke-tussenkomst.html",
    // "privacy-en-data.html",
    // "geen-discriminatie.html",
    // "maatschappelijke-impact.html",
    // "uitlegbaarheid.html",
    // "verantwoording.html",
    // ... vul aan of laat leeg voor autodiscovery via index.html
  ];

  // ---------- Helpers ----------
  const fmt = {
    date(d = new Date()) {
      const p = (n) => String(n).padStart(2, "0");
      return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
    },
    text(el) {
      return el ? el.textContent.trim() : "";
    },
    html(el) {
      return el ? el.innerHTML : "";
    },
    valById(doc, id) {
      const el = doc.getElementById(id);
      if (!el) return "";
      const v = el.value || "";
      return v.replace(/"/g, '""');
    },
    escHtml(str) {
      return (str || "").replace(/[&<>"']/g, (s) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[s]));
    },
    slug(s) {
      return (s || "")
        .toLowerCase()
        .replace(/[^\w\-]+/g, "-")
        .replace(/\-+/g, "-")
        .replace(/^\-+|\-+$/g, "");
    },
  };

  function collectPageData(doc, sourceUrl = "") {
    const data = {};
    data.title = fmt.text(doc.querySelector("h1")) || fmt.text(doc.querySelector("title"));
    data.breadcrumb = Array.from(doc.querySelectorAll(".breadcrumb a, .breadcrumb span"))
      .map((n) => fmt.text(n))
      .filter(Boolean)
      .join(" > ");
    data.user = fmt.text(doc.querySelector(".user-info .user")) || "";
    data.timestamp = fmt.text(doc.querySelector(".user-info .timestamp")) || "";
    data.principleDescription = fmt.text(doc.querySelector(".principle-description")) || "";
    data.sourceUrl = sourceUrl;
    data.cards = [];

    const cards = doc.querySelectorAll(".question-card");
    cards.forEach((card) => {
      const id = fmt.text(card.querySelector(".question-header .question-id"));
      const category = fmt.text(card.querySelector(".question-header .category"));
      const vraag = fmt.text(card.querySelector(".onderzoeksvraag-block"));

      const risicoHtml = fmt.html(card.querySelector(".maatregel-block"));

      const h4s = card.querySelectorAll(".inspectie-block h4");
      let opzetList = null;
      let bestaanList = null;
      if (h4s[0] && h4s[0].textContent.toLowerCase().includes("opzet")) {
        opzetList = h4s[0].nextElementSibling;
      }
      if (h4s[1] && h4s[1].textContent.toLowerCase().includes("bestaan")) {
        bestaanList = h4s[1].nextElementSibling;
      }
      const toItems = (ul) => (ul ? Array.from(ul.querySelectorAll("li")).map((li) => li.innerHTML.trim()) : []);

      // Formulierwaarden (optioneel aanwezig)
      const interpretation = fmt.valById(doc, `question-${id}-interpretation`);
      const citation = fmt.valById(doc, `question-${id}-citation`);
      const assessment = fmt.valById(doc, `question-${id}-assessment`);
      const motivation = fmt.valById(doc, `question-${id}-motivation`);
      const sources = fmt.valById(doc, `question-${id}-sources`);

      data.cards.push({
        id,
        category,
        vraag,
        risicoHtml,
        opzetItems: toItems(opzetList),
        bestaanItems: toItems(bestaanList),
        interpretation,
        citation,
        assessment,
        motivation,
        sources,
      });
    });

    return data;
  }

  function cardToHTML(card) {
    // Tabel weergave voor nette Word-output
    const liList = (items) => items.map((i) => `<li>${i}</li>`).join("");

    // Gebruikersinvoer onderaan elke vraag
    const userBlock = `
      <table border="1" cellspacing="0" cellpadding="8" style="border-collapse:collapse;width:100%;font-size:11pt;margin-top:6px;">
        <tr><th style="width:22%;text-align:left;">Interpretatie</th><td>${fmt.escHtml(card.interpretation)}</td></tr>
        <tr><th style="text-align:left;">Toelichting (citaten)</th><td>${fmt.escHtml(card.citation)}</td></tr>
        <tr><th style="text-align:left;">Beoordeling</th><td>${fmt.escHtml(card.assessment)}</td></tr>
        <tr><th style="text-align:left;">Toelichting op de beoordeling</th><td>${fmt.escHtml(card.motivation)}</td></tr>
        <tr><th style="text-align:left;">Bronnen</th><td>${fmt.escHtml(card.sources)}</td></tr>
      </table>
    `;

    return `
      <h3 style="margin:16px 0 6px 0;">Vraag ${fmt.escHtml(card.id)}${card.category ? ` — <span style="font-weight:normal">${fmt.escHtml(card.category)}</span>` : ""}</h3>
      <table border="1" cellspacing="0" cellpadding="8" style="border-collapse:collapse;width:100%;font-size:11pt;">
        <tr>
          <th style="text-align:left;width:22%;">Onderzoeksvraag</th>
          <td>${fmt.escHtml(card.vraag)}</td>
        </tr>
        <tr>
          <th style="text-align:left;">Risico &amp; Beheersmaatregel</th>
          <td>${card.risicoHtml || ""}</td>
        </tr>
        <tr>
          <th style="text-align:left;">Inspectiestappen – Opzet</th>
          <td><ul style="margin:0 0 0 18px;padding:0;">${liList(card.opzetItems)}</ul></td>
        </tr>
        <tr>
          <th style="text-align:left;">Inspectiestappen – Bestaan</th>
          <td><ul style="margin:0 0 0 18px;padding:0;">${liList(card.bestaanItems)}</ul></td>
        </tr>
      </table>
      ${userBlock}
    `;
  }

  function pageToSectionHTML(pageData) {
    const headerMeta = `
      <div style="font-size:10pt;color:#666;">
        ${pageData.breadcrumb ? `Pad: ${fmt.escHtml(pageData.breadcrumb)}<br>` : ""}
        ${pageData.user ? `Auteur: ${fmt.escHtml(pageData.user)}<br>` : ""}
        ${pageData.timestamp ? `Pagina-timestamp: ${fmt.escHtml(pageData.timestamp)}<br>` : ""}
        ${pageData.sourceUrl ? `Bron: ${fmt.escHtml(pageData.sourceUrl)}` : ""}
      </div>
    `;

    const cardsHtml = pageData.cards.map(cardToHTML).join("\n");
    return `
      <section style="page-break-inside:avoid;margin:0 0 28px 0;">
        <h2 style="margin:0 0 6px 0;">${fmt.escHtml(pageData.title || "Onbenoemde sectie")}</h2>
        ${headerMeta}
        ${pageData.principleDescription ? `<p style="margin:10px 0;">${fmt.escHtml(pageData.principleDescription)}</p>` : ""}
        ${cardsHtml || `<p style="color:#a00;">Geen vragen gevonden op deze pagina.</p>`}
      </section>
    `;
  }

  function wrapAsWordDoc(bodyHTML, docTitle = "Toetsingskader-export") {
    const css = `
      body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #111; }
      h1, h2, h3 { font-family: Calibri, Arial, sans-serif; }
      h1 { font-size: 18pt; margin: 0 0 8px 0; }
      h2 { font-size: 16pt; margin: 18px 0 6px 0; }
      h3 { font-size: 13pt; }
      table { border: 1px solid #666; }
      th { background:#f2f2f2; }
      .doc-cover { border-bottom:2px solid #000; padding-bottom:6px; margin-bottom:18px; }
    `;
    const now = fmt.date();
    return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>${fmt.escHtml(docTitle)}</title>
<style>${css}</style>
</head>
<body>
  <div class="doc-cover">
    <h1>${fmt.escHtml(docTitle)}</h1>
    <div style="font-size:10pt;color:#666;">Gegenereerd: ${now}</div>
  </div>
  ${bodyHTML}
</body>
</html>`;
  }

  function downloadWord(html, filename = "toetsingskader-export.doc") {
    const blob = new Blob(["\ufeff", html], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = Object.assign(document.createElement("a"), { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    URL.revokeObjectURL(url);
    a.remove();
  }

  async function exportCurrentPageToWord() {
    const data = collectPageData(document, location.pathname.split("/").pop() || location.href);
    const section = pageToSectionHTML(data);
    const html = wrapAsWordDoc(section, data.title ? `Toetsingskader – ${data.title}` : "Toetsingskader – Export");
    const fname = (data.title ? fmt.slug(`toetsingskader-${data.title}`) : "toetsingskader-export") + ".doc";
    downloadWord(html, fname);
  }

  async function exportAllPagesToWord() {
    const urls = await discoverAllPages();
    const sections = [];

    // Altijd ook huidige pagina meenemen (zekerheid en volgorde)
    const current = location.pathname.split("/").pop() || "index.html";
    const withCurrentFirst = [current, ...urls.filter((u) => u !== current)];

    for (const url of withCurrentFirst) {
      try {
        const abs = new URL(url, location.href).href;
        const res = await fetch(abs, { credentials: "same-origin" });
        if (!res.ok) continue;
        const text = await res.text();
        const parser = new DOMParser();
        const doc = parser.parseFromString(text, "text/html");
        if (doc.querySelector(".question-card")) {
          const data = collectPageData(doc, abs);
          sections.push(pageToSectionHTML(data));
        }
      } catch (e) {
        console.warn("Kan pagina niet laden:", url, e);
      }
    }

    const bodyHTML = sections.join('\n<hr style="page-break-after:always;border:0;border-top:1px dashed #ccc;margin:24px 0;">\n');
    const html = wrapAsWordDoc(bodyHTML, "Toetsingskader – Complete export");
    downloadWord(html, "toetsingskader-compleet.doc");
  }

  async function discoverAllPages() {
    if (PAGES_WHITELIST.length) return uniq(PAGES_WHITELIST);

    let urls = [];
    try {
      // Crawlen via index.html: alle links naar .html verzamelen
      const idxUrl = new URL("index.html", location.href).href;
      const res = await fetch(idxUrl, { credentials: "same-origin" });
      if (res.ok) {
        const text = await res.text();
        const doc = new DOMParser().parseFromString(text, "text/html");
        const anchors = Array.from(doc.querySelectorAll('a[href$=".html"]'));
        urls = anchors.map((a) => a.getAttribute("href")).filter(Boolean);
      }
    } catch (e) {
      console.warn("Index niet gevonden of niet leesbaar:", e);
    }

    // Filter en dedup
    return uniq(urls);
  }

  function uniq(arr) {
    const norm = (u) => (u || "").split("#")[0].split("?")[0];
    const out = [];
    const seen = new Set();
    for (const u of arr) {
      const n = norm(u);
      if (!n) continue;
      if (!seen.has(n)) {
        seen.add(n);
        out.push(n);
      }
    }
    return out;
  }

  // ---------- UI binding ----------
  function bindButtons() {
    const oneBtn = document.querySelector(".export-word-button");
    if (oneBtn) oneBtn.addEventListener("click", exportCurrentPageToWord);

    const allBtn = document.querySelector(".export-word-all-button");
    if (allBtn) allBtn.addEventListener("click", exportAllPagesToWord);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindButtons);
  } else {
    bindButtons();
  }
})();
