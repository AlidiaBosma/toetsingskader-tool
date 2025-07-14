document.addEventListener('DOMContentLoaded', () => {

    /* helper haalt veld-waarde op en ontsnapt aan dubbele aanhalingstekens */
    const val = id => (document.getElementById(id)?.value || '')
                        .replace(/"/g, '""');

    /* >>> CSV-export <<< */
    function exportToCSV() {
        const rows = [
            ["Nr.","Onderzoeksvraag","Interpretatie","Citaten","Beoordeling",
             "Toelichting beoordeling","Bronnen"],
            /* 2.03 */
            ["2.03",
             "Zijn ontwerp-/implementatiekeuzes vastgelegd, is output reproduceerbaar en is sensitiviteit getest?",
             val("question-2.03-interpretation"),
             val("question-2.03-citation"),
             val("question-2.03-assessment"),
             val("question-2.03-motivation"),
             val("question-2.03-sources")],
            /* 2.06 */
            ["2.06",
             "Welke controles garanderen juistheid en volledigheid tussen invoer en uitvoer?",
             val("question-2.06-interpretation"),
             val("question-2.06-citation"),
             val("question-2.06-assessment"),
             val("question-2.06-motivation"),
             val("question-2.06-sources")],
            /* 2.08 */
            ["2.08",
             "Wordt de model-output gemonitord en periodiek geëvalueerd?",
             val("question-2.08-interpretation"),
             val("question-2.08-citation"),
             val("question-2.08-assessment"),
             val("question-2.08-motivation"),
             val("question-2.08-sources")],
            /* 2.10 */
            ["2.10",
             "Wordt bij geautomatiseerde besluitvorming voldaan aan de relevante wet- en regelgeving?",
             val("question-2.10-interpretation"),
             val("question-2.10-citation"),
             val("question-2.10-assessment"),
             val("question-2.10-motivation"),
             val("question-2.10-sources")],
            /* 2.20 */
            ["2.20",
             "Is expliciet afgewogen proportionaliteit, subsidiariteit en dataminimalisatie?",
             val("question-2.20-interpretation"),
             val("question-2.20-citation"),
             val("question-2.20-assessment"),
             val("question-2.20-motivation"),
             val("question-2.20-sources")],
            /* 4.09 */
            ["4.09",
             "Is er sprake van security-by-design in de algoritme-omgeving?",
             val("question-4.09-interpretation"),
             val("question-4.09-citation"),
             val("question-4.09-assessment"),
             val("question-4.09-motivation"),
             val("question-4.09-sources")]
        ];

        /* Bouw CSV-string met ;-scheiding */
        const csv = rows.map(r => r.map(f => `"${f.replace(/\r?\n|\r/g,' ')}"`).join(';'))
                        .join('\n');

        const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
        const link = Object.assign(document.createElement('a'), {
            href: URL.createObjectURL(blob),
            download: 'veiligheid-kwaliteit-export.csv'
        });
        document.body.append(link);
        link.click();
        link.remove();
    }

    /* >>> Word-export (.doc) <<< */
    function exportToWord() {

        const q = (nr, text) => `
            <h3>Vraag ${nr}</h3>
            <b>Onderzoeksvraag:</b> ${text}<br>
            <b>Interpretatie:</b> ${val(\`question-${nr}-interpretation\`)}<br>
            <b>Citaten:</b> ${val(\`question-${nr}-citation\`)}<br>
            <b>Beoordeling:</b> ${val(\`question-${nr}-assessment\`)}<br>
            <b>Toelichting:</b> ${val(\`question-${nr}-motivation\`)}<br>
            <b>Bronnen:</b> ${val(\`question-${nr}-sources\`)}<br><hr>`;

        const htmlContent = `
            <h2>2.1 Veiligheid en kwaliteit</h2><hr>
            ${q('2.03','Ontwerp-/implementatiekeuzes, reproduceerbaarheid & sensitiviteit')}
            ${q('2.06','Controles op juistheid en volledigheid tussen invoer & uitvoer')}
            ${q('2.08','Monitoring en periodieke evaluatie van model-output')}
            ${q('2.10','Voldoen geautomatiseerde besluitvorming aan wet- en regelgeving?')}
            ${q('2.20','Afweging proportionaliteit, subsidiariteit & dataminimalisatie')}
            ${q('4.09','Security-by-design in de algoritme-omgeving')}
        `;

        const full = `
            <html><head><meta charset="utf-8"></head>
            <body style="font-family:Arial,Helvetica,sans-serif;">${htmlContent}</body></html>`;

        const blob = new Blob(['\ufeff', full], {type: 'application/msword'});
        const link = Object.assign(document.createElement('a'), {
            href: URL.createObjectURL(blob),
            download: 'veiligheid-kwaliteit-export.doc'
        });
        document.body.append(link);
        link.click();
        link.remove();
    }

    /* knoppen */
    document.querySelector('.export-button')?.addEventListener('click', exportToCSV);
    document.querySelector('.export-word-button')?.addEventListener('click', exportToWord);
});
