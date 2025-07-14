document.addEventListener('DOMContentLoaded', function() {
    function getValue(id) {
        return (document.getElementById(id)?.value || "").replace(/"/g, '""');
    }

    function exportToCSV() {
        const rows = [
            [
                "Nr.",
                "Onderzoeksvraag",
                "Interpretatie (in eigen woorden)",
                "Toelichting (citaten uit relevante rapporten)",
                "Beoordeling",
                "Toelichting op de beoordeling",
                "Bronnen"
            ],
            [
                "2.10",
                "Als er sprake is van geautomatiseerde besluitvorming, wordt daarbij voldaan aan de wet- en regelgeving die daarvoor geldt?",
                getValue("question-2.10-interpretation"),
                getValue("question-2.10-citation"),
                getValue("question-2.10-assessment"),
                getValue("question-2.10-motivation"),
                getValue("question-2.10-sources")
            ],
            [
                "3.10",
                "Is er sprake van automatische besluitvorming en zo ja: is dit toegestaan?",
                getValue("question-3.10-interpretation"),
                getValue("question-3.10-citation"),
                getValue("question-3.10-assessment"),
                getValue("question-3.10-motivation"),
                getValue("question-3.10-sources")
            ]
        ];
        let csv = '';
        rows.forEach(row => {
            csv += row.map(field => `"${field.replace(/\r?\n|\r/g, ' ')}"`).join(';') + '\n';
        });

        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'menselijke-tussenkomst-export.csv';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function exportToWord() {
        const htmlContent = `
            <h2>1.1 Menselijke tussenkomst</h2>
            <hr>
            <h3>Vraag 2.10</h3>
            <b>Onderzoeksvraag:</b> Als er sprake is van geautomatiseerde besluitvorming, wordt daarbij voldaan aan de wet- en regelgeving die daarvoor geldt?<br>
            <b>Interpretatie:</b> ${getValue("question-2.10-interpretation")}<br>
            <b>Toelichting (citaten):</b> ${getValue("question-2.10-citation")}<br>
            <b>Beoordeling:</b> ${getValue("question-2.10-assessment")}<br>
            <b>Toelichting op de beoordeling:</b> ${getValue("question-2.10-motivation")}<br>
            <b>Bronnen:</b> ${getValue("question-2.10-sources")}<br>
            <hr>
            <h3>Vraag 3.10</h3>
            <b>Onderzoeksvraag:</b> Is er sprake van automatische besluitvorming en zo ja: is dit toegestaan?<br>
            <b>Interpretatie:</b> ${getValue("question-3.10-interpretation")}<br>
            <b>Toelichting (citaten):</b> ${getValue("question-3.10-citation")}<br>
            <b>Beoordeling:</b> ${getValue("question-3.10-assessment")}<br>
            <b>Toelichting op de beoordeling:</b> ${getValue("question-3.10-motivation")}<br>
            <b>Bronnen:</b> ${getValue("question-3.10-sources")}<br>
        `;
        const html = `
            <html>
            <head>
                <meta charset="utf-8">
                <title>Export Menselijke Tussenkomst</title>
            </head>
            <body style="font-family:Arial,Helvetica,sans-serif;">
                ${htmlContent}
            </body>
            </html>
        `;
        const blob = new Blob(['\ufeff', html], {type: 'application/msword'});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'menselijke-tussenkomst-export.doc';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.querySelector('.export-button')?.addEventListener('click', exportToCSV);
    document.querySelector('.export-word-button')?.addEventListener('click', exportToWord);
});