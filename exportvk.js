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
                "2.03",
                "Zijn de gemaakte overwegingen van het ontwerp en de implementatie vastgelegd? Is de output reproduceerbaar? Is de sensitiviteit van de output getest op veranderingen in de input en de aannames en zijn de uitkomsten hiervan vergeleken met vooraf opgestelde acceptatiecriteria?",
                getValue("question-2.03-interpretation"),
                getValue("question-2.03-citation"),
                getValue("question-2.03-assessment"),
                getValue("question-2.03-motivation"),
                getValue("question-2.03-sources")
            ],
            [
                "2.06",
                "Welke controles zijn toegepast om de aansluiting te maken tussen de invoer en de uitvoer om zo de juistheid en volledigheid van de verwerking te garanderen?",
                getValue("question-2.06-interpretation"),
                getValue("question-2.06-citation"),
                getValue("question-2.06-assessment"),
                getValue("question-2.06-motivation"),
                getValue("question-2.06-sources")
            ],
            [
                "2.08",
                "Wordt de output van het model gemonitord? Wordt dit op regelmatige basis geëvalueerd?",
                getValue("question-2.08-interpretation"),
                getValue("question-2.08-citation"),
                getValue("question-2.08-assessment"),
                getValue("question-2.08-motivation"),
                getValue("question-2.08-sources")
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
                "2.20",
                "Is er een expliciete afweging gemaakt met betrekking tot proportionaliteit en subsidiariteit en de toepassing van dataminimalisatie hiervoor?",
                getValue("question-2.20-interpretation"),
                getValue("question-2.20-citation"),
                getValue("question-2.20-assessment"),
                getValue("question-2.20-motivation"),
                getValue("question-2.20-sources")
            ],
            [
                "4.09",
                "Is er sprake van security by design?",
                getValue("question-4.09-interpretation"),
                getValue("question-4.09-citation"),
                getValue("question-4.09-assessment"),
                getValue("question-4.09-motivation"),
                getValue("question-4.09-sources")
            ]
        ];
        let csv = '';
        rows.forEach(row => {
            csv += row.map(field => `"${field.replace(/\r?\n|\r/g, ' ')}"`).join(';') + '\n';
        });

        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'veiligheid-kwaliteit-export.csv';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function exportToWord() {
        function q(nr, vraag) {
            return `
                <h3>Vraag ${nr}</h3>
                <b>Onderzoeksvraag:</b> ${vraag}<br>
                <b>Interpretatie:</b> ${getValue("question-" + nr + "-interpretation")}<br>
                <b>Toelichting (citaten):</b> ${getValue("question-" + nr + "-citation")}<br>
                <b>Beoordeling:</b> ${getValue("question-" + nr + "-assessment")}<br>
                <b>Toelichting op de beoordeling:</b> ${getValue("question-" + nr + "-motivation")}<br>
                <b>Bronnen:</b> ${getValue("question-" + nr + "-sources")}<br>
                <hr>
            `;
        }

        const htmlContent = `
            <h2>2.1 Veiligheid en kwaliteit</h2>
            <hr>
            ${q("2.03", "Zijn de gemaakte overwegingen van het ontwerp en de implementatie vastgelegd? Is de output reproduceerbaar? Is de sensitiviteit van de output getest op veranderingen in de input en de aannames en zijn de uitkomsten hiervan vergeleken met vooraf opgestelde acceptatiecriteria?")}
            ${q("2.06", "Welke controles zijn toegepast om de aansluiting te maken tussen de invoer en de uitvoer om zo de juistheid en volledigheid van de verwerking te garanderen?")}
            ${q("2.08", "Wordt de output van het model gemonitord? Wordt dit op regelmatige basis geëvalueerd?")}
            ${q("2.10", "Als er sprake is van geautomatiseerde besluitvorming, wordt daarbij voldaan aan de wet- en regelgeving die daarvoor geldt?")}
            ${q("2.20", "Is er een expliciete afweging gemaakt met betrekking tot proportionaliteit en subsidiariteit en de toepassing van dataminimalisatie hiervoor?")}
            ${q("4.09", "Is er sprake van security by design?")}
        `;
        const html = `
            <html>
            <head>
                <meta charset="utf-8">
                <title>Export Veiligheid en Kwaliteit</title>
            </head>
            <body style="font-family:Arial,Helvetica,sans-serif;">
                ${htmlContent}
            </body>
            </html>
        `;
        const blob = new Blob(['\ufeff', html], {type: 'application/msword'});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'veiligheid-kwaliteit-export.doc';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.querySelector('.export-button')?.addEventListener('click', exportToCSV);
    document.querySelector('.export-word-button')?.addEventListener('click', exportToWord);
});
