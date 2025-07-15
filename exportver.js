document.addEventListener('DOMContentLoaded', function () {
    function getValue(id) {
        return (document.getElementById(id)?.value || "").replace(/"/g, '""');
    }

    function exportToWord() {
        const vragen = [
            { id: "1.03", vraag: "Vindt er op vastgelegde (periodieke) momenten een afweging plaats van de kansen en risico’s over het gebruik van het algoritme?" },
            { id: "1.05", vraag: "Zijn bij uitbesteding van onderdelen of activiteiten met betrekking tot het algoritme afspraken met betrokken externe partijen gemaakt en vastgelegd?" },
            { id: "1.06", vraag: "Zijn de rollen, taken, verantwoordelijkheden, eigenaarschap en bevoegdheden in het proces beschreven (inclusief eigenaarschap) en in de praktijk toegepast / op de juiste momenten betrokken?" },
            { id: "2.02", vraag: "Is er documentatie die het ontwerp en de implementatie beschrijft?" },
            { id: "2.03", vraag: "Zijn de gemaakte overwegingen van het ontwerp en de implementatie vastgelegd? Is de output reproduceerbaar? Is de sensitiviteit van de output getest?" },
            { id: "2.11", vraag: "Is het model (code en werking) gepubliceerd en beschikbaar voor belanghebbenden?" },
            { id: "2.13", vraag: "Heeft de organisatie voldoende controle over de kwaliteit van de data bij afhankelijkheid van derden?" },
            { id: "2.15", vraag: "Is de kwaliteit gewaarborgd bij training- en testdata? Zijn deze correct gescheiden verwerkt?" },
            { id: "2.17", vraag: "Worden aan de statistische aannames van het modeltype voldaan?" }
        ];

        let htmlContent = '<h2>4.1 Verantwoording procedures</h2><hr>';

        vragen.forEach(vraag => {
            htmlContent += `
                <h3>Vraag ${vraag.id}</h3>
                <b>Onderzoeksvraag:</b> ${vraag.vraag}<br>
                <b>Interpretatie:</b> ${getValue(`question-${vraag.id}-interpretation`)}<br>
                <b>Toelichting (citaten):</b> ${getValue(`question-${vraag.id}-citation`)}<br>
                <b>Beoordeling:</b> ${getValue(`question-${vraag.id}-assessment`)}<br>
                <b>Toelichting op de beoordeling:</b> ${getValue(`question-${vraag.id}-motivation`)}<br>
                <b>Bronnen:</b> ${getValue(`question-${vraag.id}-sources`)}<br>
                <hr>
            `;
        });

        const html = `
            <html>
            <head>
                <meta charset="utf-8">
                <title>Export Verantwoording Procedures</title>
            </head>
            <body style="font-family:Arial,Helvetica,sans-serif;">
                ${htmlContent}
            </body>
            </html>
        `;

        const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'verantwoording-procedures-export.doc';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.querySelector('.export-word-button')?.addEventListener('click', exportToWord);
});
