var btnSave = document.getElementById("btnSave");
btnSave.addEventListener("click", function(e) {
    var wb = XLSX.utils.book_new();
    wb.Props = {
        Title: "Sheets Tutorial",
        Author: "Hendra"
    };

    wb.SheetNames.push('Bodat');

    var ws = XLSX.utils.aoa_to_sheet([[]]);

    for (var i=0; i < ) {

    }

    wb.Sheets['Bodat'] = ws;

    ws['!ref'] = 'A1:G100';
    ws['!merges'] = [
        {s:{c:0, r:0}, e:{c:1, r:0}}
    ];
    ws['A1'] = {v: 'Hendra'};
    ws['C1'] = {v: 'Prasetya'};
    
    console.log(ws['!merges']);
    console.log(ws['!ref']);
    console.log(wb);

    XLSX.writeFile(wb, "handika.xlsx");
});
