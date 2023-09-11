Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("open-dialog").onclick = openDialog;
    }
})

let dialog = null;


function openDialog() {
    // TODO1: Call the Office Common API that opens a dialog
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
      
        // TODO2: Add callback parameter.
        function (result) {
            dialog = result.value;
            dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
          }
      );
  }


  function processMessage(arg) {
    // document.getElementById("user-name").innerHTML = arg.message;
    changeDateFormat(arg.message);
    dialog.close();
  }

  async function changeDateFormat(dateFormat) {
    await Excel.run(async (context) => {
        const selected =  context.workbook.getSelectedRange();
        selected.load("values");
        await context.sync();
        
        const allFormats = [
            "dd-mm-yyyy",
            "dd mmmm yyyy",
            "dd-mm-yy",
            "dd.m.yy",
            "dddd mmmm yyyy"
        ]

        let formats = [
            // ["dd-mm-yyyy"]
            [allFormats[parseInt(dateFormat)]]
        ];
        selected.numberFormat = formats;
        selected.format.autofitColumns();
        selected.format.autofitRows();
        await context.sync();
        console.log(selected.values);
        
    })
}