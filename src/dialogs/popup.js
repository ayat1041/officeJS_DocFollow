(async () => {
  await Office.onReady();

  // TODO1: Assign handler to the OK button.
  document.getElementById("format-1").onclick = ()=> sendStringToParentPage("0");
  document.getElementById("format-2").onclick = ()=> sendStringToParentPage("1");
  document.getElementById("format-3").onclick = ()=> sendStringToParentPage("2");
  document.getElementById("format-4").onclick = ()=> sendStringToParentPage("3");
  document.getElementById("format-5").onclick = ()=> sendStringToParentPage("4");


  // TODO2: Create the OK button handler
  function sendStringToParentPage(typeFormat) {
    // const userName = document.getElementById("name-box").value;
    Office.context.ui.messageParent(typeFormat);
  }
})();
