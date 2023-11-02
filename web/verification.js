const selectOption = document.getElementById("select-options");
const divOptionManu = document.getElementById("option-manuelle");
const divOptionAuto = document.getElementById("option-automatique");

// Changement de la disposition des buttons selon le mode de vérification
selectOption.addEventListener("change", function() {
  if (selectOption.value === "Automatique") {
    divOptionAuto.removeAttribute("hidden");
    divOptionManu.setAttribute("hidden", "hidden");
  } else {
    divOptionManu.removeAttribute("hidden");
    divOptionAuto.setAttribute("hidden", "hidden");
  }
});


const realFileBtn = document.getElementById("real-file-excel");
const excelFileBtn = document.getElementById("import-excel");
const choosedFile = document.getElementById("imported-excel");
const commenrVerificationButton = document.getElementById("test");

// *********************************************Button excel

excelFileBtn.addEventListener("click", function() {
  realFileBtn.click()
});

realFileBtn.addEventListener("change", function(){
  if (realFileBtn.value){
    const fullPath = realFileBtn.value;
    const pathSeparator = fullPath.includes("\\") ? "\\" : "/"; // Vérifie le séparateur de chemin en fonction de la plateforme
    const parts = fullPath.split(pathSeparator); // Divise le chemin en parties
    const fileName = parts[parts.length - 1]; // Obtient le dernier élément du tableau, qui est le nom du fichier    
    choosedFile.removeAttribute("hidden");
    excelFileBtn.setAttribute("hidden", "hidden");
    choosedFile.querySelector("span").innerHTML = fileName;
    if (documentFileBtn.hasAttribute("hidden")){
        commenrVerificationButton.style.backgroundColor = "#009dcc";
        commenrVerificationButton.style.color = "#ffff";
        commenrVerificationButton.disabled = false;
        }
  }
  else {
    excelFileBtn.innerHTML = '<i class="fa-regular fa-file-excel excel-icon"></i><br>Importer fichier Excel';
  }
});

choosedFile.addEventListener("click", function(){
    excelFileBtn.removeAttribute("hidden");
    choosedFile.setAttribute("hidden", "hidden");
    commenrVerificationButton.style.backgroundColor = "#f5f6f8";
    commenrVerificationButton.style.color = "#adb3c4";
    commenrVerificationButton.disabled = true;
});

// *****************************************************Button document
const realDocBtn = document.getElementById("real-file-doc");
const documentFileBtn = document.getElementById("import-doc");
const choosedDocuement = document.getElementById("imported-doc");


documentFileBtn.addEventListener("click", function() {
    realDocBtn.click()
});
  
realDocBtn.addEventListener("change", function(){
    if (realDocBtn.value){
        const fullPath = realDocBtn.value;
        const pathSeparator = fullPath.includes("\\") ? "\\" : "/"; // Vérifie le séparateur de chemin en fonction de la plateforme
        const parts = fullPath.split(pathSeparator); // Divise le chemin en parties
        const fileName = parts[parts.length - 1]; // Obtient le dernier élément du tableau, qui est le nom du fichier    
        choosedDocuement.removeAttribute("hidden");
        documentFileBtn.setAttribute("hidden", "hidden");
        choosedDocuement.querySelector("span").innerHTML = fileName;
        if (excelFileBtn.hasAttribute("hidden")){
            commenrVerificationButton.style.backgroundColor = "#009dcc";
            commenrVerificationButton.style.color = "#ffff";
            commenrVerificationButton.disabled = false;
            }
    }
    else {
        documentFileBtn.innerHTML = '<i class="fa-regular fa-file-excel excel-icon"></i><br>Importer fichier Excel';
    }
});
  
choosedDocuement.addEventListener("click", function(){
    documentFileBtn.removeAttribute("hidden");
    choosedDocuement.setAttribute("hidden", "hidden");
    commenrVerificationButton.style.backgroundColor = "#f5f6f8";
    commenrVerificationButton.style.color = "#adb3c4";
    commenrVerificationButton.disabled = true;
        
});

// *********************************************Button verification
// test

commenrVerificationButton.addEventListener('click', function() {
    eel.hello_from_python();
});


// *********************************************Button excel (automatique)

const realFileBtnAuto = document.getElementById("real-file-excel-auto");
const excelFileBtnAuto = document.getElementById("import-excel-auto");
const choosedFileAuto = document.getElementById("imported-excel-auto");
const commenrVerificationButtonAuto = document.getElementById("validation-auto");

excelFileBtnAuto.addEventListener("click", function() {
  realFileBtnAuto.click()
});

realFileBtnAuto.addEventListener("change", function(){
  if (realFileBtnAuto.value){
    const fullPath = realFileBtnAuto.value;
    const pathSeparator = fullPath.includes("\\") ? "\\" : "/"; // Vérifie le séparateur de chemin en fonction de la plateforme
    const parts = fullPath.split(pathSeparator); // Divise le chemin en parties
    const fileName = parts[parts.length - 1]; // Obtient le dernier élément du tableau, qui est le nom du fichier    
    choosedFileAuto.removeAttribute("hidden");
    excelFileBtnAuto.setAttribute("hidden", "hidden");
    choosedFileAuto.querySelector("span").innerHTML = fileName;
    commenrVerificationButtonAuto.style.backgroundColor = "#009dcc";
    commenrVerificationButtonAuto.style.color = "#ffff";
    commenrVerificationButtonAuto.disabled = false;
  }
  else {
    excelFileBtnAuto.innerHTML = '<i class="fa-regular fa-file-excel excel-icon"></i><br>Importer fichier Excel';
  }
});

choosedFileAuto.addEventListener("click", function(){
    excelFileBtnAuto.removeAttribute("hidden");
    choosedFileAuto.setAttribute("hidden", "hidden");
    commenrVerificationButtonAuto.style.backgroundColor = "#f5f6f8";
    commenrVerificationButtonAuto.style.color = "#adb3c4";
    commenrVerificationButtonAuto.disabled = true;
});
