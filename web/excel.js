const realFileBtn = document.getElementById("real-file-excel");
const excelFileBtn = document.getElementById("import-excel");
const choosedFile = document.getElementById("imported-excel");
const commenrDivisionBouton = document.getElementById("division-excel-button");

// // *********************************************Button excel import

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
    commenrDivisionBouton.style.backgroundColor = "#009dcc";
    commenrDivisionBouton.style.color = "#ffff";
    commenrDivisionBouton.disabled = false;
  }
  else {
    excelFileBtn.innerHTML = '<i class="fa-regular fa-file-excel excel-icon"></i><br>Importer fichier Excel';
  }
});

choosedFile.addEventListener("click", function(){
    excelFileBtn.removeAttribute("hidden");
    choosedFile.setAttribute("hidden", "hidden");
    commenrDivisionBouton.style.backgroundColor = "#f5f6f8";
    commenrDivisionBouton.style.color = "#adb3c4";
    commenrDivisionBouton.disabled = true;
});

// *********************************************Button division
// // test

// commenrDivisionBouton.addEventListener("click", function() {
//     eel.hello_from_python();
// });

// JavaScript (Côté client)
eel.expose(divisionExcel);

function divisionExcel() {
  // Récupérer le fichier Excel
  const excelFile = realFileBtn.value

  // Créer un objet FormData pour envoyer le fichier au backend Python
  // const formData = new FormData();
  // formData.append("excelFile", excelFile);

  // Envoyer le fichier au backend Python
  eel.division_excel(excelFile);
}

// Soumettre le formulaire en utilisant Eel
commenrDivisionBouton.addEventListener("click", (e) => {
  e.preventDefault();
  divisionExcel();
});