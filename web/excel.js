// // *********************************************Button excel import
const realFileBtn = document.getElementById("real-file-excel");
const excelFileBtn = document.getElementById("import-excel");
const choosedFile = document.getElementById("imported-excel");
const commenrDivisionBouton = document.getElementById("division-excel-button");
var file_path = ""

async function getFileName() {
  try {
    file_path = await eel.select_and_send_excel_path()();

    if (file_path) {
      const pathSeparator = file_path.includes("\\") ? "\\" : "/";
      const parts = file_path.split(pathSeparator);
      const fileName = parts[parts.length - 1];
      return fileName;
    }

  } catch (error) {
    console.error("Erreur : " + error);
  }
  return "";
}

// Écoute le clic sur excelFileBtn et exécute la fonction getFilePath
excelFileBtn.addEventListener("click", async function() {
  const NomFichier = await getFileName();
  if (NomFichier != "") {
    choosedFile.removeAttribute("hidden");
    excelFileBtn.setAttribute("hidden", "hidden");
    choosedFile.querySelector("span").innerHTML = NomFichier;
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

function divisionExcel() {
  // Envoyer le nom du fichier au backend Python

  eel.division_excel(file_path);
}

// Soumettre le formulaire en utilisant Eel
commenrDivisionBouton.addEventListener("click", function() {
  console.log(file_path)
  divisionExcel();
});



