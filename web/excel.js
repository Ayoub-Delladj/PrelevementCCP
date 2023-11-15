// fond transparent pourle popup
// Pour afficher le fond semi-transparent
function showLoadingPopup() {
  document.querySelector('.overlay').style.display = 'block';
}

// Pour masquer le fond semi-transparent
function hideLoadingPopup() {
  document.querySelector('.overlay').style.display = 'none';
}

// // *********************************************Button excel import
const realFileBtn = document.getElementById("real-file-excel");
const excelFileBtn = document.getElementById("import-excel");
const choosedFile = document.getElementById("imported-excel");
const commenrDivisionBouton = document.getElementById("division-excel-button");
const telechargerFichierBouton = document.getElementById("telecharger-fichier-segmente");
var file_path = "";
let telechargement_path = '';
let nom_du_fichier = '';

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
    // telechargerFichierBouton.style.backgroundColor = "#f5f6f8";
    // telechargerFichierBouton.style.color = "#adb3c4";
    // telechargerFichierBouton.disabled = true;
});

function divisionExcel() {
  const loadingPopup = document.getElementById("loading-popup");
  loadingPopup.style.display = "block";
  showLoadingPopup();

  // Récupérer les valeurs des champs
  var nomPage = document.getElementById('nomPage').value;
  var lignesParPage = document.getElementById('lignesParPage').value;
  var lignesParPagetraite = document.getElementById('nomPageExcel_traite').value;
  if (lignesParPagetraite.trim() === '') {
    lignesParPagetraite='Feuil1';
  }
  else{
    lignesParPagetraite=lignesParPagetraite;
  }
  // Vérifier si les champs sont vides
  if (nomPage.trim() === '') {
    if (lignesParPage.trim() === '') {
      eel.division_excel(file_path,'BATCH',971,lignesParPagetraite);
    }
    else {
      eel.division_excel(file_path, 'BATCH', lignesParPage,lignesParPagetraite);
    }
  }
  else {
    if (lignesParPage.trim() === '') {
      eel.division_excel(file_path, page_name=nomPage,971,lignesParPagetraite);
    }
    else {
      eel.division_excel(file_path, nomPage, lignesParPage,971,lignesParPagetraite);
    }
  }
       
}

// Soumettre le formulaire en utilisant Eel
commenrDivisionBouton.addEventListener("click", function() {
  divisionExcel(); 
});


eel.expose(close_loading_popup); // Expose la fonction pour être appelée depuis Python

function close_loading_popup(chemin_telechargement, nom_fichier) {
  console.log('closing popup');
  const loadingPopup = document.getElementById("loading-popup");
  loadingPopup.style.display = "none";
  hideLoadingPopup();
  telechargerFichierBouton.style.backgroundColor = "#009dcc";
  telechargerFichierBouton.style.color = "#ffff";
  telechargerFichierBouton.innerHTML = '<i class="fa-solid fa-download"></i> Télécharger le fichier segmenté';
  telechargerFichierBouton.disabled = false;
  telechargement_path = chemin_telechargement;
  nom_du_fichier = nom_fichier;
}

//------------button telecharger le fichier

const downloadPopup = document.getElementById("download-popup");
const errorPopup = document.getElementById("error-popup");
const closePopupButton = document.getElementById("close-popup-button");
const closePopupButton2 = document.getElementById("close-popup-button2");

telechargerFichierBouton.addEventListener("click", async function() {
  var folder_path = await eel.select_and_send_folder_path()();
  const downloaded_path = await eel.telecharger_fichier(nom_du_fichier, telechargement_path, folder_path)();
  downloadPopup.style.display = "block";
  document.getElementById("chemin-a-remplir").innerHTML = downloaded_path;
  telechargerFichierBouton.style.backgroundColor = "#75bb41";
  telechargerFichierBouton.style.color = "#ffff";
  telechargerFichierBouton.innerHTML = '<i class="fa-solid fa-check"></i> Téléchargement effectué';
});

closePopupButton.addEventListener("click", function() {
  downloadPopup.style.display = "none";
});

closePopupButton2.addEventListener("click", function() {
  errorPopup.style.display = "none";
});

// Écoute de l'événement de touche Échap
document.addEventListener("keydown", function(event) {
  if (event.key === "Escape") {
    downloadPopup.style.display = "none";
    errorPopup.style.display = "none";
  }
});


//gestion d'exception
eel.expose(gestion_exception);

function gestion_exception () {
  const loadingPopup = document.getElementById("loading-popup");
  loadingPopup.style.display = "none";
  hideLoadingPopup();
  errorPopup.style.display = "block";
}



window.addEventListener("beforeunload", function(event) {
  if (event.target.activeElement.tagName !== 'A') {
    eel.close_python(); // Cette fonction doit être exposée depuis Python
  }});