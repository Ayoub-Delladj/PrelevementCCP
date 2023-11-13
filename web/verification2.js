// fond transparent pourle popup
// Pour afficher le fond semi-transparent
function showLoadingPopup() {
  document.querySelector('.overlay').style.display = 'block';
}

// Pour masquer le fond semi-transparent
function hideLoadingPopup() {
  document.querySelector('.overlay').style.display = 'none';
}

const selectOption = document.getElementById("select-options");
const divOptionManu = document.getElementById("option-manuelle");
const divOptionAuto = document.getElementById("option-automatique");

let telechargement_path = '';
let nom_du_fichier = '';

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


// // *********************************************Button excel import
const realFileBtn = document.getElementById("real-file-excel");
const excelFileBtn = document.getElementById("import-excel");
const choosedFile = document.getElementById("imported-excel");
const commenrVerificationButton = document.getElementById("test");

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
// ----------------------------------------------------------------


// // *********************************************Button txt import

const documentFileBtn = document.getElementById("import-doc");
const choosedDocuement = document.getElementById("imported-doc");

var file_path2 = ""

async function getFileName2() {
  try {
    file_path2 = await eel.select_and_send_excel_path()();

    if (file_path2) {
      const pathSeparator = file_path2.includes("\\") ? "\\" : "/";
      const parts = file_path2.split(pathSeparator);
      const fileName = parts[parts.length - 1];
      return fileName;
    }

  } catch (error) {
    console.error("Erreur : " + error);
  }
  return "";
}

// Écoute le clic sur excelFileBtn et exécute la fonction getFilePath
documentFileBtn.addEventListener("click", async function() {
  const NomFichier2 = await getFileName2();
  if (NomFichier2 != "") {
    choosedDocuement.removeAttribute("hidden");
    documentFileBtn.setAttribute("hidden", "hidden");
    choosedDocuement.querySelector("span").innerHTML = NomFichier2;
    if (excelFileBtn.hasAttribute("hidden")){
      commenrVerificationButton.style.backgroundColor = "#009dcc";
      commenrVerificationButton.style.color = "#ffff";
      commenrVerificationButton.disabled = false;
      }
  }
  else {
    documentFileBtn.innerHTML = '<i class="fa-regular fa-file image-icon"></i></i><br>Importer un document';
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
  const loadingPopup = document.getElementById("loading-popup");
  loadingPopup.style.display = "block";
  showLoadingPopup();
  var Pagetraite1 = document.getElementById('nomPageExcel_1').value;
  var colonnetraite1 = document.getElementById('nomColonneExcel1').value;

  if (Pagetraite1.trim() === '') {
    Pagetraite1='Feuil1'
    }
  else{
    Pagetraite1=Pagetraite1
    }
  if (colonnetraite1.trim() === '') {
    colonnetraite1='compte'
    }
  else{
    colonnetraite1=colonnetraite1
    }

  var Pagetraite2 = document.getElementById('nomPageExcel_2').value;
  var colonnetraite2 = document.getElementById('nomColonneExcel2').value;

  if (Pagetraite2.trim() === '') {
    Pagetraite2='Feuil1'
    }
  else{
    Pagetraite2=Pagetraite2
    }
  if (colonnetraite2.trim() === '') {
    colonnetraite2='numero de compte'
    }
  else{
    colonnetraite2=colonnetraite2
    }

  eel.verification_compte(file_path, Pagetraite1, colonnetraite1, file_path2, Pagetraite2, colonnetraite2);

});


// *********************************************Button excel (automatique)



const excelFileBtnAuto = document.getElementById("import-excel-auto");
const choosedFileAuto = document.getElementById("imported-excel-auto");
const commenrVerificationButtonAuto = document.getElementById("validation-auto");
var file_path3 = ""

async function getFileName3() {
  try {
    file_path3 = await eel.select_and_send_excel_path()();

    if (file_path3) {
      const pathSeparator = file_path3.includes("\\") ? "\\" : "/";
      const parts = file_path3.split(pathSeparator);
      const fileName = parts[parts.length - 1];
      return fileName;
    }

  } catch (error) {
    console.error("Erreur : " + error);
  }
  return "";
}

// Écoute le clic sur excelFileBtn et exécute la fonction getFilePath
excelFileBtnAuto.addEventListener("click", async function() {
  const NomFichier3 = await getFileName3();
  if (NomFichier3 != "") {
    choosedFileAuto.removeAttribute("hidden");
    excelFileBtnAuto.setAttribute("hidden", "hidden");
    choosedFileAuto.querySelector("span").innerHTML = NomFichier3;
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


const telechargerFichierBouton = document.getElementById("telecharger-manuelle")

eel.expose(close_loading_popup); // Expose la fonction pour être appelée depuis Python

function close_loading_popup(chemin_telechargement, nom_fichier) {
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
  }
});