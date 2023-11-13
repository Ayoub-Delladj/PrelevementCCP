const realFileBtn = document.getElementById("real-file-excel");
const excelFileBtn = document.getElementById("import-excel");
const choosedFile = document.getElementById("imported-excel");
const commenrVerificationButton = document.getElementById("test");
const downloadButton = document.getElementById("dowload-button");

// *********************************************Button excel

var file_path_excl = ""

async function getExclName() {
  try {
    file_path_excl = await eel.select_and_send_excel_path()();

    if (file_path_excl) {
      const pathSeparator = file_path_excl.includes("\\") ? "\\" : "/";
      const parts = file_path_excl.split(pathSeparator);
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
  const NomFichier = await getExclName();
  if (NomFichier != "") {
    choosedFile.removeAttribute("hidden");
    excelFileBtn.setAttribute("hidden", "hidden");
    choosedFile.querySelector("span").innerHTML = NomFichier;
    if (imageFileBtn.hasAttribute("hidden")){
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
// *****************************************************Button images
const realImgBtn = document.getElementById("real-file-image");
const imageFileBtn = document.getElementById("import-image");
const choosedImage = document.getElementById("imported-image");

var file_path_img = ""

async function getImgName() {
  try {
    file_path_img = await eel.select_and_send_pdf_path()();

    if (file_path_img) {
      const pathSeparator = file_path_img.includes("\\") ? "\\" : "/";
      const parts = file_path_img.split(pathSeparator);
      const fileName = parts[parts.length - 1];
      return fileName;
    }

  } catch (error) {
    console.error("Erreur : " + error);
  }
  return "";
}


// fond transparent pourle popup
// Pour afficher le fond semi-transparent
function showLoadingPopup() {
  document.querySelector('.overlay').style.display = 'block';
}

// Pour masquer le fond semi-transparent
function hideLoadingPopup() {
  document.querySelector('.overlay').style.display = 'none';
}

// Écoute le clic sur excelFileBtn et exécute la fonction getFilePath
imageFileBtn.addEventListener("click", async function() {
  const NomFichier = await getImgName();
  if (NomFichier != "") {
    choosedImage.removeAttribute("hidden");
    imageFileBtn.setAttribute("hidden", "hidden");
    choosedImage.querySelector("span").innerHTML = NomFichier;
    if (excelFileBtn.hasAttribute("hidden")){
      commenrVerificationButton.style.backgroundColor = "#009dcc";
      commenrVerificationButton.style.color = "#ffff";
      commenrVerificationButton.disabled = false;
      }
    
  }
  else {
    imageFileBtn.innerHTML = '<i class="fa-regular fa-images image-icon"></i><br>Importer les images';
  }
});

choosedImage.addEventListener("click", function(){
    imageFileBtn.removeAttribute("hidden");
    choosedImage.setAttribute("hidden", "hidden");
    commenrVerificationButton.style.backgroundColor = "#f5f6f8";
    commenrVerificationButton.style.color = "#adb3c4";
    commenrVerificationButton.disabled = true;
});
// *********************************************Button verification

commenrVerificationButton.addEventListener('click', function() {
  const loadingPopup = document.getElementById("loading-popup");
  loadingPopup.style.display = "block";
  showLoadingPopup();

  eel.verification_CCP(file_path_excl, file_path_img);

});

//*********************************affichage du résultat */
eel.expose(close_loading_popup); // Expose la fonction pour être appelée depuis Python

function close_loading_popup(df) {
  if (df) {
    console.log(df);
    const loadingPopup = document.getElementById("loading-popup");
    loadingPopup.style.display = "none";
    hideLoadingPopup();
    document.querySelector('#dataframe-row table').innerHTML = df;
    document.getElementById("dataframe-row").removeAttribute("hidden");
    downloadButton.style.backgroundColor = "#009dcc";
    downloadButton.style.color = "#ffff";
    downloadButton.innerHTML = '<i class="fa-solid fa-download"></i> Télécharger le fichier segmenté';
    downloadButton.disabled = false;
  }
  else {
    console.log(df);
    document.getElementById("dataframe").innerHTML = "<p>Pas d'erreurs dans la vérification des comptes CCP</p>";
  }
}

//gestion d'exception
eel.expose(gestion_exception);

function gestion_exception () {
  const loadingPopup = document.getElementById("loading-popup");
  loadingPopup.style.display = "none";
  hideLoadingPopup();
  errorPopup.style.display = "block";
}

//button telecharger 
const downloadPopup = document.getElementById("download-popup");
const closePopupButton = document.getElementById("close-popup-button");
const closePopupButton2 = document.getElementById("close-popup-button2");
const errorPopup = document.getElementById("error-popup");

downloadButton.addEventListener("click", async function() {
  var folder_path = await eel.select_and_send_folder_path()();
  // const downloaded_path = await eel.telecharger_fichier(nom_du_fichier, telechargement_path, folder_path)();
  downloadPopup.style.display = "block";
  // document.getElementById("chemin-a-remplir").innerHTML = downloaded_path;
  downloadButton.style.backgroundColor = "#75bb41";
  downloadButton.style.color = "#ffff";
  downloadButton.innerHTML = '<i class="fa-solid fa-check"></i>  Téléchargement effectué';
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


// test
window.addEventListener("beforeunload", function(event) {
  if (event.target.activeElement.tagName !== 'A') {
    eel.close_python(); // Cette fonction doit être exposée depuis Python
  }
});