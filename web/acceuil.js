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


imageFileBtn.addEventListener("click", function() {
    realImgBtn.click()
});
  
realImgBtn.addEventListener("change", function(){
    if (realImgBtn.value){
        const fullPath = realImgBtn.value;
        const pathSeparator = fullPath.includes("\\") ? "\\" : "/"; // Vérifie le séparateur de chemin en fonction de la plateforme
        const parts = fullPath.split(pathSeparator); // Divise le chemin en parties
        const fileName = parts[parts.length - 1]; // Obtient le dernier élément du tableau, qui est le nom du fichier    
        choosedImage.removeAttribute("hidden");
        imageFileBtn.setAttribute("hidden", "hidden");
        choosedImage.querySelector("span").innerHTML = fileName;
        if (excelFileBtn.hasAttribute("hidden")){
            commenrVerificationButton.style.backgroundColor = "#009dcc";
            commenrVerificationButton.style.color = "#ffff";
            commenrVerificationButton.disabled = false;
            }
    }
    else {
        imageFileBtn.innerHTML = '<i class="fa-regular fa-file-excel excel-icon"></i><br>Importer fichier Excel';
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
// test

commenrVerificationButton.addEventListener('click', function() {
    eel.hello_from_python();
});