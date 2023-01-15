var modal_about = document.getElementById("aboutGoSModal");
var modal_contact = document.getElementById("contactGoSModal");
var btn_about = document.getElementById("aboutBtn");
var btn_contact = document.getElementById("contactBtn");
var about_close = document.getElementsByClassName("close-about")[0];
var contact_close = document.getElementsByClassName("close-contact")[0];


btn_about.onclick = function () { //about us open
        modal_about.style.display = "block"; 
}
btn_contact.onclick = function () { //contact us open BROKEn
    modal_contact.style.display = "block";
}


about_close.onclick = function () { //about us close BROKEN
    modal_about.style.display = "none";
    
}
contact_close.onclick = function () { // contact us close BROKEN
    modal_contact.style.display = "none";

}
window.onclick = function (event) { // add comment pls idk what this is
    if (event.target == modal_about) {
        modal_about.style.display = "none";
    }
  
}
