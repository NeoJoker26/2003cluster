// Get the modal
var modalAbout = document.getElementById("about_modal");

// Get the button that opens the modal
var aboutBtn = document.getElementById("aboutBtn");

// Get the <span> element that closes the modal
var aboutclose = document.getElementsByClassName("about_close")[0];

// When the user clicks the button, open the modal 
aboutBtn.onclick = function () {
    modalAbout.style.display = "block";
}

// When the user clicks on <span> (x), close the modal
aboutclose.onclick = function () {
    modalAbout.style.display = "none";
}

// When the user clicks anywhere outside of the modal, close it
window.onclick = function (event) {
    if (event.target == modalAbout) {
        modalAbout.style.display = "none";
    }
}

// Get the modal
var modalContact = document.getElementById("contact_modal");

// Get the button that opens the modal
var contactBtn = document.getElementById("contactBtn");

// Get the <span> element that closes the modal
var contactclose = document.getElementsByClassName("contact_close")[0];

// When the user clicks the button, open the modal 
contactBtn.onclick = function () {
    modalContact.style.display = "block";
}

// When the user clicks on <span> (x), close the modal
contactclose.onclick = function () {
    modalContact.style.display = "none";
}

// When the user clicks anywhere outside of the modal, close it
window.onclick = function (event) {
    if (event.target == modalContact) {
        modalContact.style.display = "none";
    }
}
