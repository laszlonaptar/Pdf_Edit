function addMitarbeiter() {
    const container = document.getElementById("mitarbeiterContainer");
    const clone = container.children[0].cloneNode(true);
    container.appendChild(clone);
}

document.getElementById("arbeitsForm").addEventListener("submit", function(e) {
    e.preventDefault();
    alert("A beküldés működik – a backend még nem kapcsolódik.");
});