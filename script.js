document.getElementById("arbeitsForm").addEventListener("submit", function(e) {
    e.preventDefault();
    alert("Formular wird verarbeitet...");

    const formData = new FormData(e.target);
    const values = Object.fromEntries(formData.entries());
    console.log("Erfasste Daten:", values);

    // Ide helyezhető az XLS-generálási logika vagy egy fetch()-hívás a backendhez
});
