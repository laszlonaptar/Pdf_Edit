
document.getElementById("pdfForm").addEventListener("submit", function (event) {
    event.preventDefault();

    alert("Formular wird verarbeitet...");

    const formData = new FormData(event.target);
    const data = Object.fromEntries(formData.entries());

    fetch("https://pdf-edit.onrender.com/generate", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify(data),
    })
    .then((response) => {
        if (!response.ok) {
            throw new Error("Server returned an error");
        }
        return response.blob();
    })
    .then((blob) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "arbeitsnachweis.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        alert("Datei erfolgreich heruntergeladen.");
    })
    .catch((error) => {
        console.error("Fehler:", error);
        alert("Fehler beim Senden des Formulars.");
    });
});
