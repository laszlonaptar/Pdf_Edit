document.getElementById("excelForm").addEventListener("submit", async function (e) {
    e.preventDefault();

    const form = e.target;
    const formData = new FormData(form);

    try {
        const response = await fetch("/generate", {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Hiba a válaszban:", errorText);
            alert("Hiba történt a fájl generálásakor.");
            return;
        }

        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = "arbeitsnachweis.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(downloadUrl);
    } catch (error) {
        console.error("Hiba a kérés során:", error);
        alert("Nem sikerült a kapcsolat a szerverrel.");
    }
});
