document.addEventListener("DOMContentLoaded", function () {
    let addWorkerBtn = document.getElementById("add-worker");
    let workersContainer = document.getElementById("workers-container");
    let workerCount = 1;
    const maxWorkers = 5;

    function addNumberValidation(input) {
        input.addEventListener("input", function () {
            this.value = this.value.replace(/\D/g, ""); // csak számok
        });
    }

    // Már létező mezőre is beállítjuk a validációt
    let firstAusweis = document.querySelector('input[name="ausweis1"]');
    if (firstAusweis) {
        addNumberValidation(firstAusweis);
    }

    addWorkerBtn.addEventListener("click", function () {
        if (workerCount < maxWorkers) {
            workerCount++;
            let workerDiv = document.createElement("div");
            workerDiv.classList.add("worker");

            workerDiv.innerHTML = `
                <label>Nachname:</label>
                <input name="nachname${workerCount}" type="text" required />
                <label>Vorname:</label>
                <input name="vorname${workerCount}" type="text" required />
                <label>Ausweis-Nr.:</label>
                <input name="ausweis${workerCount}" type="text" required />
            `;

            workersContainer.appendChild(workerDiv);

            // Új mezőre is beállítjuk a szám validációt
            let newAusweis = workerDiv.querySelector(`input[name="ausweis${workerCount}"]`);
            addNumberValidation(newAusweis);
        }
    });
});
