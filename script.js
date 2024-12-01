let workbook = null;

// Charger le fichier Excel
document.getElementById('loadFileButton').addEventListener('click', () => {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
        alert('Veuillez choisir un fichier Excel');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result;
        workbook = XLSX.read(data, { type: 'binary' });

        loadSheetsData();
    };

    reader.readAsBinaryString(file);
});

// Charger les données des feuilles Excel
function loadSheetsData() {
    if (!workbook) return;

    const playersSheet = workbook.Sheets['Joueurs'];
    const coachesSheet = workbook.Sheets['Entraîneurs'];
    const clubsSheet = workbook.Sheets['Clubs'];
    const exercisesSheet = workbook.Sheets['Exercices'];

    // Charger les joueurs dans la sélection
    const playersData = XLSX.utils.sheet_to_json(playersSheet, { header: 1 });
    const coachesData = XLSX.utils.sheet_to_json(coachesSheet, { header: 1 });
    const clubsData = XLSX.utils.sheet_to_json(clubsSheet, { header: 1 });
    const exercisesData = XLSX.utils.sheet_to_json(exercisesSheet, { header: 1 });

    const playerSelect = document.getElementById('playerSelect');
    const coachSelect = document.getElementById('coachSelect');
    const exerciseSelect = document.getElementById('exerciseSelect');

    playersData.forEach(row => {
        const option = document.createElement('option');
        option.value = row[0]; // Nom du joueur
        option.textContent = row[0]; // Nom du joueur
        playerSelect.appendChild(option);
    });

    coachesData.forEach(row => {
        const option = document.createElement('option');
        option.value = row[0]; // Nom de l'entraîneur
        option.textContent = row[0]; // Nom de l'entraîneur
        coachSelect.appendChild(option);
    });

    exercisesData.forEach(row => {
        const option = document.createElement('option');
        option.value = row[0]; // Nom de l'exercice
        option.textContent = row[0]; // Nom de l'exercice
        exerciseSelect.appendChild(option);
    });
}

// Ajouter un joueur
document.getElementById('addPlayer').addEventListener('click', () => {
    const playerName = prompt("Entrez le nom du joueur");
    if (!playerName) return;

    const playersSheet = workbook.Sheets['Joueurs'];
    const playersData = XLSX.utils.sheet_to_json(playersSheet, { header: 1 });

    // Ajouter le joueur dans la feuille
    playersData.push([playerName]);
    const newSheet = XLSX.utils.aoa_to_sheet(playersData);
    workbook.Sheets['Joueurs'] = newSheet;

    loadSheetsData(); // Recharger les données
});

// Ajouter un entraîneur
document.getElementById('addCoach').addEventListener('click', () => {
    const coachName = prompt("Entrez le nom de l'entraîneur");
    if (!coachName) return;

    const coachesSheet = workbook.Sheets['Entraîneurs'];
    const coachesData = XLSX.utils.sheet_to_json(coachesSheet, { header: 1 });

    // Ajouter l'entraîneur dans la feuille
    coachesData.push([coachName]);
    const newSheet = XLSX.utils.aoa_to_sheet(coachesData);
    workbook.Sheets['Entraîneurs'] = newSheet;

    loadSheetsData(); // Recharger les données
});

// Ajouter un club
document.getElementById('addClub').addEventListener('click', () => {
    const clubName = prompt("Entrez le nom du club");
    if (!clubName) return;

    const clubsSheet = workbook.Sheets['Clubs'];
    const clubsData = XLSX.utils.sheet_to_json(clubsSheet, { header: 1 });

    // Ajouter le club dans la feuille
    clubsData.push([clubName]);
    const newSheet = XLSX.utils.aoa_to_sheet(clubsData);
    workbook.Sheets['Clubs'] = newSheet;

    loadSheetsData(); // Recharger les données
});

// Exporter le fichier Excel mis à jour
document.getElementById('exportExcel').addEventListener('click', () => {
    const clubName = document.getElementById('clubName').value;
    const coachName = document.getElementById('coachSelect').value;
    const playerName = document.getElementById('playerSelect').value;
    const exercise = document.getElementById('exerciseSelect').value;

    if (!clubName || !coachName || !playerName || !exercise) {
        alert('Tous les champs doivent être remplis');
        return;
    }

    // Exporter le fichier mis à jour
    XLSX.writeFile(workbook, 'Club_Sport.xlsx');
});
