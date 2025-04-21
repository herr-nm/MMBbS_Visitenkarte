# Erforderliche Module installieren (falls nicht vorhanden)
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
if (-not (Get-Module -ListAvailable -Name QRCodeGenerator)) {
    Install-Module -Name QRCodeGenerator -Force -Scope CurrentUser
}

# Importieren der Module
Import-Module ImportExcel
Import-Module QRCodeGenerator

# Pfad zur Excel-Datei
$excelFile = "visitenkarten.xlsx"

# Lese die Daten aus der Excel-Tabelle
$data = Import-Excel -Path $excelFile

# Standardwerte
$companyLogo = "mmbbs_logo.png"
$defaultPhone = "+4951164619811"
$baseEmail = "@mmbbs.de"
$mapLat = "52.3205168"
$mapLon = "9.815254"
$mapPopup = "MMBbS"

# Fixe Links
$linkedInSchool = "https://www.linkedin.com/school/mmbbs/"
$instagramSchool = "https://www.instagram.com/multi_media_bbs/"

# Erstelle den Ausgabeordner
$outputFolder = "Visitenkarten"
if (!(Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
}

# Schleife durch jede Zeile in der Excel-Datei
foreach ($person in $data) {
    $name = $person.Name
    $title = $person.Titel
    $phone = if ($person.Telefon) { $person.Telefon } else { $defaultPhone }
    $email = "$($person.Nachname.ToLower())$baseEmail"
    $linkedin = $person.LinkedIn
    $profilePic = $person.Profilbild
    $fileName = $person.Dateiname

    # Erstelle die vCard (.vcf)
    $vcardContent = @"
BEGIN:VCARD
VERSION:3.0
FN:$name
TITLE:$title
TEL:$phone
EMAIL:$email
ORG:MMBbS
ADR:;;Expo Plaza 3;Hannover;;30539;Germany
END:VCARD
"@

    $vcardPath = "$outputFolder\$fileName.vcf"
    Set-Content -Path $vcardPath -Value $vcardContent

    # Erstelle einen QR-Code für die vCard
    $qrPath = "$outputFolder\$fileName.png"
    New-QRCode -Content "https://herr-nm.de/$fileName.vcf" -OutFile $qrPath -Size 300

    # HTML-Visitenkarte erstellen
    $htmlPath = "$outputFolder\$fileName.html"
    $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Digitale Visitenkarte - $name</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
</head>
<body>
    <div class="container card-container">
        <div class="profile-card text-center">
            <img src="$companyLogo" alt="Firmenlogo" class="company-logo" width="200">
            <img src="$profilePic" alt="Profilbild" class="profile-img">
            <h3>$name</h3>
            <p>$title</p>
        </div>

        <div class="info-card text-center">
            <a href="$fileName.vcf" download class="btn btn-light w-100 mb-2">
                <i class="fas fa-download"></i> Kontakt speichern
            </a>
        </div>

        <div class="info-card text-center">
            <button class="btn btn-light w-100" type="button" data-bs-toggle="collapse" data-bs-target="#qrCollapse">
                <i class="fas fa-qrcode"></i> QR-Code anzeigen
            </button>
            <div class="collapse mt-3" id="qrCollapse">
                <img src="$fileName.png" alt="QR-Code" class="img-fluid">
            </div>
        </div>

        <div class="info-card">
            <h5>Kontakt</h5>
            <p><i class="fas fa-phone"></i> <a href="tel:$phone">$phone</a></p>
            <p><i class="fas fa-envelope"></i> <a href="mailto:$email">$email</a></p>
"@

    if ($linkedin) { $htmlContent += "<p><i class='fab fa-linkedin'></i> <a href='$linkedin' target='_blank'>LinkedIn</a></p>`n" }

    $htmlContent += @"
        </div>

        <div class="social-card">
            <h5>Social Media</h5>
            <p><i class="fab fa-linkedin"></i> <a href="$linkedInSchool" target="_blank">MMBbS auf LinkedIn</a></p>
            <p><i class="fab fa-instagram"></i> <a href="$instagramSchool" target="_blank">MMBbS auf Instagram</a></p>
        </div>

        <div class="address-card">
            <h5>Adresse der MMBbS</h5>
            <p>Expo Plaza 3, 30539 Hannover</p>
            <div id="map"></div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://unpkg.com/leaflet/dist/leaflet.js"></script>
    <script>
        var map = L.map('map').setView([$mapLat, $mapLon], 15);
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; OpenStreetMap contributors'
        }).addTo(map);
        L.marker([$mapLat, $mapLon]).addTo(map)
            .bindPopup('$mapPopup')
            .openPopup();
    </script>
</body>
</html>
"@

    Set-Content -Path $htmlPath -Value $htmlContent
}
