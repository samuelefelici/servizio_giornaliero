ABSENCE_CODES = {
    "PC","AA","AF","AO","AP","AS","AM","PADm","PAD","PA",
    "CM8","CM","PD","FP","FU","F1/2","F","AGP","I4","I91","I",
    "MA3","MA20","MA","ES","FI","S","PSn","PSm3","PSm2","PAL",
    "P626","PS30","PS","PZ","PE","PPR","PP","PVM","SG","NF" 
}

RESIDENZA_RENAME = {
    "CASTELFIDARDO": "C.FID.",
    "FILOTTRANO": "FILOT",
}
DEFAULT_SORT = ["Residenza","Categoria","Turno","Inizio"]  # editabile
HEADER_PROBE = "Cognome e Nome"
EXPECTED_COLUMNS = [
    "Cognome e Nome","Matricola","Categoria","Residenza",
    "Turno","Inizio","Fine","Indennità e note"
]
TITLE = "Servizio Giornaliero"
