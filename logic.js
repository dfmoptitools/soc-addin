// ======================================================
// LOGIQUE MÉTIER - équivalent du script VBA
// ======================================================

// Configuration - équivalent des Const VBA
const CONFIG = {
    PREFIX_SOC_ID: "SOC-",
    PATTERN_SOC_START: "[SOC x ",

    CAT_EN_COURS: "En cours Eitan",
    CAT_ATTENTE_CLIENT: "Attente retour client",
    CAT_TRAITE: "Traité",

    MOTS_EXCLUS: ["SOC", "DFM", "X"],

    PREFIXES_SUJET: ["RE:", "TR:", "FW:", "FWD:", "RÉF.:", "REF:"],

    MARQUEURS_FIN_MAIL: [
        "Si cette application ne vous semble pas légitime",
        "Bien à vous",
        "Kind regards",
        "Cordialement",
        "Bonne journée",
        "Best regards"
    ],

    NOM_DOSSIER_TRAITEE: "Traitée",
    NOM_DOSSIER_ALERTES_SOC: "Alertes SOC"
};


// ======================================================
// Normalise une chaîne pour comparaison
// (équivalent VBA Normaliser)
// ======================================================
function normaliser(s) {
    if (!s) return "";
    return s.toUpperCase()
            .trim()
            .replace(/-/g, "")
            .replace(/_/g, "")
            .replace(/\s/g, "")
            .replace(/\./g, "");
}


// ======================================================
// Retire itérativement les préfixes du sujet
// Gère les chaînes type "RE: TR: RE: FW:"
// ======================================================
function nettoyerPrefixesSujet(sujet) {
    let resultat = (sujet || "").trim();
    let modifie = true;

    while (modifie) {
        modifie = false;
        for (const prefixe of CONFIG.PREFIXES_SUJET) {
            if (resultat.toUpperCase().startsWith(prefixe.toUpperCase())) {
                resultat = resultat.substring(prefixe.length).trim();
                modifie = true;
            }
        }
    }

    return resultat;
}


// ======================================================
// Coupe le corps HTML au premier marqueur de fin trouvé
// ======================================================
function couperFinMail(htmlBody) {
    if (!htmlBody) return "";

    let posMin = -1;
    const lowerBody = htmlBody.toLowerCase();

    for (const marqueur of CONFIG.MARQUEURS_FIN_MAIL) {
        const pos = lowerBody.indexOf(marqueur.toLowerCase());
        if (pos > -1) {
            if (posMin === -1 || pos < posMin) {
                posMin = pos;
            }
        }
    }

    if (posMin > -1) {
        return htmlBody.substring(0, posMin);
    }
    return htmlBody;
}


// ======================================================
// Extrait l'ID client (le plus long nombre)
// et le nom de société (mots alphabétiques)
// depuis un sujet contenant [SOC x ...]
// ======================================================
function extraireIdEtSociete(sujet) {
    const resultat = { clientID: "", nomSociete: "" };

    const posDebut = sujet.toUpperCase().indexOf(CONFIG.PATTERN_SOC_START.toUpperCase());
    if (posDebut < 0) return resultat;

    const debut = posDebut + CONFIG.PATTERN_SOC_START.length;
    const posFin = sujet.indexOf("]", debut);
    if (posFin < 0) return resultat;

    let tmp = sujet.substring(debut, posFin).trim();

    // --- ÉTAPE 1 : Extraire le plus long nombre ---
    let longestNumber = "";
    let currentNumber = "";

    for (let i = 0; i < tmp.length; i++) {
        const c = tmp[i];
        if (c >= "0" && c <= "9") {
            currentNumber += c;
        } else {
            if (currentNumber.length > longestNumber.length) {
                longestNumber = currentNumber;
            }
            currentNumber = "";
        }
    }
    if (currentNumber.length > longestNumber.length) {
        longestNumber = currentNumber;
    }

    if (longestNumber) {
        // Enlever les zéros de gauche : "02756" -> "2756"
        resultat.clientID = String(parseInt(longestNumber, 10));
    }

    // --- ÉTAPE 2 : Extraire les mots alphabétiques ---
    // Remplacer séparateurs par des espaces
    let tmpMots = tmp.replace(/[_\-.\/]/g, " ");
    const words = tmpMots.split(/\s+/);
    const validWords = [];

    for (const w of words) {
        const word = w.trim();
        if (!word) continue;

        // Vérifier que ce sont uniquement des lettres
        if (!/^[a-zA-ZÀ-ÿ]+$/.test(word)) continue;

        // Vérifier qu'il n'est pas dans la liste d'exclusion
        if (CONFIG.MOTS_EXCLUS.includes(word.toUpperCase())) continue;

        validWords.push(word);
    }

    resultat.nomSociete = validWords.join(" ");
    return resultat;
}


// ======================================================
// Cherche les emails dans le fichier JSON de contacts
// Format attendu du JSON :
// [
//   { "clientId": "30824", "nomSociete": "ATDOMCO", "email": "x@y.com" },
//   { "clientId": "2756", "nomSociete": "I Tech Transfert", "email": "a@b.com" },
//   ...
// ]
// Match par ID OU par nomSociete normalisé
// ======================================================
function rechercherEmails(contacts, clientID, nomSociete) {
    if (!Array.isArray(contacts)) return [];

    const nomNormalise = normaliser(nomSociete);
    const emails = new Set();

    for (const c of contacts) {
        if (!c.email) continue;

        // Match par ID
        if (clientID && c.clientId && String(c.clientId) === String(clientID)) {
            emails.add(c.email);
            continue;
        }

        // Match par nom de société normalisé
        if (nomNormalise && c.nomSociete) {
            if (normaliser(c.nomSociete) === nomNormalise) {
                emails.add(c.email);
            }
        }
    }

    return Array.from(emails);
}


// ======================================================
// Construit le corps HTML pour la réponse "Légitime"
// ======================================================
function corpsLegitime() {
    return "Bonjour,<br><br>" +
           "Cette action est bien légitime.<br>" +
           "Merci pour votre alerte.<br>";
}


// ======================================================
// Construit le corps HTML pour la réponse "Clôture"
// ======================================================
function corpsCloture() {
    return "Bonjour,<br><br>" +
           "Merci pour votre retour. Je prends bien note et clôture l'alerte.<br><br>" +
           "En vous souhaitant une belle journée.<br>";
}


// ======================================================
// Remplace "Bonjour à tous" par "Bonjour" dans le corps
// (équivalent VBA)
// ======================================================
function corrigerBonjour(htmlBody) {
    if (!htmlBody) return "";
    return htmlBody.replace(/Bonjour à tous,/gi, "Bonjour,");
}
