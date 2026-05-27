// ======================================================
// INTERFACE TASKPANE - boutons & Office.js
// ======================================================

let contactsData = null;

Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;

    // Charger les contacts du localStorage si déjà importés
    const stored = localStorage.getItem("socContacts");
    if (stored) {
        try {
            contactsData = JSON.parse(stored);
            afficherInfoContacts();
        } catch (e) {
            console.error("Contacts JSON invalides en localStorage", e);
            localStorage.removeItem("socContacts");
        }
    }

    // Liaison des boutons
    document.getElementById("btnLoadContacts").addEventListener("click", () => {
        document.getElementById("jsonFile").click();
    });

    document.getElementById("jsonFile").addEventListener("change", chargerContacts);
    document.getElementById("btnClearContacts").addEventListener("click", effacerContacts);

    document.getElementById("btnAlerteSOC").addEventListener("click", actionAlerteSOC);
    document.getElementById("btnLegitime").addEventListener("click", actionLegitime);
    document.getElementById("btnCloture").addEventListener("click", actionCloture);
});


// ======================================================
// Gestion du fichier de contacts
// ======================================================
function chargerContacts(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = JSON.parse(e.target.result);
            if (!Array.isArray(data)) {
                afficherStatus("Le JSON doit être un tableau []", "error");
                return;
            }
            contactsData = data;
            localStorage.setItem("socContacts", JSON.stringify(data));
            afficherInfoContacts();
            afficherStatus(`✅ ${data.length} contacts chargés`, "success");
        } catch (err) {
            afficherStatus("Erreur de lecture JSON : " + err.message, "error");
        }
    };
    reader.readAsText(file);
}

function effacerContacts() {
    if (!confirm("Effacer les contacts importés ?")) return;
    localStorage.removeItem("socContacts");
    contactsData = null;
    afficherInfoContacts();
    afficherStatus("Contacts effacés", "info");
}

function afficherInfoContacts() {
    const info = document.getElementById("contactsInfo");
    const btnClear = document.getElementById("btnClearContacts");
    if (contactsData && contactsData.length) {
        info.textContent = `${contactsData.length} contacts chargés ✅`;
        btnClear.classList.remove("hidden");
    } else {
        info.textContent = "Aucun fichier de contacts chargé.";
        btnClear.classList.add("hidden");
    }
}


// ======================================================
// Affichage statut + debug
// ======================================================
function afficherStatus(message, type) {
    const el = document.getElementById("status");
    el.textContent = message;
    el.className = "status " + type;
    el.classList.remove("hidden");
}

function afficherDebug(infos) {
    document.getElementById("debugInfo").textContent = JSON.stringify(infos, null, 2);
}


// ======================================================
// ACTION 1 : ALERTE SOC
// Équivalent VBA Bouton_Alerte_SOC
// ======================================================
async function actionAlerteSOC() {
    try {
        const item = Office.context.mailbox.item;
        if (!item) {
            afficherStatus("Aucun mail sélectionné", "error");
            return;
        }

        // 1. Nettoyer le sujet
        const sujetOriginal = item.subject || "";
        const sujetNettoye = nettoyerPrefixesSujet(sujetOriginal);

        // 2. Extraire ID et nom société
        const { clientID, nomSociete } = extraireIdEtSociete(sujetNettoye);

        afficherDebug({
            sujetOriginal,
            sujetNettoye,
            clientID,
            nomSociete
        });

        if (!clientID && !nomSociete) {
            afficherStatus("⚠️ Aucun ID/nom trouvé dans le sujet. Format attendu : [SOC x <id>_<nom>]", "warning");
            return;
        }

        // 3. Chercher les destinataires
        let emails = [];
        if (contactsData) {
            emails = rechercherEmails(contactsData, clientID, nomSociete);
        }

        if (emails.length === 0) {
            afficherStatus(`⚠️ Aucun contact trouvé pour SOC-${clientID} / ${nomSociete}`, "warning");
        }

        // 4. Récupérer et nettoyer le corps HTML
        const htmlBody = await getBodyAsync(item, Office.CoercionType.Html);
        let htmlNettoye = corrigerBonjour(htmlBody);
        htmlNettoye = couperFinMail(htmlNettoye);

        // 5. Ajouter les catégories sur le mail source
        await ajouterCategories(item, [CONFIG.CAT_EN_COURS, CONFIG.CAT_ATTENTE_CLIENT]);

        // 6. Ouvrir un nouveau mail pré-rempli
        Office.context.mailbox.displayNewMessageFormAsync({
            toRecipients: emails,
            subject: sujetNettoye,
            htmlBody: htmlNettoye,
            // L'importance haute n'est pas configurable via displayNewMessageFormAsync
            // L'utilisateur doit la définir manuellement dans le nouveau mail si besoin
        }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                afficherStatus(`✅ Mail créé (${emails.length} destinataire(s))`, "success");
            } else {
                afficherStatus("Erreur création mail : " + result.error.message, "error");
            }
        });

    } catch (err) {
        afficherStatus("Erreur : " + err.message, "error");
        console.error(err);
    }
}


// ======================================================
// ACTION 2 : LÉGITIME
// Équivalent VBA Bouton_Repondre_Tous_Legitime
// ======================================================
async function actionLegitime() {
    try {
        const item = Office.context.mailbox.item;
        if (!item) {
            afficherStatus("Aucun mail sélectionné", "error");
            return;
        }

        // 1. Ajouter les catégories
        await ajouterCategories(item, [CONFIG.CAT_TRAITE, CONFIG.CAT_EN_COURS]);

        // 2. Ouvrir la réponse à tous pré-remplie
        // Note : on utilise displayReplyAllForm avec un objet pour pré-remplir le corps
        item.displayReplyAllForm({
            htmlBody: corpsLegitime()
        });

        afficherStatus("✅ Réponse créée. Le déplacement vers 'Traitée' doit être fait manuellement.", "info");

        // ⚠️ IMPORTANT : le déplacement vers "Alertes SOC > Traitée"
        // nécessite Microsoft Graph API (l'API Office.js de base ne permet
        // pas de lister les dossiers ni de déplacer un mail par nom de dossier).
        // Voir README.md pour activer cette fonctionnalité.

    } catch (err) {
        afficherStatus("Erreur : " + err.message, "error");
        console.error(err);
    }
}


// ======================================================
// ACTION 3 : CLÔTURE
// Équivalent VBA Bouton_Repondre_Tous_Cloture
// ======================================================
async function actionCloture() {
    try {
        const item = Office.context.mailbox.item;
        if (!item) {
            afficherStatus("Aucun mail sélectionné", "error");
            return;
        }

        // 1. Ajouter la catégorie
        await ajouterCategories(item, [CONFIG.CAT_TRAITE]);

        // 2. Ouvrir la réponse à tous pré-remplie
        item.displayReplyAllForm({
            htmlBody: corpsCloture()
        });

        afficherStatus("✅ Réponse de clôture créée", "success");

    } catch (err) {
        afficherStatus("Erreur : " + err.message, "error");
        console.error(err);
    }
}


// ======================================================
// HELPERS OFFICE.JS
// ======================================================

// Récupère le corps HTML d'un mail (async)
function getBodyAsync(item, coercionType) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(coercionType, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// Récupère les catégories actuelles
function getCategoriesAsync(item) {
    return new Promise((resolve, reject) => {
        item.categories.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || []);
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// Ajoute des catégories si absentes
// ⚠️ NOTE IMPORTANTE : les catégories doivent EXISTER au préalable
// dans la liste des catégories maîtresses Outlook, sinon l'ajout échoue.
// Crée-les manuellement dans Outlook avant utilisation :
// "En cours Eitan", "Attente retour client", "Traité"
async function ajouterCategories(item, nouvellesCategories) {
    const actuelles = await getCategoriesAsync(item);
    const noms = actuelles.map(c => c.displayName);

    const aAjouter = nouvellesCategories
        .filter(nom => !noms.some(n => n.toLowerCase() === nom.toLowerCase()))
        .map(nom => ({ displayName: nom }));

    if (aAjouter.length === 0) return;

    return new Promise((resolve, reject) => {
        item.categories.addAsync(aAjouter.map(c => c.displayName), (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            } else {
                // Erreur fréquente : catégorie non définie dans la mailbox
                console.warn("Catégorie non ajoutée : " + result.error.message);
                resolve(); // on continue quand même
            }
        });
    });
}
