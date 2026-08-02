/* ==========================================================================
   Chauffage Demarche Pascal — comportements
   Un seul fichier, sans dépendance. Le site reste utilisable sans lui : les
   liens restent des liens, le formulaire un formulaire.
   ========================================================================== */

(function () {
  "use strict";

  var DESTINATAIRE = "pascaldemarche@hotmail.be";

  /* ======================================================================
     DÉMONSTRATION — à retirer quand le site devient définitif

     `actif: false` suffit à livrer le site : la bannière disparaît, plus rien
     ne compte, le reste est inchangé. Voir README.md.

     Ceci s'exécute dans le navigateur du visiteur : c'est une échéance
     commerciale, pas une serrure.
     ====================================================================== */
  var DEMO = {
    actif: true,
    expire: "2026-08-10T20:00:00+02:00",
    contact: "info@ineart.be",
  };

  function restant() {
    return new Date(DEMO.expire).getTime() - Date.now();
  }

  /* La précision augmente à mesure que l'échéance approche, là où elle devient
     utile. Un décompte à la seconde sur une vitrine attire l'œil au mauvais
     endroit. */
  function formatRestant(ms) {
    var minutes = Math.floor(ms / 60000);
    var heures = Math.floor(minutes / 60);
    var jours = Math.floor(heures / 24);
    if (jours >= 2) return jours + " jours";
    if (heures >= 24) return "1 jour et " + (heures - 24) + " h";
    if (heures >= 1) return heures + " h " + String(minutes % 60).padStart(2, "0");
    return minutes + " min";
  }

  function expirer() {
    var page = document.createElement("div");
    page.className = "expire";
    page.innerHTML =
      '<div class="expire__box">' +
      '<svg viewBox="0 0 32 32" aria-hidden="true">' +
      '<path d="M16 3c3.4 4.6 5.7 7 5.7 10.3a5.7 5.7 0 0 1-11.4 0C10.3 10 12.6 7.6 16 3z" fill="#e2601c"/>' +
      '<path d="M16 29.5c-2.4 0-4.3-1.9-4.3-4.2 0-2.5 2.4-4.5 4.3-7 1.9 2.5 4.3 4.5 4.3 7 0 2.3-1.9 4.2-4.3 4.2z" fill="#1f6fb2"/>' +
      "</svg>" +
      "<h1>Cette démonstration est terminée</h1>" +
      "<p>Ce site était une proposition présentée à Chauffage Demarche Pascal " +
      "par IneWeb. La période de consultation est écoulée.</p>" +
      '<a class="btn btn--eau" href="mailto:' + DEMO.contact +
      '?subject=Site%20Chauffage%20Demarche">Nous écrire</a>' +
      "</div>";
    document.body.replaceChildren(page);
    document.title = "Démonstration terminée — Chauffage Demarche Pascal";
  }

  function demarrerDemo() {
    var banniere = document.querySelector("[data-demo]");
    if (!DEMO.actif || !banniere) return;
    if (restant() <= 0) return expirer();

    var cible = banniere.querySelector("[data-demo-countdown]");
    banniere.hidden = false;

    function tic() {
      var ms = restant();
      if (ms <= 0) return expirer();
      cible.textContent = formatRestant(ms);
    }

    tic();
    setInterval(tic, restant() > 3600000 ? 60000 : 10000);
  }

  demarrerDemo();

  /* --- Demande de devis ---------------------------------------------------
     Le formulaire compose un e-mail déjà rédigé plutôt que d'appeler un
     service tiers : pas de serveur à maintenir, pas de compte à créer, et
     aucune coordonnée de client qui transite par un intermédiaire.

     L'objet du message reprend la nature de la demande. Une chaudière en panne
     et une salle de bain à rénover n'appellent pas la même réactivité ; le
     voir dans sa boîte de réception sans ouvrir le message a de la valeur.
     ---------------------------------------------------------------------- */
  var form = document.querySelector("[data-quote]");

  if (form) {
    form.addEventListener("submit", function (event) {
      event.preventDefault();

      var data = new FormData(form);
      function lire(nom) {
        return (data.get(nom) || "").toString().trim();
      }

      var nom = lire("nom");
      var tel = lire("tel");
      var objet = lire("objet");
      var message = lire("message");

      var lignes = ["Bonjour,", ""];
      lignes.push(objet
        ? "Je vous contacte à propos de : " + objet.toLowerCase() + "."
        : "Je souhaite obtenir un devis.");
      if (message) lignes.push("", message);
      lignes.push("", "— " + (nom || "Nom à compléter"));
      lignes.push("Téléphone : " + (tel || "à compléter"));

      window.location.href =
        "mailto:" + DESTINATAIRE +
        "?subject=" + encodeURIComponent(objet || "Demande de devis") +
        "&body=" + encodeURIComponent(lignes.join("\n"));
    });
  }

  /* --- Navigation : marquer la section en cours ---------------------------
     Sur une page unique, savoir où l'on se trouve est la seule réponse à
     « où suis-je ». L'observateur ne fait qu'ajouter une couleur : s'il ne
     s'exécute pas, la navigation reste parfaitement utilisable.
     ---------------------------------------------------------------------- */
  var liens = document.querySelectorAll('.nav__links a[href^="#"]');
  if (!liens.length || !("IntersectionObserver" in window)) return;

  var parId = {};
  var sections = [];

  liens.forEach(function (lien) {
    var section = document.querySelector(lien.getAttribute("href"));
    if (!section) return;
    parId[section.id] = lien;
    sections.push(section);
  });

  var observateur = new IntersectionObserver(
    function (entrees) {
      entrees.forEach(function (entree) {
        var lien = parId[entree.target.id];
        if (!lien) return;
        lien.style.color = entree.isIntersecting ? "var(--noir)" : "";
      });
    },
    { rootMargin: "-45% 0px -50% 0px" }
  );

  sections.forEach(function (section) { observateur.observe(section); });
})();
