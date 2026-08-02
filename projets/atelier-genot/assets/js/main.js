/* ==========================================================================
   Atelier Genot — comportements
   Un seul fichier, sans dépendance. Le site fonctionne entièrement sans lui :
   le formulaire retombe sur un envoi classique, la navigation reste des ancres.
   ========================================================================== */

(function () {
  "use strict";

  /* --- Destinataire du formulaire ----------------------------------------
     L'adresse de l'atelier. Une seule occurrence, ici.
     ---------------------------------------------------------------------- */
  var DESTINATAIRE = "ateliergenot@hotmail.com";

  /* ======================================================================
     DÉMONSTRATION — à retirer quand le site devient définitif

     Le site s'annonce comme une proposition et affiche le temps qu'il lui
     reste. Passé la date, la page est remplacée par un écran d'expiration.

     Pour livrer le site pour de bon : mettre `actif` à false. La bannière
     disparaît, plus rien ne compte, le reste du site est inchangé.

     À SAVOIR : ceci s'exécute dans le navigateur du visiteur. Quelqu'un qui
     désactive JavaScript ou recule l'horloge de sa machine verra la page
     malgré tout. C'est une échéance commerciale, pas une serrure. Pour
     fermer réellement l'accès, il faut retirer les fichiers du serveur — le
     README explique comment automatiser ça.
     ====================================================================== */
  var DEMO = {
    actif: true,
    // Date et heure de fin, au format ISO. Fuseau belge : +02:00 en été,
    // +01:00 en hiver.
    expire: "2026-08-08T20:00:00+02:00",
    contact: "info@ineart.be",
  };

  function restant() {
    return new Date(DEMO.expire).getTime() - Date.now();
  }

  /* Formulation volontairement approximative : « encore 5 jours » se lit
     mieux que « 4 j 23 h 11 min 06 s », et un décompte à la seconde sur une
     page de vitrine attire l'œil au mauvais endroit. La précision augmente
     à mesure que l'échéance approche, là où elle devient utile. */
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
      '<img class="expire__logo" src="assets/img/logo.png" alt="Atelier Genot" width="398" height="320">' +
      "<h1>Cette démonstration est terminée</h1>" +
      "<p>Ce site était une proposition présentée à l'Atelier Genot par IneWeb. " +
      "La période de consultation est écoulée.</p>" +
      '<a class="btn btn--ember" href="mailto:' + DEMO.contact +
      '?subject=Site%20Atelier%20Genot">Nous écrire</a>' +
      "</div>";
    document.body.replaceChildren(page);
    document.title = "Démonstration terminée — Atelier Genot";
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
    // Une fois par minute suffit tant qu'il reste plus d'une heure ; on
    // resserre ensuite pour que le dernier quart d'heure reste juste.
    setInterval(tic, restant() > 3600000 ? 60000 : 10000);
  }

  demarrerDemo();

  /* --- Demande de devis ---------------------------------------------------
     Le formulaire compose un e-mail déjà rédigé plutôt que d'appeler un
     service tiers : pas de serveur à maintenir, pas de compte à créer, et
     aucune donnée du visiteur ne transite par un intermédiaire.
     ---------------------------------------------------------------------- */
  var form = document.querySelector("[data-quote]");

  if (form) {
    form.addEventListener("submit", function (event) {
      event.preventDefault();

      var data = new FormData(form);
      var nom = (data.get("nom") || "").toString().trim();
      var tel = (data.get("tel") || "").toString().trim();
      var type = (data.get("type") || "").toString().trim();
      var message = (data.get("message") || "").toString().trim();

      var sujet = "Demande de devis — " + (type || "projet");
      var corps = [
        "Bonjour,",
        "",
        "Je souhaite obtenir un devis pour : " + (type || "un projet") + ".",
        "",
        message || "(description à compléter)",
        "",
        "— " + (nom || "Nom à compléter"),
        "Téléphone : " + (tel || "à compléter"),
      ].join("\n");

      window.location.href =
        "mailto:" + DESTINATAIRE +
        "?subject=" + encodeURIComponent(sujet) +
        "&body=" + encodeURIComponent(corps);
    });
  }

  /* --- Navigation : marquer la section en cours ---------------------------
     Sur un site d'une seule page, savoir où l'on se trouve est la seule
     réponse à « où suis-je ». L'observateur ne fait qu'ajouter une classe :
     s'il ne s'exécute pas, la navigation reste parfaitement utilisable.
     ---------------------------------------------------------------------- */
  var links = document.querySelectorAll('.nav__links a[href^="#"]');
  if (!links.length || !("IntersectionObserver" in window)) return;

  var byId = {};
  var sections = [];

  links.forEach(function (link) {
    var section = document.querySelector(link.getAttribute("href"));
    if (!section) return;
    byId[section.id] = link;
    sections.push(section);
  });

  var observer = new IntersectionObserver(
    function (entries) {
      entries.forEach(function (entry) {
        var link = byId[entry.target.id];
        if (!link) return;
        link.style.color = entry.isIntersecting ? "var(--on-dark)" : "";
      });
    },
    { rootMargin: "-45% 0px -50% 0px" }
  );

  sections.forEach(function (section) { observer.observe(section); });
})();
