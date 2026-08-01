/* ==========================================================================
   Atelier Genot — comportements
   Un seul fichier, sans dépendance. Le site fonctionne entièrement sans lui :
   le formulaire retombe sur un envoi classique, la navigation reste des ancres.
   ========================================================================== */

(function () {
  "use strict";

  /* --- Destinataire du formulaire ----------------------------------------
     À REMPLIR : l'adresse de l'atelier. Une seule occurrence, ici.
     ---------------------------------------------------------------------- */
  var DESTINATAIRE = "contact@ateliergenot.be";

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
