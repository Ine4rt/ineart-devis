/* ==========================================================================
   Garage J.C.A.S. — comportements
   Un seul fichier, sans dépendance. Le site reste utilisable sans lui : les
   cases à cocher restent des cases, les liens restent des liens.
   ========================================================================== */

(function () {
  "use strict";

  var DESTINATAIRE = "jcas.ceuppens@gmail.com";

  /* ======================================================================
     DÉMONSTRATION — à retirer quand le site devient définitif

     Voir README.md. `actif: false` suffit à livrer le site : la bannière
     disparaît, plus rien ne compte, le reste est inchangé.

     Ceci s'exécute dans le navigateur du visiteur : c'est une échéance
     commerciale, pas une serrure.
     ====================================================================== */
  var DEMO = {
    actif: true,
    expire: "2026-08-09T20:00:00+02:00",
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
      '<span class="expire__mark">JCAS</span>' +
      "<h1>Cette démonstration est terminée</h1>" +
      "<p>Ce site était une proposition présentée au Garage J.C.A.S. par IneWeb. " +
      "La période de consultation est écoulée.</p>" +
      '<a class="btn btn--jaune" href="mailto:' + DEMO.contact +
      '?subject=Site%20Garage%20JCAS">Nous écrire</a>' +
      "</div>";
    document.body.replaceChildren(page);
    document.title = "Démonstration terminée — Garage J.C.A.S.";
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

  /* --- Sélection des besoins ----------------------------------------------
     Le visiteur ne cherche pas une liste de prestations, il cherche ce qui ne
     va pas sur sa voiture. Les cases cochées composent le message ; le
     compteur montre que la sélection est prise en compte, sans quoi on ne sait
     pas si le clic a servi à quelque chose.
     ---------------------------------------------------------------------- */
  function besoinsCoches() {
    return Array.prototype.map.call(
      document.querySelectorAll('input[name="besoin"]:checked'),
      function (input) { return input.value; }
    );
  }

  var compteur = document.querySelector("[data-quote-count]");

  function majCompteur() {
    if (!compteur) return;
    var liste = besoinsCoches();
    if (!liste.length) {
      compteur.textContent = "Aucun besoin sélectionné pour l'instant.";
      return;
    }
    /* La liste énumérée plutôt qu'un simple total : « Freins, Pneus » se
       vérifie d'un coup d'œil, « 2 besoins » demande de recompter. */
    compteur.innerHTML = "Votre demande : <b>" + liste.join(", ") + "</b>";
  }

  document.querySelectorAll('input[name="besoin"]').forEach(function (input) {
    input.addEventListener("change", majCompteur);
  });
  majCompteur();

  /* --- Envoi de la demande -------------------------------------------------
     Les deux formulaires — la grille de besoins et le bloc de contact —
     aboutissent au même e-mail. Chacun apporte ce qu'il a : le premier la
     liste des besoins, le second les coordonnées et le véhicule. Composer un
     message plutôt qu'appeler un service tiers évite un compte de plus, un
     abonnement de plus, et des coordonnées de clients hébergées ailleurs.
     ---------------------------------------------------------------------- */
  function texte(nom) {
    var champ = document.querySelector('[name="' + nom + '"]');
    return champ ? champ.value.trim() : "";
  }

  function envoyer(event) {
    event.preventDefault();

    var besoins = besoinsCoches();
    var nom = texte("nom");
    var tel = texte("tel");
    var vehicule = texte("vehicule");
    var message = texte("message");

    var sujet = besoins.length
      ? "Demande de rendez-vous — " + besoins.join(", ")
      : "Demande de rendez-vous";

    var lignes = ["Bonjour,", ""];
    lignes.push(besoins.length
      ? "Je souhaite un rendez-vous pour : " + besoins.join(", ") + "."
      : "Je souhaite un rendez-vous au garage.");
    if (vehicule) lignes.push("", "Véhicule : " + vehicule);
    if (message) lignes.push("", message);
    lignes.push("", "— " + (nom || "Nom à compléter"));
    lignes.push("Téléphone : " + (tel || "à compléter"));

    window.location.href =
      "mailto:" + DESTINATAIRE +
      "?subject=" + encodeURIComponent(sujet) +
      "&body=" + encodeURIComponent(lignes.join("\n"));
  }

  document.querySelectorAll("[data-quote]").forEach(function (form) {
    form.addEventListener("submit", envoyer);
  });

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
        lien.style.color = entree.isIntersecting ? "var(--nuit)" : "";
      });
    },
    { rootMargin: "-45% 0px -50% 0px" }
  );

  sections.forEach(function (section) { observateur.observe(section); });
})();
