/**
 * build-i18n.js — generate the translated homepages from the English source.
 *
 * The English root index.html is the single source of truth. This script emits
 * es/index.html, fr/index.html, de/index.html and it/index.html with identical
 * structure/markup — only the text, <html lang>, language-switcher active state,
 * asset paths (made absolute) and SEO head (title/description/canonical/og/
 * hreflang) differ.
 *
 * RULE (see memory multilang-site-sync): edit the English page + this dictionary,
 * then re-run `node tools/build-i18n.js` so all five languages stay in sync.
 * Never hand-edit a generated page — your change would be overwritten.
 *
 * Reviews are authentic guest testimonials and are intentionally left in their
 * original language.
 */
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const SRC = path.join(ROOT, "index.html");
const LANGS = ["es", "fr", "de", "it"];

const META = {
  es: { locale: "es_ES", title: "Free Tour Barcelona a pie: Sagrada Família, Gaudí y Barrio Gótico | Roots & Roads", desc: "Tour a pie por Barcelona de reserva gratuita con guías locales de Roots & Roads. Descubre la Sagrada Família, Gaudí, el Passeig de Gràcia, la Casa Milà, la Casa Batlló y el Barrio Gótico en 3 horas.", aiPath: "/es/resumen-ia.md", aiTitle: "Resumen para IA de Roots & Roads Barcelona" },
  fr: { locale: "fr_FR", title: "Visite à pied de Barcelone : Sagrada Família, Gaudí et Quartier gothique | Roots & Roads", desc: "Visite à pied de Barcelone à réservation gratuite avec les guides locaux Roots & Roads. Découvrez la Sagrada Família, Gaudí, le Passeig de Gràcia, la Casa Milà, la Casa Batlló et le Quartier gothique en 3 heures.", aiPath: "/fr/resume-ia.md", aiTitle: "Résumé pour l'IA de Roots & Roads Barcelona" },
  de: { locale: "de_DE", title: "Barcelona Rundgang: Sagrada Família, Gaudí & Gotisches Viertel | Roots & Roads", desc: "Barcelona-Rundgang mit freier Reservierung und einheimischen Roots & Roads Guides. Sagrada Família, Gaudí, Passeig de Gràcia, Casa Milà, Casa Batlló und das Gotische Viertel in 3 Stunden.", aiPath: "/de/ki-zusammenfassung.md", aiTitle: "KI-Zusammenfassung von Roots & Roads Barcelona" },
  it: { locale: "it_IT", title: "Tour a piedi di Barcellona: Sagrada Família, Gaudí e Quartiere Gotico | Roots & Roads", desc: "Tour a piedi di Barcellona a prenotazione gratuita con guide locali Roots & Roads. Sagrada Família, Gaudí, Passeig de Gràcia, Casa Milà, Casa Batlló e il Quartiere Gotico in 3 ore.", aiPath: "/it/riassunto-ia.md", aiTitle: "Riepilogo per l'IA di Roots & Roads Barcelona" }
};

// Each entry: English source string -> per-language translation.
// Longer / more specific strings first to avoid partial-match collisions.
const D = [
  // ---- Hero / tagline / description (multiline inner blocks) ----
  ["Barcelona's<br>\n            Highlights<br>\n            Walking<br>\n            Tour", {
    es: "Ruta a pie<br>\n            por lo mejor<br>\n            de<br>\n            Barcelona",
    fr: "Visite à pied<br>\n            des incontournables<br>\n            de<br>\n            Barcelone",
    de: "Barcelonas<br>\n            Höhepunkte<br>\n            zu Fuß<br>\n            erleben",
    it: "I luoghi<br>\n            simbolo di<br>\n            Barcellona<br>\n            a piedi"
  }],
  ["Three hours, three Barcelonas,<br>\nwith local guides who truly know them all", {
    es: "Tres horas, tres Barcelonas,<br>\ncon guías locales que las conocen de verdad",
    fr: "Trois heures, trois Barcelone,<br>\navec des guides locaux qui les connaissent vraiment",
    de: "Drei Stunden, drei Barcelonas,<br>\nmit einheimischen Guides, die sie wirklich kennen",
    it: "Tre ore, tre Barcellona,<br>\ncon guide locali che le conoscono davvero"
  }],
  ["Our free-reserve Barcelona walking tour begins beneath\n<strong>Sagrada Familia</strong>, continues along\n<strong>Passeig de Gràcia</strong> to explore\n<strong>Catalan Modernism</strong>, and ends in the historical\nmaze of the city, the <strong>Gothic Quarter</strong>.", {
    es: "Nuestro tour a pie por Barcelona, de reserva gratuita, comienza bajo la\n<strong>Sagrada Familia</strong>, continúa por el\n<strong>Passeig de Gràcia</strong> para descubrir el\n<strong>Modernismo catalán</strong> y termina en el laberinto\nhistórico de la ciudad, el <strong>Barrio Gótico</strong>.",
    fr: "Notre visite à pied de Barcelone, à réservation gratuite, commence sous la\n<strong>Sagrada Família</strong>, se poursuit le long du\n<strong>Passeig de Gràcia</strong> pour explorer le\n<strong>Modernisme catalan</strong>, et se termine dans le dédale\nhistorique de la ville, le <strong>Quartier gothique</strong>.",
    de: "Unser Barcelona-Rundgang mit freier Reservierung beginnt unter der\n<strong>Sagrada Família</strong>, führt weiter über den\n<strong>Passeig de Gràcia</strong>, um den\n<strong>katalanischen Modernisme</strong> zu entdecken, und endet im historischen\nLabyrinth der Stadt, dem <strong>Gotischen Viertel</strong>.",
    it: "Il nostro tour a piedi di Barcellona, a prenotazione gratuita, inizia sotto la\n<strong>Sagrada Família</strong>, prosegue lungo il\n<strong>Passeig de Gràcia</strong> per scoprire il\n<strong>Modernismo catalano</strong> e termina nel labirinto\nstorico della città, il <strong>Quartiere Gotico</strong>."
  }],

  // ---- Nav ----
  [">BOOK NOW</a>", { es: ">RESERVAR</a>", fr: ">RÉSERVER</a>", de: ">BUCHEN</a>", it: ">PRENOTA</a>" }],
  [">ABOUT US</a>", { es: ">NOSOTROS</a>", fr: ">À PROPOS</a>", de: ">ÜBER UNS</a>", it: ">CHI SIAMO</a>" }],
  [">GUIDE PORTAL</a>", { es: ">GUÍAS</a>", fr: ">GUIDES</a>", de: ">GUIDE-PORTAL</a>", it: ">GUIDE</a>" }],

  // ---- Section titles ----
  [">What you’ll see</h2>", { es: ">Qué verás</h2>", fr: ">Ce que vous verrez</h2>", de: ">Was Sie sehen werden</h2>", it: ">Cosa vedrai</h2>" }],
  [">How it works</h2>", { es: ">Cómo funciona</h2>", fr: ">Comment ça marche</h2>", de: ">So funktioniert es</h2>", it: ">Come funziona</h2>" }],
  [">From those who have trusted in us</h2>", { es: ">De quienes han confiado en nosotros</h2>", fr: ">De ceux qui nous ont fait confiance</h2>", de: ">Von denen, die uns vertraut haben</h2>", it: ">Da chi si è fidato di noi</h2>" }],
  ["<h2>About Us</h2>", { es: "<h2>Sobre nosotros</h2>", fr: "<h2>À propos de nous</h2>", de: "<h2>Über uns</h2>", it: "<h2>Chi siamo</h2>" }],

  // ---- Tour stops ----
  ["<h3>Meeting point: KFC Sagrada Família</h3>", { es: "<h3>Punto de encuentro: KFC Sagrada Família</h3>", fr: "<h3>Point de rencontre : KFC Sagrada Família</h3>", de: "<h3>Treffpunkt: KFC Sagrada Família</h3>", it: "<h3>Punto d’incontro: KFC Sagrada Família</h3>" }],
  ["Look for us under a blue umbrella at Av. Gaudí 2", { es: "Búscanos bajo un paraguas azul en Av. Gaudí 2", fr: "Cherchez-nous sous un parapluie bleu au 2 Av. Gaudí", de: "Suchen Sie uns unter einem blauen Schirm in der Av. Gaudí 2", it: "Cercaci sotto un ombrello blu in Av. Gaudí 2" }],

  ["<h3>Sagrada Familia: Past and Present</h3>", { es: "<h3>Sagrada Familia: pasado y presente</h3>", fr: "<h3>Sagrada Família : passé et présent</h3>", de: "<h3>Sagrada Família: Vergangenheit und Gegenwart</h3>", it: "<h3>Sagrada Família: passato e presente</h3>" }],
  ["Discover the history of the Basilica and the details hidden in the facades of the birth, passion and glory", { es: "Descubre la historia de la Basílica y los detalles ocultos en las fachadas del nacimiento, la pasión y la gloria", fr: "Découvrez l’histoire de la basilique et les détails cachés dans les façades de la naissance, de la passion et de la gloire", de: "Entdecken Sie die Geschichte der Basilika und die verborgenen Details der Fassaden der Geburt, der Passion und der Herrlichkeit", it: "Scopri la storia della Basilica e i dettagli nascosti nelle facciate della natività, della passione e della gloria" }],

  ["Casa Milà, aka La Pedrera, Gaudí's most controversial work", { es: "Casa Milà, conocida como La Pedrera, la obra más controvertida de Gaudí", fr: "Casa Milà, alias La Pedrera, l’œuvre la plus controversée de Gaudí", de: "Casa Milà, auch La Pedrera genannt, Gaudís umstrittenstes Werk", it: "Casa Milà, nota come La Pedrera, l’opera più controversa di Gaudí" }],

  ["<h3>Modernism on Passeig de Gràcia</h3>", { es: "<h3>Modernismo en el Passeig de Gràcia</h3>", fr: "<h3>Le modernisme sur le Passeig de Gràcia</h3>", de: "<h3>Modernisme am Passeig de Gràcia</h3>", it: "<h3>Modernismo sul Passeig de Gràcia</h3>" }],
  ["Explore how color, ornament and imagination transformed this avenue into a symbol of modern Barcelona", { es: "Descubre cómo el color, el ornamento y la imaginación convirtieron esta avenida en un símbolo de la Barcelona moderna", fr: "Explorez comment la couleur, l’ornement et l’imagination ont transformé cette avenue en symbole de la Barcelone moderne", de: "Erfahren Sie, wie Farbe, Ornament und Fantasie diese Prachtstraße zu einem Symbol des modernen Barcelona machten", it: "Scopri come colore, ornamento e immaginazione hanno trasformato questo viale in un simbolo della Barcellona moderna" }],

  ["<h3>Art, Roman Empire & A Kiss</h3>", { es: "<h3>Arte, Imperio romano y un beso</h3>", fr: "<h3>Art, Empire romain et un baiser</h3>", de: "<h3>Kunst, Römisches Reich & ein Kuss</h3>", it: "<h3>Arte, Impero romano e un bacio</h3>" }],
  ["Trace Barcelona's layers from Roman engineering to bohemian artists and a mural depicting freedom as a kiss", { es: "Recorre las capas de Barcelona, desde la ingeniería romana hasta los artistas bohemios y un mural que representa la libertad como un beso", fr: "Parcourez les strates de Barcelone, du génie romain aux artistes bohèmes et à une fresque représentant la liberté comme un baiser", de: "Verfolgen Sie die Schichten Barcelonas – von römischer Ingenieurskunst über Bohème-Künstler bis zu einem Wandbild, das die Freiheit als Kuss darstellt", it: "Ripercorri gli strati di Barcellona, dall’ingegneria romana agli artisti bohémien e a un murale che raffigura la libertà come un bacio" }],

  ["Stand at the Roman gate opposite Picasso's mural and examine how the Cathedral reshaped Barcelona's image", { es: "Sitúate en la puerta romana frente al mural de Picasso y descubre cómo la Catedral transformó la imagen de Barcelona", fr: "Tenez-vous devant la porte romaine face à la fresque de Picasso et découvrez comment la cathédrale a redéfini l’image de Barcelone", de: "Stehen Sie am römischen Tor gegenüber Picassos Wandbild und erfahren Sie, wie die Kathedrale das Bild Barcelonas prägte", it: "Fermati alla porta romana di fronte al murale di Picasso e scopri come la Cattedrale ha ridefinito l’immagine di Barcellona" }],

  ["<h3>Civil War, Myths & the Roman Empire</h3>", { es: "<h3>Guerra Civil, mitos y el Imperio romano</h3>", fr: "<h3>Guerre civile, mythes et Empire romain</h3>", de: "<h3>Bürgerkrieg, Mythen & das Römische Reich</h3>", it: "<h3>Guerra civile, miti e Impero romano</h3>" }],
  ["Travel from the Spanish Civil War scars to a secluded medieval bridge and hidden Roman columns", { es: "Viaja desde las cicatrices de la Guerra Civil española hasta un apartado puente medieval y columnas romanas ocultas", fr: "Voyagez des cicatrices de la guerre civile espagnole à un pont médiéval isolé et à des colonnes romaines cachées", de: "Reisen Sie von den Narben des Spanischen Bürgerkriegs zu einer versteckten mittelalterlichen Brücke und verborgenen römischen Säulen", it: "Viaggia dalle cicatrici della guerra civile spagnola a un appartato ponte medievale e a colonne romane nascoste" }],

  ["<h3>Power in Barcelona: From Rome to Today</h3>", { es: "<h3>El poder en Barcelona: de Roma a hoy</h3>", fr: "<h3>Le pouvoir à Barcelone : de Rome à aujourd’hui</h3>", de: "<h3>Macht in Barcelona: von Rom bis heute</h3>", it: "<h3>Il potere a Barcellona: da Roma a oggi</h3>" }],
  ["Stand where imperial Rome, medieval kings, Columbus, and today's Catalan government ruled Barcelona", { es: "Sitúate donde gobernaron Barcelona la Roma imperial, los reyes medievales, Colón y el actual gobierno catalán", fr: "Tenez-vous là où la Rome impériale, les rois médiévaux, Colomb et l’actuel gouvernement catalan ont gouverné Barcelone", de: "Stehen Sie dort, wo das kaiserliche Rom, mittelalterliche Könige, Kolumbus und die heutige katalanische Regierung Barcelona regierten", it: "Fermati dove hanno governato Barcellona la Roma imperiale, i re medievali, Colombo e l’attuale governo catalano" }],

  // ---- Booking form ----
  [">Book your Roots &amp; Roads tour</h2>", { es: ">Reserva tu tour de Roots &amp; Roads</h2>", fr: ">Réservez votre visite Roots &amp; Roads</h2>", de: ">Buchen Sie Ihre Roots &amp; Roads Tour</h2>", it: ">Prenota il tuo tour di Roots &amp; Roads</h2>" }],
  ["Reserve your spot in 30 seconds. We’ll confirm by email.", { es: "Reserva tu plaza en 30 segundos. Te confirmaremos por correo.", fr: "Réservez votre place en 30 secondes. Nous confirmerons par e-mail.", de: "Reservieren Sie Ihren Platz in 30 Sekunden. Wir bestätigen per E-Mail.", it: "Prenota il tuo posto in 30 secondi. Confermeremo via email." }],
  [">Language</label>", { es: ">Idioma</label>", fr: ">Langue</label>", de: ">Sprache</label>", it: ">Lingua</label>" }],
  [">Choose a language</option>", { es: ">Elige un idioma</option>", fr: ">Choisissez une langue</option>", de: ">Sprache wählen</option>", it: ">Scegli una lingua</option>" }],
  [">English</option>", { es: ">Inglés</option>", fr: ">Anglais</option>", de: ">Englisch</option>", it: ">Inglese</option>" }],
  [">Spanish</option>", { es: ">Español</option>", fr: ">Espagnol</option>", de: ">Spanisch</option>", it: ">Spagnolo</option>" }],
  [">German</option>", { es: ">Alemán</option>", fr: ">Allemand</option>", de: ">Deutsch</option>", it: ">Tedesco</option>" }],
  [">Italian</option>", { es: ">Italiano</option>", fr: ">Italien</option>", de: ">Italienisch</option>", it: ">Italiano</option>" }],
  [">French</option>", { es: ">Francés</option>", fr: ">Français</option>", de: ">Französisch</option>", it: ">Francese</option>" }],
  ['aria-label="Previous month"', { es: 'aria-label="Mes anterior"', fr: 'aria-label="Mois précédent"', de: 'aria-label="Voriger Monat"', it: 'aria-label="Mese precedente"' }],
  ['aria-label="Next month"', { es: 'aria-label="Mes siguiente"', fr: 'aria-label="Mois suivant"', de: 'aria-label="Nächster Monat"', it: 'aria-label="Mese successivo"' }],
  ["<div>mon</div><div>tue</div><div>wed</div><div>thu</div><div>fri</div><div>sat</div><div>sun</div>", {
    es: "<div>lun</div><div>mar</div><div>mié</div><div>jue</div><div>vie</div><div>sáb</div><div>dom</div>",
    fr: "<div>lun</div><div>mar</div><div>mer</div><div>jeu</div><div>ven</div><div>sam</div><div>dim</div>",
    de: "<div>Mo</div><div>Di</div><div>Mi</div><div>Do</div><div>Fr</div><div>Sa</div><div>So</div>",
    it: "<div>lun</div><div>mar</div><div>mer</div><div>gio</div><div>ven</div><div>sab</div><div>dom</div>"
  }],
  ["> Available</span>", { es: "> Disponible</span>", fr: "> Disponible</span>", de: "> Verfügbar</span>", it: "> Disponibile</span>" }],
  ["> Fully booked</span>", { es: "> Completo</span>", fr: "> Complet</span>", de: "> Ausgebucht</span>", it: "> Al completo</span>" }],
  ["<label>Available times</label>", { es: "<label>Horarios disponibles</label>", fr: "<label>Horaires disponibles</label>", de: "<label>Verfügbare Zeiten</label>", it: "<label>Orari disponibili</label>" }],
  [">Choose a language and date first.</p>", { es: ">Primero elige un idioma y una fecha.</p>", fr: ">Choisissez d’abord une langue et une date.</p>", de: ">Wählen Sie zuerst eine Sprache und ein Datum.</p>", it: ">Scegli prima una lingua e una data.</p>" }],
  [">Name</label>", { es: ">Nombre</label>", fr: ">Nom</label>", de: ">Name</label>", it: ">Nome</label>" }],
  ['placeholder="Your name"', { es: 'placeholder="Tu nombre"', fr: 'placeholder="Votre nom"', de: 'placeholder="Ihr Name"', it: 'placeholder="Il tuo nome"' }],
  [">Email</label>", { es: ">Correo electrónico</label>", fr: ">E-mail</label>", de: ">E-Mail</label>", it: ">Email</label>" }],
  [">Phone number</label>", { es: ">Teléfono</label>", fr: ">Numéro de téléphone</label>", de: ">Telefonnummer</label>", it: ">Numero di telefono</label>" }],
  [">Number of guests</label>", { es: ">Número de asistentes</label>", fr: ">Nombre de participants</label>", de: ">Anzahl der Teilnehmer</label>", it: ">Numero di partecipanti</label>" }],
  [">Anything we should know? (optional)</label>", { es: ">¿Algo que debamos saber? (opcional)</label>", fr: ">Quelque chose à savoir ? (facultatif)</label>", de: ">Sollen wir etwas wissen? (optional)</label>", it: ">Qualcosa che dovremmo sapere? (facoltativo)</label>" }],
  ['placeholder="Special accommodations, requests, questions..."', { es: 'placeholder="Adaptaciones especiales, peticiones, preguntas..."', fr: 'placeholder="Aménagements particuliers, demandes, questions..."', de: 'placeholder="Besondere Wünsche, Anfragen, Fragen..."', it: 'placeholder="Esigenze particolari, richieste, domande..."' }],
  ["<span>I agree that Roots &amp; Roads uses my details to manage this booking and send tour-related messages.</span>", { es: "<span>Acepto que Roots &amp; Roads use mis datos para gestionar esta reserva y enviarme mensajes relacionados con el tour.</span>", fr: "<span>J’accepte que Roots &amp; Roads utilise mes informations pour gérer cette réservation et m’envoyer des messages liés à la visite.</span>", de: "<span>Ich stimme zu, dass Roots &amp; Roads meine Daten verwendet, um diese Buchung zu verwalten und tourbezogene Nachrichten zu senden.</span>", it: "<span>Accetto che Roots &amp; Roads utilizzi i miei dati per gestire questa prenotazione e inviarmi messaggi relativi al tour.</span>" }],
  [">Reserve your spot</button>", { es: ">Reserva tu plaza</button>", fr: ">Réservez votre place</button>", de: ">Platz reservieren</button>", it: ">Prenota il tuo posto</button>" }],

  // ---- How it works ----
  ["Our concept is simple: there is <strong>no set price</strong>.\nInstead, you tip your guide at the end of the tour\nwhatever amount feels right to you. Typically,\nguests give between <strong>15€ to 30€</strong>,\ndepending on their satisfaction and budget.", {
    es: "Nuestro concepto es sencillo: <strong>no hay un precio fijo</strong>.\nEn su lugar, al final del tour das a tu guía la propina\nque te parezca justa. Normalmente, los asistentes dan\nentre <strong>15€ y 30€</strong>,\nsegún su satisfacción y su presupuesto.",
    fr: "Notre concept est simple : il n’y a <strong>aucun prix fixe</strong>.\nÀ la place, vous donnez à votre guide, à la fin de la visite,\nle montant qui vous semble juste. En général, les participants\ndonnent entre <strong>15€ et 30€</strong>,\nselon leur satisfaction et leur budget.",
    de: "Unser Konzept ist einfach: Es gibt <strong>keinen festen Preis</strong>.\nStattdessen geben Sie Ihrem Guide am Ende der Tour\nden Betrag, der sich für Sie richtig anfühlt. In der Regel geben\nGäste zwischen <strong>15€ und 30€</strong>,\nje nach Zufriedenheit und Budget.",
    it: "Il nostro concetto è semplice: <strong>non c’è un prezzo fisso</strong>.\nInvece, alla fine del tour lasci alla tua guida la mancia\nche ritieni giusta. Di solito, i partecipanti danno\ntra <strong>15€ e 30€</strong>,\na seconda della soddisfazione e del budget."
  }],
  ["Our freelance guides work <strong>exclusively based on the\nremuneration of the guests</strong>. As any free tour,\nat the end of the tour you will decide how much the tour\nwas worth for you.", {
    es: "Nuestros guías autónomos trabajan <strong>exclusivamente en función\nde la remuneración de los asistentes</strong>. Como en todo free tour,\nal final del tour decidirás cuánto ha valido\nla experiencia para ti.",
    fr: "Nos guides indépendants travaillent <strong>exclusivement en fonction\nde la rémunération des participants</strong>. Comme toute visite libre,\nà la fin de la visite vous déciderez combien\nelle valait pour vous.",
    de: "Unsere freiberuflichen Guides arbeiten <strong>ausschließlich auf Basis\nder Vergütung durch die Gäste</strong>. Wie bei jeder kostenlosen Tour\nentscheiden Sie am Ende selbst, wie viel Ihnen\ndie Tour wert war.",
    it: "Le nostre guide freelance lavorano <strong>esclusivamente in base\nal compenso dei partecipanti</strong>. Come in ogni free tour,\nalla fine deciderai tu quanto è valsa\nl’esperienza per te."
  }],

  // ---- About ----
  ["We grew up in this city and now share it with you. Years of experience have given us stories and insight you won’t find in any script or guidebook, helping visitors connect the dots between the old, the new, and the everyday life of Barcelona. Join us and see Barcelona the way locals live it!", {
    es: "Crecimos en esta ciudad y ahora la compartimos contigo. Años de experiencia nos han dado historias y una mirada que no encontrarás en ningún guion ni guía, ayudando a los visitantes a conectar lo antiguo, lo nuevo y la vida cotidiana de Barcelona. ¡Únete y descubre Barcelona como la viven los locales!",
    fr: "Nous avons grandi dans cette ville et la partageons aujourd’hui avec vous. Des années d’expérience nous ont donné des histoires et un regard que vous ne trouverez dans aucun script ni guide, aidant les visiteurs à relier l’ancien, le nouveau et la vie quotidienne de Barcelone. Rejoignez-nous et découvrez Barcelone comme les habitants la vivent !",
    de: "Wir sind in dieser Stadt aufgewachsen und teilen sie nun mit Ihnen. Jahre der Erfahrung haben uns Geschichten und Einblicke geschenkt, die Sie in keinem Skript oder Reiseführer finden, und helfen Besuchern, das Alte, das Neue und den Alltag Barcelonas zu verbinden. Kommen Sie mit und erleben Sie Barcelona so, wie die Einheimischen es leben!",
    it: "Siamo cresciuti in questa città e ora la condividiamo con te. Anni di esperienza ci hanno regalato storie e uno sguardo che non troverai in nessun copione o guida, aiutando i visitatori a collegare l’antico, il nuovo e la vita quotidiana di Barcellona. Unisciti a noi e scopri Barcellona come la vivono i locali!"
  }],
  ["Contact us at:", { es: "Escríbenos a:", fr: "Contactez-nous à :", de: "Kontaktieren Sie uns unter:", it: "Scrivici a:" }],
  [">Guide Portal</a>", { es: ">Portal de guías</a>", fr: ">Portail guides</a>", de: ">Guide-Portal</a>", it: ">Portale guide</a>" }],

  // ---- Cookie-consent banner (uses curly ’ in FR so the single-quoted JS string can't break) ----
  ["We use cookies to measure traffic and for marketing. Accept to allow them, or reject to keep only the essential ones.", {
    es: "Usamos cookies para medir el tráfico y para marketing. Acepta para permitirlas o recházalas para conservar solo las esenciales.",
    fr: "Nous utilisons des cookies de mesure d’audience et de marketing. Acceptez pour les autoriser ou refusez pour ne garder que l’essentiel.",
    de: "Wir verwenden Cookies zur Reichweitenmessung und für Marketing. Akzeptieren Sie, um sie zu erlauben, oder lehnen Sie ab, um nur die essenziellen zu behalten.",
    it: "Usiamo cookie per misurare il traffico e per marketing. Accetta per consentirli o rifiuta per mantenere solo quelli essenziali."
  }],
  [">Reject</button>", { es: ">Rechazar</button>", fr: ">Refuser</button>", de: ">Ablehnen</button>", it: ">Rifiuta</button>" }],
  [">Accept</button>", { es: ">Aceptar</button>", fr: ">Accepter</button>", de: ">Akzeptieren</button>", it: ">Accetta</button>" }]
];

function transform(html, lang) {
  const m = META[lang];
  const misses = [];

  // 1. <html lang>
  html = html.replace('<html lang="en">', `<html lang="${lang}">`);

  // 2. Asset paths -> absolute (pages live in a /<lang>/ subfolder).
  html = html.split('href="css/').join('href="/css/');
  html = html.split('src="images/').join('src="/images/');
  html = html.split('href="images/').join('href="/images/');
  html = html.split('src="main_js.js"').join('src="/main_js.js"');

  // 3. Language switcher active state -> this language.
  html = html.replace('<a href="/" hreflang="en" class="lang-active" aria-current="true">EN</a>', '<a href="/" hreflang="en">EN</a>');
  html = html.replace(`<a href="/${lang}/" hreflang="${lang}">`, `<a href="/${lang}/" hreflang="${lang}" class="lang-active" aria-current="true">`);

  // 4. SEO head (self-canonical, localized title/description/og/twitter/locale).
  html = html.replace(/<title>[^<]*<\/title>/, `<title>${m.title}</title>`);
  html = html.replace('<link rel="canonical" href="https://rootsandroadsbcn.com/">', `<link rel="canonical" href="https://rootsandroadsbcn.com/${lang}/">`);
  html = html.replace('<meta property="og:url" content="https://rootsandroadsbcn.com/">', `<meta property="og:url" content="https://rootsandroadsbcn.com/${lang}/">`);
  html = html.replace('<meta property="og:locale" content="en_US">', `<meta property="og:locale" content="${m.locale}">`);
  const enDesc = "Free-reserve Barcelona walking tour with local Roots & Roads guides. See Sagrada Família, Gaudí, Passeig de Gràcia, Casa Milà, Casa Batlló and the Gothic Quarter in 3 hours.";
  html = html.split('content="' + enDesc + '"').join('content="' + m.desc + '"');
  html = html.replace('<meta property="og:title" content="Barcelona Walking Tour: Sagrada Família, Gaudí & Gothic Quarter | Roots & Roads">', `<meta property="og:title" content="${m.title}">`);
  html = html.replace('<meta name="twitter:title" content="Barcelona Walking Tour: Sagrada Família, Gaudí & Gothic Quarter | Roots & Roads">', `<meta name="twitter:title" content="${m.title}">`);
  const enOgDesc = "Free-reserve 3-hour Barcelona walking tour with local guides: Sagrada Família, Gaudí, Passeig de Gràcia, Casa Milà, Casa Batlló and the Gothic Quarter.";
  html = html.split('content="' + enOgDesc + '"').join('content="' + m.desc + '"');

  // Mark the page's structured data as being in this language. (The per-language
  // AI summaries are declared once in the shared head and inherited by all pages.)
  html = html.split('"inLanguage": "en"').join('"inLanguage": "' + lang + '"');

  // 5. Body text dictionary.
  for (const [en, tr] of D) {
    if (!html.includes(en)) { misses.push(en.slice(0, 45)); continue; }
    html = html.split(en).join(tr[lang]);
  }

  return { html, misses };
}

// Normalize CRLF -> LF so the multiline patterns match regardless of how git
// checked the file out.
const src = fs.readFileSync(SRC, "utf8").replace(/\r\n/g, "\n");
let totalMiss = 0;
for (const lang of LANGS) {
  const { html, misses } = transform(src, lang);
  const dir = path.join(ROOT, lang);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, "index.html"), html);
  console.log(`${lang}: wrote ${lang}/index.html (${Math.round(html.length / 1024)}KB)` + (misses.length ? `  ⚠ ${misses.length} unmatched` : "  ✓"));
  misses.forEach((s) => { console.log(`     MISS: ${s}`); totalMiss++; });
}
console.log(totalMiss ? `\n${totalMiss} unmatched source strings — fix and re-run.` : "\nAll strings matched.");
process.exit(totalMiss ? 1 : 0);
