# Banc d'essai de l'export « Groupes et EDT »

Le bouton **⬇︎ Excel** de l'écran Groupes et EDT produit le planning mural du
club (onglets `edt`, `prof`, `terrains`, `joueurs`) plus les quatre onglets de
détail. Le relire depuis le navigateur demande de cliquer, d'ouvrir Excel, et de
recommencer à chaque retouche. Ce banc fait la même chose en une commande.

```bash
cd site/banc
npm install          # une seule fois : il ne tire qu'ExcelJS
node banc-edt.js --apercu
```

Le classeur sort dans `sortie/groupes-cca.xlsx`, et `--apercu` écrit en plus
`sortie/apercu-edt.html` — l'onglet `edt` rejoué en HTML avec ses couleurs, ses
cadres et ses fusions. C'est le moyen de juger la mise en page sans ouvrir Excel :
un pavé qui ne tombe pas juste s'y voit tout de suite.

## Options

| option | effet |
|---|---|
| `--apercu` | écrit aussi l'aperçu HTML de l'onglet `edt` |
| `--sans-contraintes` | ignore `contraintes-tennis.xlsx` — le chemin dégradé : ni budget de prof, ni plage de courts |
| `--etat=<chemin>` | un autre état que `../../solveur/etat-ecran.json` |
| `--sortie=<dossier>` | un autre dossier de sortie que `sortie/` |

## Ce qu'il teste, et ce qu'il ne teste pas

Le code n'est **pas** dupliqué ici : le script lit `../index.html`, en extrait le
dernier bloc `<script>` — celui de l'écran —, le déballe de son IIFE et
l'exécute sous node avec un DOM en carton. Il appelle ensuite les vraies
fonctions `feuilleEdt`, `feuilleProfs`, `feuilleTerrains`, `feuilleJoueurs` et
`feuillesDetail`. Si l'une d'elles est renommée, le banc s'arrête en le disant
plutôt que de passer à côté.

En revanche il ne teste ni l'écran lui-même, ni le chargement d'ExcelJS depuis
le CDN, ni le téléchargement : tout ce qui touche au navigateur est remplacé par
des bouchons. Une régression sur un bouton ne se verra pas ici.

L'état de départ est celui que l'agent écrit à côté des classeurs
(`solveur/etat-ecran.json`) : le banc travaille donc sur le planning réel du
club, pas sur des données inventées.

## Et le déploiement ?

Ce dossier est exclu de Vercel par `../.vercelignore` : le site publié reste
`index.html` et rien d'autre.
