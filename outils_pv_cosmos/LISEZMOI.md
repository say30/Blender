# PV COSMOS : remplir le PV Excel depuis un ticket COSMOS

Deux clics au lieu de recopier les champs à la main :

1. sur le ticket COSMOS, un clic sur le favori **PV COSMOS** ;
2. dans le PV Excel, **Ctrl + Maj + I**.

## Fichiers

| Fichier | Rôle |
|---|---|
| `installer_favori.html` | Page à ouvrir une fois pour installer le favori (glisser le bouton dans la barre des favoris). |
| `ImportCosmos.bas` | Module VBA à importer dans `PV_Vide.xlsm` (macro `Importer_Cosmos`, raccourci Ctrl+Maj+I). |
| `pv_cosmos.js` | Code source lisible du favori. |
| `build.js`, `installer_favori.template.html`, `pv_cosmos.min.js` | Régénèrent `installer_favori.html` après une modification de `pv_cosmos.js`. |

## Installation (une seule fois, sur le PC de travail)

### Le favori
1. Ouvre `installer_favori.html` dans Edge ou Chrome.
2. Fais glisser le bouton **PV COSMOS** dans la barre des favoris (Ctrl+Maj+B pour l'afficher).

### Le module Excel
Le PV est protégé (feuilles et projet VBA). On ne le modifie donc pas : le module va dans le **classeur de macros personnelles** (PERSONAL.XLSB) et agit sur le PV ouvert.
1. Dans Excel, va dans Affichage > Macros > Enregistrer une macro…, choisis « Classeur de macros personnelles », clique sur OK, puis arrête l'enregistrement.
2. Fais **Alt + F11**, sélectionne VBAProject (PERSONAL.XLSB), puis Fichier > Importer un fichier… > `ImportCosmos.bas`.
3. Alt+F8 > PERSONAL.XLSB!Importer_Cosmos > Options, puis tape le raccourci `I` majuscule (Ctrl+Maj+I).
4. Quand Excel le demande en fermant, enregistre le classeur de macros personnelles.

Les cellules remplies par la macro sont des cellules de saisie, non verrouillées : la protection des feuilles ne gêne pas. Si une cellule refuse l'écriture, la macro le signale.

## Ce qui est récupéré

| Champ du PV | Source dans COSMOS |
|---|---|
| N° dossier | « Numéro de la demande » (DEM…) |
| Nom, prénom | Fiche du bénéficiaire, sinon le champ « Bénéficiaire » |
| Site | Emplacement du bénéficiaire (ex. FRANCE/MEDITERRANEE/**NIMES**/…) |
| Bâtiment / étage / bureau | Fiche utilisateur, si COSMOS les a (sinon à saisir) |
| iPhone installé | Note de travail « Modèle / S/N / IMEI » du ticket, sinon l'iPhone réservé dans les actifs |
| iPhone restitué | Onglet « Détail de la demande » : « Matériel à récupérer ? » |
| PC restitué | Onglet « Détail de la demande » : « Matériel à restituer / désinstaller » (NetBIOS - S/N - modèle) |
| PC, station, écrans installés | Actifs affectés à / réservés pour le bénéficiaire (hors matériel à restituer) |

Le panneau montre aussi la liste des actifs du bénéficiaire. Les boutons **→ installé** et **→ restitué** permettent de corriger ou compléter un champ en un clic.

Les états, accessoires, tests et l'accompagnement restent à remplir dans le PV.

## Sécurité

Le favori ne fait que **lire** ce que ton compte voit déjà dans COSMOS : le ticket, les tâches de la même demande, la fiche utilisateur et les actifs. Il passe par l'API de COSMOS, ou par les pages normales si l'API est refusée. Il n'écrit rien et n'envoie rien ailleurs que dans ton presse-papier.
