# PV COSMOS : PV et charte iPhone créés automatiquement depuis un ticket COSMOS

1. Sur le ticket COSMOS (SCTASK ou INC) : favori **PV COSMOS**, puis bouton **Créer les PV**.
2. Excel (ouvert) crée les fichiers dans `Documents\COSMOS\NOM Prénom\AAAA-MM-JJ\` et ouvre le dossier.

## Fichiers

| Fichier | Rôle |
|---|---|
| `installer_favori.html` | Page pour installer le favori (glisser le bouton dans la barre des favoris). |
| `PVCosmosAuto.bas` | Module VBA à mettre dans PERSONAL.XLSB (surveillance du dossier Téléchargements, création des PV et de la charte). |
| `pv_cosmos.js` | Code source lisible du favori. |
| `build.js`, `installer_favori.template.html`, `pv_cosmos.min.js` | Régénèrent `installer_favori.html` après une modification de `pv_cosmos.js`. |

## Installation

1. Dossier `C:\Users\<toi>\Documents\COSMOS` contenant `PV_Vide.xlsm` et `Charte_Iphone_Vide.xlsm` (ils ne sont jamais modifiés).
2. Favori : ouvrir `installer_favori.html` et glisser le bouton **PV COSMOS** dans la barre des favoris.
3. Excel : créer le classeur de macros personnelles (Affichage > Macros > Enregistrer une macro… > Classeur de macros personnelles, puis arrêter l'enregistrement), Alt+F11, sélectionner VBAProject (PERSONAL.XLSB), Fichier > Importer `PVCosmosAuto.bas`. Enregistrer, fermer Excel puis le rouvrir.
   Réglages en haut du module : `DOSSIER_COSMOS` (vide = Documents\COSMOS), `TECHNICIEN`.
   Facultatif : Alt+F8 > PERSONAL.XLSB!Importer_Cosmos > Options > `I` (Ctrl+Maj+I = traiter tout de suite).

## Fonctionnement

- Le favori lit le ticket (formulaire, notes de travail, tâches de la même demande, fiche utilisateur, actifs affectés / réservés) et coche les documents à créer :
  Installation PC, Désinstallation PC, iPhone (changement / installation / restitution), Désinstallation PC + iPhone, Charte iPhone.
- Site, bâtiment, étage et bureau sont obligatoires (cases rouges). Ils sont cherchés dans l'emplacement du ticket, la ligne « Localisation : » de la description, la fiche utilisateur et l'emplacement du matériel en service.
- **Créer les PV** télécharge `PVCOSMOS_<ticket>.txt`. Le module Excel le détecte (toutes les 3 s), crée :
  - `NOM Prénom - Installation PC.xlsm`, `- Désinstallation PC.xlsm`, `- Changement iPhone.xlsm` (ou Installation / Restitution iPhone), `- Désinstallation PC + iPhone.xlsm` ;
  - `NOM Prénom - Charte iPhone.xlsm` et `.pdf` ;
  
  puis supprime la demande et ouvre le dossier. Un fichier déjà présent n'est jamais écrasé (suffixe `(2)`). Journal : `Documents\COSMOS\journal_pv.csv`.
- Doublons : le favori retient dans Edge les tickets traités. Sur le ticket, un bandeau vert s'affiche. Sur une liste COSMOS, un clic sur le favori colore en vert les tickets déjà faits.

## Sécurité

Le favori ne fait que lire ce que ton compte voit dans COSMOS. Il n'écrit rien dans COSMOS et ne transmet rien, sauf le fichier de demande enregistré dans ton dossier Téléchargements.
