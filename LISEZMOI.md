# XLSnip – installation (gratuite)

Un complément Excel doit être hébergé en HTTPS. GitHub Pages le fait gratuitement.

## 1. Héberger les fichiers
1. Créez un compte GitHub, puis un dépôt public nommé `xlsnip`.
2. Déposez-y : `taskpane.html`, `taskpane.js`, `logo.png`, `logo-mark.png`, `icon-16.png`, `icon-32.png`, `icon-80.png` (et, facultatif, le dossier `fonts/`).
3. Dépôt → Settings → Pages → Branch `main` / racine → Save.
4. L'adresse sera `https://VOTRE-COMPTE.github.io/xlsnip/`.

## 2. Adapter le manifeste
Dans `manifest.xml`, remplacez toutes les occurrences de `VOTRE-COMPTE` par votre nom de compte GitHub.

## Charte graphique
- Couleurs de la charte : or `#deba4d`, indigo `#4f5fa7`, bleu `#3f97d3` (dégradé bleu → indigo sur les boutons principaux, or en soulignement et accents). Le vert et le rouge ne servent qu'aux pointages Valide / Invalide.
- Le logo est affiché dans l'en-tête (`logo.png`) et dans l'écran d'accueil (`logo-mark.png`). Les icônes du ruban Excel reprennent le symbole du logo.
- Thème clair et sombre : le volet suit le thème du système.
- **Typographie Urbane** : police commerciale, non disponible en ligne. Par défaut, le volet utilise Montserrat (très proche). Pour activer Urbane, créez un dossier `fonts/` à côté de `taskpane.html` et déposez-y `Urbane-Regular.woff2`, `Urbane-DemiBold.woff2` et `Urbane-Bold.woff2` (conversion depuis vos fichiers de police, si la licence l'autorise).

## 3. Installer dans Excel
- **Excel sur le web** : Accueil → Compléments → Plus de compléments → Gérer mes compléments → Charger mon complément → choisir `manifest.xml`.
- **Excel Windows** : placez `manifest.xml` dans un dossier partagé (clic droit → Propriétés → Partage). Fichier → Options → Centre de gestion de la confidentialité → Paramètres du Centre → Catalogues de compléments approuvés → ajoutez le chemin réseau (`\\PC\dossier`), cochez « Afficher dans le menu ». Redémarrez Excel, puis Insertion → Mes compléments → Dossier partagé → XLSnip.
- **Excel Mac** : copiez `manifest.xml` dans `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`, redémarrez Excel, Insertion → Mes compléments.

Le bouton « Ouvrir XLSnip » apparaît ensuite dans l'onglet Accueil ; il ouvre et ferme le volet.

## 4. Utilisation
1. **+ Importer** : PDF ou images (JPG/PNG, convertis en PDF). Le texte natif est lu ; les pages scannées passent en OCR (français + anglais).
2. Les documents sont stockés **dans le fichier Excel** (enregistrez le classeur).
3. Choisissez **Texte**, **Nombre** ou **Tableau**, cliquez une cellule Excel, puis tracez une zone sur le document.
   Pour pointer, choisissez **✓ Valide** (symbole ✓, fond vert clair) ou **✗ Invalide** (symbole ✗, fond rouge clair) : la zone tracée n'a pas besoin de contenir du texte.
4. En cliquant plus tard sur une cellule extraite, le document s'ouvre à la bonne page avec la zone en bleu (option « Suivre la source »).

## Plusieurs zones dans une même case (Ctrl) et somme
- **Texte + Ctrl** : après un premier snip, maintenez **Ctrl** (⌘ sur Mac) en relâchant le clic d'une nouvelle zone : son texte est ajouté à la suite, séparé par un espace (« Quelle belle » + « journée ensoleillée » → « Quelle belle journée ensoleillée »). La cellule active doit rester celle du snip précédent.
- **Σ Somme** : additionne tous les nombres de la zone (une zone peut couvrir toute une colonne de montants). Avec **Ctrl**, la zone suivante s'ajoute au total de la case (2, 3 et 5 → 10).
- Les zones d'une même case sont toutes surlignées dans le document ; le clic droit supprime ou retire l'ensemble du snip.
- Ctrl ne s'applique pas aux modes Nombre, Tableau, Valide et Invalide.

## Compression des documents
- Le bouton **Compresser** (en haut du volet) demande confirmation, puis convertit chaque document en pages JPEG allégées (niveaux Standard, Forte ou Maximale). Les snips et le texte détecté sont conservés.
- Si un document n'est pas réellement plus léger après compression, l'original est conservé.
- Le texte détecté est aussi réenregistré sous une forme plus compacte (gain supplémentaire sur les documents scannés).
- Opération **irréversible** : gardez une copie du classeur. Enregistrez ensuite le classeur pour que le gain apparaisse.
- La taille de chaque document est indiquée dans la liste déroulante.

## Niveaux de gris et doublons
- **Niveaux de gris** : une fenêtre liste les documents avec des cases à cocher, et les liens **Tout sélectionner** / **Tout désélectionner**. Les documents cochés sont convertis en images en niveaux de gris (même niveau de compression que « Compresser »). Un document qui ne devient pas plus léger est laissé en couleur. Action irréversible.
- **Doublons** : détecte les documents strictement identiques (même contenu, quel que soit le nom), propose de supprimer les copies et conserve le premier exemplaire. Les snips des copies sont rattachés à l'exemplaire conservé, donc rien n'est perdu. Un document compressé ou converti n'est plus identique à son original.

## Suivi des snips
- Chaque snip est rattaché à sa cellule par une plage nommée masquée (`XLSnip_…`) : il suit les insertions/suppressions de lignes et de colonnes, les déplacements et le renommage de la feuille.
- Si les cellules sont supprimées, ou si leur contenu est entièrement effacé, la zone surlignée disparaît du document (au bout d'une demi-seconde environ) et la couleur de fond de la cellule est réinitialisée.
- Un snip de tableau n'est retiré que si **toutes** ses cellules sont vides.

## Agrandir / réduire la visionneuse
- Saisissez la **barre de séparation** située juste au-dessus du document (curseur ↕) et **faites glisser** en maintenant le clic : vers le haut, la zone d'outils (logo, import, options) se replie et la visionneuse s'agrandit ; vers le bas, elle se réduit.
- **Double-clic** sur la barre : replie tout (seul le bandeau du haut reste visible) ou rétablit la taille normale. Au clavier : barre sélectionnée, flèches ↑/↓, Début, Fin.
- Les modes de snip et la navigation entre pages restent toujours visibles. La hauteur choisie est mémorisée.
- Les outils d'allègement (Compresser, Niveaux de gris, Doublons) sont regroupés derrière le bouton **Alléger**, replié par défaut pour libérer de la place.
- La largeur du volet se règle avec son bord gauche (fonction native d'Excel).

## Clic droit sur une zone du document
- **Insérer l'image de la zone dans Excel** : l'extrait du document est inséré comme image à la position de la cellule active (largeur limitée à 360 pt).
- **Supprimer le snip et vider la cellule** : retire la zone, vide la cellule et réinitialise son formatage.
- **Retirer le snip (garder la valeur)** : retire la zone et le formatage, mais conserve la valeur dans la cellule.

## Limites
- Word/Excel/PowerPoint ne sont pas convertis automatiquement : exportez-les d'abord en PDF.
- L'OCR et les bibliothèques se chargent depuis Internet (pdf.js, Tesseract.js, jsPDF).
- Documents très lourds (> ~20 Mo) : le classeur grossit d'autant.
- La détection des colonnes est heuristique : vérifiez les tableaux complexes.
- Code non testé dans Excel : des ajustements sont probables.
