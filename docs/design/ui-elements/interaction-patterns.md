
# <a name="interaction-patterns-for-office-add-ins"></a>Modèles d’interaction pour les compléments Office


Les Compléments Office peuvent améliorer l’expérience de création et de productivité, tout en connectant le contenu des applications hôtes Office à des flux de travail web de plus grande envergure. Un certain nombre de scénarios courants s’appliquent aux composants de contenu, du volet Office et Outlook que vous pouvez développer. Cet article décrit certains des scénarios les plus courants et fournit des modèles d’interaction recommandés pour l’expérience utilisateur associée au complément. Vous pouvez décomposer, combiner ou mélanger et associer ces modèles d’interaction en fonction de vos scénarios uniques.

 **Scénarios courants pour les compléments**

| Type de complément | Scénarios courants |
| ------ | ------ |
|  Contenu  |  Visualisation de données <br> Widgets et outils  |
|  Volet de tâches  |  Transformation et traitement des données <br> Création plus efficace <br> Localisation de contenu et insertion de données <br> Publication ou téléchargement de contenu vers un service web  |
|  Outlook  |  Pontage entre le contenu du courrier électronique et une application externe <br> Informations supplémentaires sur le contenu d’un courrier électronique ou d’un rendez-vous <br> Informations contribuant à améliorer votre productivité  |

## <a name="visualize-data-with-a-content-add-in"></a>Visualisation des données avec un complément de contenu


Cet exemple présente un complément de contenu pour Excel qui génère un graphique à partir des données d’une feuille de calcul.

Dans ce modèle d’interaction, le complément ne devient actif que lorsque vous sélectionnez et liez des données pour générer le graphique. Il est important de communiquer l’objet du complément et la procédure d’activation dans la vue initiale du complément. 

**Complément de contenu pour Excel qui génère un graphique à partir des données d’une feuille de calcul**
<br>
![Application de contenu pour Excel qui génère un graphique à partir des données d’une feuille de calcul](../../images/off15appUXFig01.png)
<br>
<ul><li><p>Afin de renforcer l’idée selon laquelle l’utilisateur doit effectuer une action avant de choisir le bouton, affichez les instructions avec un bouton désactivé (A).</p></li><li><p>Une fois que vous avez sélectionné une plage de cellules, le bouton <span class="ui">Créer un graphique</span> devient actif (B-C).</p></li><li><p>La visualisation remplit le conteneur et remplace la vue précédente (D).</p></li><li><p>Affichez tout élément d’interface utilisateur supplémentaire sur le bord inférieur du complément avec un bouton de paramètres (engrenage) pour vous permettre d’accéder à une vue dans laquelle vous pouvez réinitialiser ou gérer le complément.</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments qui nécessitent que vous sélectionniez les données avant l’activation</p></li></ul>

## <a name="transform-content-with-a-task-pane-add-in"></a>Transformation du contenu avec un complément du volet Office


Cet exemple présente un complément du volet Office qui traduit le texte de votre document dans une autre langue.

Dans ce modèle d’interaction, vous devez d’abord sélectionner le texte à traduire dans le document.

**Complément du volet Office qui traduit le texte de votre document dans une autre langue**
<br>
![Application de volet de tâches qui traduit le texte de votre document dans une autre langue](../../images/off15appUXFig02.png)
<br>
<ul><li><p>Communiquez l’objet du complément dans un titre et indiquez que l’utilisateur doit d’abord effectuer une sélection (A).</p></li><li><p>Le menu de langue et le bouton <span class="ui">Traduire</span> sont désactivés, renforçant l’idée que l’utilisateur doit effectuer une action pour poursuivre. Après sélection du contenu dans le document, ces deux éléments deviennent actifs (D).</p></li><li><p>Une fois que l’utilisateur choisit <span class="ui">Traduire</span>, l’interface utilisateur se développe pour afficher le contenu traduit, ainsi qu’un bouton permettant de le réinsérer dans le document (E).</p></li><li><p>Vous pouvez fournir un bouton <span class="ui">Effacer</span> ou <span class="ui">Réinitialiser</span> qui rétablit la vue initiale.</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments qui nécessitent que vous sélectionniez des données avant l’activation</p></li><li><p>Interface utilisateur qui se déroule ou est révélée au fur et à mesure de votre progression dans un scénario</p></li></ul>

## <a name="process-data-with-a-task-pane-add-in"></a>Données de processus avec un complément du volet Office


Cet exemple présente un complément du volet Office qui vérifie les données dans Excel.

Dans ce modèle d’interaction, vous devez sélectionner une plage de cellules dans la feuille de calcul pour commencer.

**Complément du volet Office qui vérifie les données dans Excel**
<br>
![Application de volet de tâches qui vérifie des données dans Excel](../../images/off15appUXFig03.png)
<br>
<ul><li><p>L’objet du complément est décrit dans le titre. Les instructions vous aident à commencer.</p></li><li><p>Le bouton <span class="ui">Envoyer les données sélectionnées</span> est désactivé, renforçant l’idée que l’utilisateur doit effectuer une action pour progresser (A).</p></li><li><p>Une fois que l’utilisateur a sélectionné une plage de cellules dans sa feuille de calcul (B), le bouton <span class="ui">Envoyer les données sélectionnées</span> devient actif.</p></li><li><p>Une fois que l’utilisateur a cliqué sur ce bouton, l’interface utilisateur est remplacée par la plage de cellules sélectionnée, une barre de progression et un bouton <span class="ui">Annuler</span>.</p></li><li><p>La barre de progression indique l’état du processus et le bouton <span class="ui">Annuler</span> permet de l’interrompre (D).</p></li><li><p>Lorsque le processus est terminé, les résultats sont automatiquement affichés (E). La sélection d’un élément dans la liste active la cellule correspondante dans la feuille de calcul.</p></li></ul>Convient mieux pour :
<ul><li><p>Processus d’une durée indéterminée</p></li></ul>

## <a name="analyze-content-with-a-task-pane-add-in"></a>Analyse du contenu avec un complément du volet Office


Cet exemple présente un complément du volet Office qui affiche les définitions des mots que vous tapez.

Dans ce modèle d’interaction, vous devez d’abord sélectionner le texte dans le document pour afficher les résultats.

**Complément du volet Office qui affiche les définitions des mots au fur et à mesure de la saisie**
<br>
![Application de volet de tâches qui affiche les définitions de mot au fur et à mesure de la saisie](../../images/off15appUXFig04.png)
<br>
<ul><li><p>Un titre explique l’objet du complément et comment commencer (A).</p></li><li><p>La recherche automatique est activée par défaut, avec la possibilité de la désactiver (B).</p></li><li><p>Une fois que vous effectuez une sélection, le complément affiche le contenu correspondant (D).</p></li><li><p>Fournissez un lien pour afficher plus d’informations (E).</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments qui renvoient automatiquement du contenu au fur et à mesure de la création</p></li><li><p>Compléments qui nécessitent que vous sélectionniez du contenu avant l’activation</p></li></ul>

## <a name="locate-content-with-a-task-pane-add-in"></a>Localisation de contenu avec un complément du volet Office


Cet exemple présente un complément du volet Office pour la recherche de contenu.

Dans ce modèle d’interaction, entrez une chaîne dans la zone de recherche ou choisissez parmi une liste de contenus sélectionnés pour commencer.

**Complément du volet Office pour la recherche de contenu**
<br>
![Application de volet de tâches permettant de rechercher du contenu](../../images/off15appUXFig05.png)
<br>
<ul><li><p>La vue initiale contient une zone <span class="ui">Recherche</span> (A) et une liste de contenus sélectionnés (B).</p></li><li><p>Lorsque l’utilisateur entre une chaîne dans la zone de recherche, l’icône de recherche est remplacée par une icône de fermeture (C).</p></li><li><p>Choisissez l’icône Fermer pour revenir à la vue initiale.</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments qui renvoient automatiquement du contenu au fur et à mesure de la création</p></li><li><p>Compléments qui nécessitent que vous sélectionniez du contenu avant l’activation</p></li></ul>

## <a name="insert-media-with-a-task-pane-add-in"></a>Insertion d’un élément multimédia avec un complément du volet Office


Dans ce modèle d’interaction, vous pouvez sélectionner une image à partir des résultats de la recherche pour l’insérer dans le document.

**Complément du volet Office pour l’insertion d’une image**
<br>
![Application de volet de tâches permettant d’insérer une image](../../images/off15appUXFig06.png)
<br>
<ul><li><p>Vous avez filtré la liste des résultats de recherche (A) et sélectionné le contenu à insérer (B).</p></li><li><p>Une vue détaillée du contenu sélectionné est affichée (C) avec un bouton permettant de revenir à la liste.</p></li><li><p>Un bouton <span class="ui">Insérer une photo</span> se trouve dans le pied de page (D). Lorsque vous cliquez sur ce bouton, l’image est insérée dans le document.</p></li><li><p>Une courte description de la provenance de l’image est incluse avec le contenu inséré (E). </p></li><li><p>L’interface utilisateur du complément confirme visuellement la réussite de l’action.</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments permettant d’insérer du contenu</p></li></ul>

## <a name="insert-selected-text-with-a-task-pane-add-in"></a>Insertion du texte sélectionné avec un complément du volet Office


Dans ce modèle d’interaction, vous sélectionnez du texte à partir des résultats de la recherche pour l’insérer dans le document.

**Complément du volet Office pour l’insertion de texte**
<br>
![Application de volet de tâches permettant d’insérer du texte](../../images/off15appUXFig07.png)
<br>
<ul><li><p>Vous avez déjà localisé une portion de contenu (A).</p></li><li><p>Un bouton <span class="ui">Insérer une sélection</span> désactivé est affiché dans le pied de page (B).</p></li><li><p>Lorsque vous sélectionnez une chaîne de texte (C), le bouton <span class="ui">Insérer une sélection</span> devient actif.</p></li><li><p>Une fois que l’utilisateur choisit ce bouton, le complément insère le texte sélectionné dans le document avec une référence à la source du contenu (E).</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments de recherche et d’insertion de contenu</p></li></ul>

## <a name="publish-to-a-web-service-with-a-task-pane-add-in"></a>Publication sur un service web avec un complément du volet Office


Cet exemple présente un complément du volet Office pour la publication d’un document en tant qu’article de blog.

Dans ce modèle d’interaction, vous avez terminé d’écrire le contenu d’un document et vous souhaitez le publier sur votre blog.

**Complément du volet Office pour la publication d’un document en tant qu’article de blog**
<br>
![Application de volet de tâches pour la publication d’un document en tant qu’article de blog](../../images/off15appUXFig08.png)
<br>
<ul><li><p>Tout d’abord, un formulaire de connexion s’affiche pour entrer vos informations d’identification (A).</p></li><li><p>Des liens de création de compte et de gestion des problèmes de connexion classiques sont fournis (B). La sélection de ces liens ouvre les pages correspondantes dans un navigateur.</p></li><li><p>Lorsque vous êtes connecté, le complément affiche un formulaire permettant créer un nouvel article de blog (C).</p></li><li><p>Le nom du compte avec lequel vous vous êtes connecté (et sous lequel vous effectuerez vos publications) apparaît en haut du complément. Le complément fournit un lien pour afficher un aperçu de l’article (D). Sélectionnez-le pour afficher l’aperçu dans un navigateur.</p></li><li><p>Après la sélection de <span class="ui">Créer un article</span>, le complément affiche une vue confirmant que le contenu du document a été publié (E).</p></li><li><p>Le complément fournit un lien permettant d’afficher l’article dans un navigateur (F), ainsi qu’un bouton permettant de créer un autre article (G).</p></li></ul>Convient mieux pour :
<ul><li><p>Compléments qui génèrent, publient ou partagent du contenu sur les réseaux sociaux, les sites de blog et les services web</p></li><li><p>Compléments qui nécessitent que vous vous connectiez à un service</p></li></ul>

## <a name="get-more-information-about-people-with-an-outlook-add-in"></a>Obtention d’informations supplémentaires sur des personnes avec un complément Outlook


 **Exemple 1**

**Page de résultats et de détails**
<br>
![Page de résultats et de détails](../../images/off15appUXFig09.jpg)
<br>
Convient mieux pour :
<ul><li><p>Présentation de l’étendue de votre contenu si vous disposez d’ensembles de données volumineux qu’il serait utile de mettre en avant</p></li><li><p>Pages de détails qui utilisent la taille complète du conteneur de complément</p></li><li><p>Modèles de navigation qui bénéficient d’un flux « aller-retour »</p></li></ul>
 **Exemple 2**

**Page de détails avec navigation persistante**
<br>
![Page de détails avec navigation persistante](../../images/off15appUXFig10.jpg)
<br>
Convient mieux pour :
<ul><li><p>Affichage, par défaut, du premier résultat d’un ensemble de données</p></li><li><p>Structures de navigation fonctionnant comme des onglets (navigation linéaire à un seul niveau)</p></li><li><p>Réduction des actions de sélection en maintenant la navigation disponible en permanence</p></li><li><p>Fourniture d’espace pour afficher la navigation en permanence</p></li></ul>

## <a name="get-more-information-about-content-with-an-outlook-add-in"></a>Obtention d’informations supplémentaires sur le contenu avec un complément Outlook


 **Exemple 1**

**Page de résultats et de détails**
<br>
![Page de résultats et de détails](../../images/off15appUXFig11.jpg)
<br>
Convient mieux pour :
<ul><li><p>Présentation de l’étendue de votre contenu si vous disposez d’ensembles de données volumineux qu’il serait utile d’afficher</p></li><li><p>Sélection ou choix exigé avant l’affichage de détails supplémentaires</p></li><li><p>Pages de détails qui utilisent la taille complète du conteneur de complément</p></li><li><p>Modèles de navigation qui bénéficient d’un flux « aller-retour »</p></li></ul>
 **Exemple 2**

**Page de détails avec contenu secondaire**
<br>
![Page de détails avec contenu secondaire](../../images/off15appUXFig12.jpg)
<br>
Convient mieux pour :
<ul><li><p>Situations dans lesquelles vous souhaitez mettre en avant un élément de contenu</p></li><li><p>Présentation de votre contenu sans interaction de l’utilisateur</p></li><li><p>Navigation persistante (pouvant être ajoutée à ce modèle pour un mélange de simplicité et de facilité de navigation)</p></li></ul>

## <a name="connect-to-an-online-service-and-present-data"></a>Connexion à un service en ligne et présentation des données


Ces exemples illustrent des modèles d’interaction pour l’obtention de données et de contenu à partir d’un service en ligne. Ils peuvent être utilisés dans les trois types de complément : compléments de contenu, compléments du volet Office et compléments Outlook.

 **Exemple 1**

**Carrousel**
<br>
![Carrousel](../../images/off15appUXFig13.jpg)
<br>
Convient mieux pour :
<ul><li><p>Petites quantités de données qui peuvent être exposées individuellement ou par groupe</p></li><li><p>Exposition de contenu sous forme de galerie, comme un diaporama ou une galerie d’images</p></li><li><p>Situations dans lesquelles un modèle de navigation suivant/précédent fonctionne bien</p></li></ul>
 **Exemple 2**

**Assistant**
<br>
![Assistant](../../images/off15appUXFig14.jpg)
<br>
Convient mieux pour :
<ul><li><p>Contenu à présenter dans un ordre spécifique</p></li><li><p>Grandes quantités de contenu adaptées à une consommation sous forme de série de petits éléments</p></li><li><p>Expériences de consommation de type livre</p></li><li><p>Situations dans lesquelles une série d’étapes ou d’actions est nécessaire pour effectuer une tâche</p></li></ul>
 **Exemple 3**

**Formulaire, résultats et détails**
<br>
![Formulaire, résultats et détails](../../images/off15appUXFig15.jpg)
<br>
Convient mieux pour :
<ul><li><p>Compléments qui nécessitent la saisie de données</p></li><li><p>Point d’entrée dans un modèle de résultats et de détails</p></li></ul>

## <a name="additional-resources"></a>Ressources supplémentaires



- [Instructions de conception pour les compléments Office](../add-in-design.md)
    
