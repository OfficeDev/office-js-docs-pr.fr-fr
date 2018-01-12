# Créer votre premier complément OneNote
<a id="build-your-first-onenote-add-in" class="xliff"></a>

Cet article vous guide tout au long de la procédure de création d’un complément de volet de tâches qui permet d’ajouter du texte à une page OneNote.

L’image suivante présente le complément que vous allez créer.

   ![Complément OneNote généré à partir de cette procédure pas à pas](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## Étape 1 : Configurer votre environnement de développement et créer un projet de complément
<a id="step-1-set-up-your-dev-environment-and-create-an-add-in-project" class="xliff"></a>
Suivez les instructions de la rubrique [Créer un complément Office à l’aide d’un éditeur](../get-started/create-an-office-add-in-using-any-editor.md) pour installer la configuration requise et exécuter le générateur Office Yeoman afin de créer un projet de complément. Le tableau suivant indique les attributs de projet à sélectionner dans le générateur Yeoman.

| Option | Valeur |
|:------|:------|
| Nouveau sous-dossier | (accepter la valeur par défaut) |
| Nom du complément | Complément OneNote |
| Application Office prise en charge | (sélectionner OneNote) |
| Créer un complément | Oui, je souhaite un nouveau complément |
| Ajouter [TypeScript](https://www.typescriptlang.org/) | Non |
| Choisir l’infrastructure | Jquery |

<a name="develop"></a>
## Étape 2 : Modifier le complément
<a id="step-2-modify-the-add-in" class="xliff"></a>
Vous pouvez modifier les fichiers de complément en utilisant un éditeur de texte ou IDE. Si vous n’avez pas encore essayé de Visual Studio Code, vous pouvez le [télécharger gratuitement](https://code.visualstudio.com/) sous Windows, Mac OSX et Linux.

1 - Ouvrez **index.html** dans le répertoire du projet. 

2 - Remplacez l’élément `<main>` par le code suivant. Cette option ajoute une zone de texte et un bouton à l’aide des [composants de la structure de l’interface utilisateur d’Office](http://dev.office.com/fabric/components).

```html
<main class="ms-welcome__main">
   <br />
   <p class="ms-font-l">Enter content below</p>
   <div class="ms-TextField ms-TextField--placeholder">
       <textarea id="textBox" rows="5"></textarea>
   </div>
   <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
        <span class="ms-Button-label">Add Outline</span>
        <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
        <span class="ms-Button-description">Adds the content above to the current page.</span>
    </button>
</main>
```

3 - Ouvrez **app.js** (ou app.ts si vous utilisez TypeScript) dans le répertoire du projet. Modifiez la fonction **Office.initialize** pour ajouter un événement de clic au bouton permettant d’**ajouter un plan**, comme indiqué ci-dessous.

```js
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
   $(document).ready(function () {
       app.initialize();
       
       // Set up event handler for the UI.
       $('#addOutline').click(addOutlineToPage);
   });
};
```
 
4 - Remplacez la méthode **run** par la méthode **addOutlineToPage** suivante. Cela permet d’obtenir le contenu de la zone de texte et de l’ajouter à la page.

```js
// Add the contents of the text area to the page.
function addOutlineToPage() {        
   OneNote.run(function (context) {
      var html = '<p>' + $('#textBox').val() + '</p>';
      
       // Get the current page.
       var page = context.application.getActivePage();
       
       // Queue a command to load the page with the title property.             
       page.load('title'); 
       
       // Add an outline with the specified HTML to the page.
       var outline = page.addOutline(40, 90, html);
       
       // Run the queued commands, and return a promise to indicate task completion.
       return context.sync()
           .then(function() {
               console.log('Added outline to page ' + page.title);
           })
           .catch(function(error) {
               app.showNotification("Error: " + error); 
               console.log("Error: " + error); 
               if (error instanceof OfficeExtension.Error) { 
                   console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
               } 
           }); 
       });
}
```

<a name="test"></a>
## Étape 3 : Test du complément sur OneNote Online
<a id="step-3-test-the-add-in-on-onenote-online" class="xliff"></a>
1 - Démarrez le serveur HTTPS.  

  a. Ouvrez une invite **cmd**/Terminal et accédez au dossier du projet de complément. 
  
  b. Exécutez la commande, comme illustré ci-dessous.

  ```
  C:\your-local-path\onenote add-in\> npm start
  ```

2 - Installez le certificat auto-signé en tant que certificat approuvé. Vous ne devrez effectuer cette opération qu’une seule fois sur votre ordinateur pour l’ensemble des projets de compléments créés avec le générateur Office Yeoman. Pour plus d’informations, reportez-vous à la rubrique relative à l’[ajout de certificats auto-signés en tant que certificats racine approuvés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

3 - Accédez à [OneNote Online](https://www.onenote.com/notebooks) et ouvrez le Bloc-notes.

4 - Sélectionnez **Insérer > Compléments Office**. Cette action ouvre la boîte de dialogue Compléments Office.

  - Si vous êtes connecté avec votre compte de consommateur, sélectionnez l’onglet **MES COMPLÉMENTS**, puis choisissez **Télécharger mon complément**.
  
  - Si vous êtes connecté avec votre compte professionnel ou scolaire, sélectionnez l’onglet **MON ORGANISATION**, puis choisissez **Télécharger mon complément**. 
  
  L’image suivante montre l’onglet **MES COMPLÉMENTS** pour les blocs-notes de consommateurs.

  ![Boîte de dialogue Compléments Office affichant l’onglet MES COMPLÉMENTS](../../images/onenote-office-add-ins-dialog.png)

5 - Dans la boîte de dialogue Télécharger le complément, accédez à **onenote-add-in-manifest.xml** dans le dossier de projet, puis choisissez **Télécharger**. Pendant le test, votre fichier manifeste est stocké dans un espace de stockage local du navigateur.

6 - Le complément s’ouvre dans un iFrame à côté de la page OneNote. Entrez du texte dans la zone de texte, puis choisissez **Ajouter un plan**. Le texte que vous avez entré est ajouté à la page. 

## Conseils et résolution des problèmes
<a id="troubleshooting-and-tips" class="xliff"></a>
- Vous pouvez déboguer le complément à l’aide des outils de développement de votre navigateur. Lorsque vous utilisez le serveur web Gulp et le débogage dans Internet Explorer ou Chrome, vous pouvez enregistrer les modifications localement et simplement actualiser l’iFrame du complément.

- Lorsque vous examinez un objet OneNote, les propriétés qui sont actuellement disponibles affichent les valeurs réelles. Les propriétés qui doivent être chargées sont affichées comme *non définies*. Développez le nœud `_proto_` pour visualiser les propriétés qui sont définies sur l’objet, mais qui ne sont pas encore chargées.

![Objet OneNote déchargé dans le débogueur](../../images/onenote-debug.png)

- Vous devez activer le contenu mixte dans le navigateur si votre complément utilise des ressources HTTP. Les compléments de production doivent uniquement utiliser des ressources HTTPS sécurisées.

- Les compléments de volet Office peuvent être ouverts à partir de n’importe quel emplacement, mais les compléments de contenu peuvent uniquement être insérés à l’intérieur d’un contenu de page normal (et non dans des titres, des images, des iFrames, etc.). 

## Ressources supplémentaires
<a id="additional-resources" class="xliff"></a>

- [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](onenote-add-ins-programming-overview.md)
- [Référence de l’API JavaScript de OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
