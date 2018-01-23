# <a name="powerpoint-add-ins"></a>Compléments PowerPoint

Vous pouvez utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations de vos utilisateurs sur toutes les plateformes, y compris Windows, iOS, Office Online et Mac. Vous pouvez créer l’un des deux types de compléments :

- Utilisez des **compléments de contenu** pour ajouter du contenu HTML5 dynamique à vos présentations. Par exemple, consultez le complément [Diagrammes LucidChart pour PowerPoint](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), qui vous permet d’injecter un diagramme interactif de LucidChart dans votre support de présentation.
- Utilisez des **compléments de volet Office** pour faire apparaître des informations de référence ou insérer des données dans la diapositive via un service. Par exemple, consultez le complément [Images Shutterstock](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), qui vous permet d’ajouter des photos professionnelles à votre présentation. 

>
  **Remarque :** Lorsque vous créez votre complément, si vous envisagez de le [publier](../publish/publish.md) dans Office Store, assurez-vous que vous respectez les [stratégies de validation Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability)).

## <a name="powerpoint-add-in-scenarios"></a>Scénarios de complément PowerPoint

Les exemples de code figurant dans l’article vous présentent certaines tâches de base en matière de développement de compléments de contenu pour PowerPoint. 

Pour afficher des informations, ces exemples dépendent de la fonction `app.showNotification`, qui est incluse dans les modèles de projet de compléments Office Visual Studio. Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devrez remplacer la fonction `showNotification` par votre propre code. Plusieurs de ces exemples dépendent également de l’objet `globals` qui est déclaré en dehors de la portée de ces fonctions : `var globals = {activeViewHandler:0, firstSlideId:0};`

Pour obtenir ces exemples de code, votre projet doit faire référence à la [bibliothèque Office.js v1.1 ou version ultérieure](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged

Si vous créez un complément de contenu, vous devrez obtenir la vue active de la présentation et gérer l’événement ActiveViewChanged dans le cadre de votre gestionnaire Office.Initialize.


- La fonction  `getActiveFileView` appelle la méthode [Document.getActiveViewAsync](http://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Mode Plan**) ou « lecture » ( **Diaporama** ou **Mode Lecture**).


- La fonction `registerActiveViewChanged` appelle la méthode [addHandlerAsync](http://dev.office.com/reference/add-ins/shared/document.addhandlerasync) afin d’inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged). 
> Remarque : Dans PowerPoint Online, l’événement [Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged) ne se déclenche jamais, car le mode diaporama est considéré comme une nouvelle session. Dans ce cas, le complément doit extraire la vue active lors du chargement, comme indiqué ci-dessous.



```js

//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}


function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
           app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>Accéder à une diapositive spécifique dans la présentation

La fonction  `getSelectedRange` appelle la méthode [Document.getSelectedDataAsync](http://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) pour obtenir un objet JSON renvoyé par `asyncResult.value` et qui contient un tableau intitulé « diapositives » répertoriant les ID, les titres et les index de la série de diapositives sélectionnée (ou uniquement de la diapositive en cours). Elle enregistre également l’ID de la première diapositive de la série sélectionnée dans une variable globale.


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

La fonction  `goToFirstSlide` appelle la méthode [Document.goToByIdAsync](http://dev.office.com/reference/add-ins/shared/document.gotobyidasync) pour accéder à l’ID de la première diapositive stockée par la fonction `getSelectedRange` ci-dessus.




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## <a name="navigate-between-slides-in-the-presentation"></a>Naviguer entre les diapositives de la présentation

La fonction `goToSlideByIndex` appelle la méthode **Document.goToByIdAsync** pour passer à la diapositive suivante dans la présentation.


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>Obtenir l’URL de la présentation

La fonction `getFileUrl` appelle la méthode [Document.getFileProperties](http://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) pour obtenir l’URL du fichier de présentation.


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```



## <a name="additional-resources"></a>Ressources supplémentaires
- [Exemples de code PowerPoint](https://dev.office.com/code-samples#?filters=powerpoint)

- [Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Utiliser des thèmes de document dans vos compléments PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
