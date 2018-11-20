# <a name="word-javascript-api-overview"></a>Présentation des API JavaScript pour Word

Word propose un ensemble complet d’API que vous pouvez utiliser pour créer des compléments qui interagissent avec les métadonnées et le contenu du document. Ces API permettent de créer des expériences attrayantes qui s’intègrent à Word et l’étendent. Vous pouvez importer et exporter du contenu, assembler de nouveaux documents provenant de différentes sources de données et réaliser une intégration avec des flux de travail de document pour créer des solutions de document personnalisées.

Vous pouvez utiliser deux API JavaScript pour interagir avec les objets et les métadonnées d’un document Word :

- API JavaScript pour Word : introduite dans Office 2016.
- [Interface API JavaScript pour Office](../javascript-api-for-office.md) (Office.js) : introduite dans Office 2013.

## <a name="word-javascript-api"></a>API JavaScript pour Word

L’API JavaScript pour Word est chargée par Office.js. Elle offre une nouvelle façon d’interagir avec les objets tels que les documents et les paragraphes. Ainsi, vous n’utilisez plus d’API asynchrones individuelles pour extraire et mettre à jour chacun de ces objets. L’API JavaScript pour Word fournit des objets JavaScript de « proxy » qui correspondent aux objets réels utilisés dans Word. Vous pouvez interagir avec ces objets de proxy en lisant et en écrivant leurs propriétés de façon synchronisée, et en appelant des méthodes synchrones pour effectuer des opérations les concernant. Ces interactions avec les objets de proxy ne sont pas immédiatement appliquées dans le script en cours d’exécution. La méthode **context.sync** synchronise l’état de vos objets JavaScript en cours d’exécution et celui des objets réels en exécutant des instructions mises en file d’attente et en récupérant les propriétés des objets Word chargés pour les utiliser dans votre script.

## <a name="javascript-api-for-office"></a>Interface API JavaScript pour Office

Vous pouvez référencer Office.js à partir des emplacements suivants :

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js : utilisez cette ressource pour les compléments de production.
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js : utilisez cette ressource quand vous essayez les fonctionnalités d’aperçu.

Si vous utilisez [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), vous pouvez télécharger les [outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx) pour obtenir des modèles de projets qui incluent Office.js.  Vous pouvez également utiliser [nuget pour obtenir Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).

Si vous utilisez TypeScript et que vous avez npm, vous pouvez obtenir les définitions TypeScript en tapant `typings install office-js --ambient` dans votre interface de ligne de commande.

## <a name="running-word-add-ins"></a>Exécution de compléments Word

Pour exécuter votre complément, utilisez un gestionnaire d’événements Office.initialize. Pour plus d’informations sur l’initialisation du complément, reportez-vous à la section [Présentation de l’API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

Les compléments qui ciblent Word 2016 ou version ultérieure s’exécutent en transmettant une fonction dans la méthode **Word.run()**. La fonction transmise dans la méthode **run** doit contenir un argument de contexte. Cet [objet de contexte](/javascript/api/word/word.requestcontext) est différent de celui que vous obtenez de l’objet Office, même s’il sert également à interagir avec l’environnement d’exécution de Word. L’objet de contexte permet d’accéder au modèle objet de l’API JavaScript pour Word. L’exemple suivant montre comment initialiser et exécuter un complément Word à l’aide de la méthode **Word.run()**.

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Synchronisation de documents Word avec des objets de proxy de l’API JavaScript pour Word

Le modèle objet de l’API JavaScript pour Word est associé de façon relativement libre aux objets dans Word. Les objets de l’API JavaScript pour Word sont des proxys pour des objets dans un document Word. Les actions effectuées sur les objets de proxy ne sont pas réalisées dans Word tant que l’état du document n’a pas été synchronisé. Inversement, l’état du document Word n’est pas répercuté sur les objets de proxy tant que l’état du document n’a pas été synchronisé. Pour synchroniser l’état du document, vous exécutez la méthode **context.sync()**. L’exemple suivant présente la création d’un objet Body de proxy et une file de commandes permettant de charger la propriété de texte sur l’objet Body de proxy, puis la synchronisation du corps dans le document Word avec l’objet de proxy correspondant à l’aide de la méthode **context.sync()**.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>Exécution d’un lot de commandes

Les objets de proxy Word utilisent des méthodes pour accéder au modèle objet et le mettre à jour. Ces méthodes sont exécutées l’une après l’autre, dans l’ordre dans lequel elles ont incluses dans la file d’attente du lot. Toutes les commandes en attente dans le lot sont exécutées lorsque la méthode context.sync() est appelée.

L’exemple suivant montre comment fonctionne la file d’attente de commandes. Lorsque la méthode **context.sync()** est appelée, la commande visant à charger le corps du texte est exécutée dans Word. C’est ensuite la commande visant à insérer du texte dans le corps de Word qui est appliquée. Les résultats sont alors renvoyés vers l’objet Body de proxy. La valeur de la propriété **body.text** dans l’API JavaScript pour Word est la valeur du corps du document de Word <u>avant</u> l’insertion du texte dans le document Word.


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a>Spécifications ouvertes de l’API JavaScript pour Word

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Word, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](../openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline pour les API JavaScript pour Word et donnez votre avis sur nos spécifications de conception.

## <a name="word-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Word

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour Word, consultez l’article [Ensembles de conditions requises de l’API JavaScript pour Word](../requirement-sets/word-api-requirement-sets.md).

## <a name="word-javascript-api-reference"></a>Référence d’API JavaScript pour Word

Pour en savoir plus sur l’API JavaScript pour Word, consultez la [documentation de référence de l’API JavaScript pour Word](/javascript/api/word).

## <a name="see-also"></a>Voir aussi

* [Présentation des compléments Word](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [Vue d’ensemble de la plateforme des compléments Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [Exemples de compléments Word sur GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
