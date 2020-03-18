---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Word
description: L’API JavaScript pour Word permet de créer des compléments pour Word.
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: bbd9fa23ffd5e25555f2d0d5e0022ebc2e81a534
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719643"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Concepts fondamentaux de programmation avec l’API JavaScript pour Word

Cet article décrit les concepts de base de l’utilisation de [l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md) pour créer des compléments pour Word 2016 ou version ultérieure.

## <a name="referencing-officejs"></a>Referencing Office.js

Vous pouvez référencer Office.js à partir des emplacements suivants :

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` : utilisez cette ressource pour les compléments de production.

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` : utilisez cette ressource pour essayer les fonctionnalités en préversion.

## <a name="word-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Word

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour Word, voir [Ensembles de conditions requises de l’API JavaScript pour Word](../reference/requirement-sets/word-api-requirement-sets.md).

## <a name="running-word-add-ins"></a>Exécution de compléments Word

Pour exécuter votre complément, utilisez un gestionnaire d’événements `Office.initialize`. Pour plus d’informations sur l’initialisation du complément, voir [Présentation de l’API](../develop/understanding-the-javascript-api-for-office.md).

Les compléments qui ciblent Word 2016 ou version ultérieure s’exécutent en transmettant une fonction dans la méthode `Word.run()`. La fonction transmise dans la méthode `run` doit contenir un argument de contexte. Cet [objet de contexte](/javascript/api/word/word.requestcontext) est différent de celui que vous obtenez de l’objet Office, même s’il sert également à interagir avec l’environnement d’exécution de Word. L’objet de contexte permet d’accéder au modèle objet de l’API JavaScript pour Word. L’exemple suivant montre comment initialiser et exécuter un complément Word à l’aide de la méthode `Word.run()`.

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

### <a name="asynchronous-nature-of-word-apis"></a>Nature asynchrone des API pour Word

L’API JavaScript pour Word est chargée par Office.js. L’API JavaScript pour Word change la façon d’interagir avec des objets tels que des documents et des paragraphes. Ainsi, vous ne devez plus utiliser d’API asynchrones individuelles pour extraire et mettre à jour chacun de ces objets. L’API JavaScript pour Word fournit des objets JavaScript « proxy » qui correspondent aux objets en direct s’exécutant dans Word. Vous pouvez interagir avec ces objets proxy en lisant et écrivant leurs propriétés de façon synchronisée, et en appelant des méthodes synchrones pour effectuer des opérations sur ces objets. Ces interactions avec des objets proxy n’ont pas lieu immédiatement dans le script en cours d’exécution. La méthode `context.sync` synchronise l’état de vos objets JavaScript en cours d’exécution et celui des objets réels en exécutant des instructions en file d’attente et en récupérant des propriétés d’objets Word chargés pour les utiliser dans votre script.

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Synchronisation de documents Word avec des objets de proxy de l’API JavaScript pour Word

Le modèle objet de l’API JavaScript pour Word est associé de façon relativement libre aux objets dans Word. Les objets de l’API JavaScript pour Word sont des proxys pour des objets dans un document Word. Les actions effectuées sur les objets de proxy ne sont pas réalisées dans Word tant que l’état du document n’a pas été synchronisé. Inversement, l’état du document Word n’est pas répercuté sur les objets de proxy tant que l’état du document n’a pas été synchronisé. Pour synchroniser l’état du document, vous exécutez la méthode `context.sync()`. L’exemple suivant présente la création d’un objet Body de proxy et une file de commandes permettant de charger la propriété de texte sur l’objet Body de proxy, puis la synchronisation du corps dans le document Word avec l’objet de proxy correspondant à l’aide de la méthode `context.sync()`.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>Exécution d’un lot de commandes

Les objets de proxy Word utilisent des méthodes pour accéder au modèle objet et le mettre à jour. Ces méthodes sont exécutées séquentiellement, dans l’ordre de leur mise en file d’attente dans le lot. Toutes les commandes en file d’attente dans le lot sont exécutées lors de l’appel de la méthode `context.sync()`.

L’exemple suivant montre comment fonctionne la file d’attente de commandes. Lors de l’appel de la méthode `context.sync()`, la commande visant à charger le corps du texte est exécutée dans Word. C’est ensuite la commande visant à insérer du texte dans le corps de Word qui est appliquée. Les résultats sont alors renvoyés vers l’objet Body de proxy. La valeur de la propriété `body.text` dans le code JavaScript Word est la même que celle du corps du document de Word <u>avant</u> l’insertion du texte dans le document Word.

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

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

## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md)
- [Créer votre premier complément Word](../quickstarts/word-quickstart.md)
- [Didacticiel sur les compléments Word](../tutorials/word-tutorial.md)
- [Référence d’API JavaScript pour Word](/javascript/api/word)