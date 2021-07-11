---
title: Utilisation du modèle de l’API propre à l’application
description: Découvrez le modèle d’API basé sur la promesse pour les compléments Excel, OneNote et Word.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5cf1d088dfa883e5df9eaba25e395857cfce9f5c
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350063"
---
# <a name="using-the-application-specific-api-model"></a>Utilisation du modèle de l’API propre à l’application

Cet article décrit l’utilisation du modèle d’API pour la création de compléments dans Excel, Word et OneNote. Il présente les concepts fondamentaux de l’utilisation des API basées sur la promesse.

> [!NOTE]
> Ce modèle n’est pas pris en charge par les clients Office 2013. Utilisez les [Modèles communs de l’API](office-javascript-api-object-model.md) pour fonctionner avec ces versions d’Office. Pour consulter les notes sur la disponibilité complète des plateformes, consultez les [disponibilités de l’application et de la plateforme cliente Office pour les compléments Office](../overview/office-add-in-availability.md).

> [!TIP]
> Les exemples de cette page utilisent les API JavaScript Excel, mais les concepts s’appliquent également aux API JavaScript OneNote, Visio et Word.

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Nature asynchrone des API basées sur la promesse

Les compléments Office sont des sites web qui apparaissent à l’intérieur d’un conteneur de navigateur au sein des applications Office, telles qu’Excel. Ce conteneur est incorporé dans l’application Office sur des plateformes de bureau, telles qu’Office sur Windows, et s’exécute dans un iFrame HTML, dans Office pour le web. En raison de considérations en relation avec les performances, les API Office.js ne peuvent pas interagir de façon synchronisée avec les applications Office sur toutes les plateformes. Par conséquent, l’appel de l’API `sync()` dans Office.js renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est résolue lorsque l’application Excel termine les actions de lecture ou d’écriture demandées. En outre, vous pouvez mettre en file d’attente plusieurs actions, comme la définition des propriétés ou l’appel de méthodes, et les exécuter en tant que lot de commandes avec un seul appel à `sync()`, au lieu d’envoyer une demande distincte pour chaque action. Les sections suivantes décrivent comment effectuer cette tâche à l’aide des API `run()` et `sync()`.

## <a name="run-function"></a>*fonction .run

`Excel.run`, `Word.run`et `OneNote.run` exécutent une fonction qui spécifie les actions à effectuer dans Excel, Word et OneNote. `*.run` crée automatiquement un contexte de demande que vous pouvez utiliser pour interagir avec des objets Office. Lorsque `*.run` a terminé, une promesse est résolue et tous les objets alloués lors de l’exécution sont automatiquement publiés.

L’exemple suivant vous montre comment utiliser `Excel.run`. Le même modèle est également utilisé avec Word et OneNote.

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a>Contexte de demande

L’application Office et votre complément s’exécutent selon deux processus différents. Dans la mesure où ils utilisent différents environnements d’exécution différents, les compléments nécessitent un objet `RequestContext` pour connecter votre complément à des objets dans Office tels que des feuilles de calcul, des plages, des paragraphes et des tableaux. Cet objet `RequestContext` est fourni en tant qu’argument lors de l’appel de `*.run`.

## <a name="proxy-objects"></a>Objets proxy

Les objets JavaScript Office que vous déclarez et utilisez avec les API basées sur la promesse sont des objets proxy. Les méthodes que vous appelez ou les propriétés que vous définissez ou chargez sur les objets proxy sont simplement ajoutées à une file d’attente de commandes en attente. Lorsque vous appelez la méthode `sync()` dans le contexte de la demande (par exemple, `context.sync()`), les commandes en attente sont envoyées à l’application Office et s’exécutent. Ces API sont essentiellement centrées sur les lots. Vous pouvez mettre en file d’attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la méthode `sync()` pour exécuter le lot de commandes mises en file d’attente.

Par exemple, l’extrait de code suivant déclare l’objet JavaScript [Excel.Range](/javascript/api/excel/excel.range), `selectedRange`, pour référencer une plage sélectionnée dans la feuille de calcul Excel, et définit certaines propriétés sur cet objet. L’objet `selectedRange` est un objet proxy. Les propriétés définies et la méthode appelée sur cet objet ne seront pas répercutées dans le document Excel tant que votre complément n’a pas appelé `context.sync()`.

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>Conseil de performance : réduire le nombre d’objets proxy créés

Éviter de créer le même objet proxy à plusieurs reprises. Au lieu de cela, si vous avez besoin du même objet proxy pour plus d’une opération, créez-le une seule fois et affectez-le à une variable, puis utilisez cette variable dans votre code.

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a>sync()

La méthode `sync()` concernant le contexte de demande synchronise l’état entre des objets proxy et des objets dans le document Office. La méthode `sync()` exécute les commandes mises en file d’attente concernant le contexte de demande et récupère des valeurs pour les propriétés qui doivent être chargées dans les objets proxy. La méthode `sync()` est exécutée de façon asynchrone et renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est résolue lorsque la méthode `sync()` est terminée.

L’exemple suivant montre une fonction de traitement par lot qui définit un objet proxy JavaScript local (`selectedRange`), charge une propriété de cet objet et utilise ensuite le modèle de promesses JavaScript pour appeler `context.sync()` afin de synchroniser l’état entre les objets proxy et les objets du document Excel.

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

Dans l’exemple précédent, `selectedRange` est configuré et sa propriété `address` est chargée lorsque `context.sync()` est appelé.

Étant donné que `sync()` est une opération asynchrone, vous devez toujours renvoyer l’objet `Promise` pour vous assurer que l’opération de `sync()` se termine avant que le script continue à s’exécuter. Si vous utilisez TypeScript ou ES6+ JavaScript, vous pouvez `await` l’appel `context.sync()` au lieu de renvoyer la promesse.

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>Conseil de performance : réduire le nombre d’appels de synchronisation

Dans l’API JavaScript Excel, `sync()` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel sur le web. Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez `sync()` et mettre en file d’attente autant de modifications que possible avant d’appeler. Pour plus d’informations sur l’optimisation des performances avec `sync()`, consultez [Évitez d’utiliser la méthode context.sync dans des boucles](../concepts/correlated-objects-pattern.md).

### <a name="load"></a>load()

Pour pouvoir lire les propriétés d’un objet proxy, vous devez charger explicitement les propriétés pour remplir l’objet proxy avec les données du document Office, puis effectuer l’appel `context.sync()`. Par exemple, si vous créez un objet proxy pour référencer une plage sélectionnée, puis que vous voulez lire la propriété `address` de la plage sélectionnée, vous devez charger la propriété `address` avant de la lire. Pour demander le chargement des propriétés d’un objet proxy, appelez la méthode `load()` de l’objet et spécifiez les propriétés à charger. L’exemple suivant illustre la propriété `Range.address` chargée pour `myRange`.

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> Si vous effectuez uniquement un appel de méthodes ou que vous avez des propriétés sur un objet proxy, vous n’avez pas besoin d’effectuer l’appel de la méthode `load()`. La méthode `load()` n’est requise que lorsque vous souhaitez lire les propriétés d’un objet proxy.

À l’instar des demandes de définition de propriétés ou d’appel de méthodes sur des objets proxy, des demandes de chargement de propriétés sur des objets proxy sont ajoutées à la file d’attente des commandes sur le contexte de demande, qui s’exécutera la prochaine fois que vous appellerez la méthode `sync()`. Vous pouvez mettre en file d’attente autant d’appels `load()` sur le contexte de la demande que nécessaire.

#### <a name="scalar-and-navigation-properties"></a>Propriétés scalaires et de navigation

Il existe deux catégories de propriétés: **scalaire** et **de navigation**. Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON. Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont affectés, au lieu d’affecter directement la propriété. Par exemple, les membres `name` et `position` sur l’objet [Excel.Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que `protection` et `tables` sont des propriétés de navigation.

Votre complément peut utiliser des propriétés de navigation comme chemin d’accès pour charger des propriétés scalaires spécifiques. Le code suivant met en file d’attente une commande `load` pour le nom de la police utilisée par un objet `Excel.Range`, sans charger d’autres informations.

```js
someRange.load("format/font/name")
```

Vous pouvez également définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès. Par exemple, vous pouvez définir la taille de police de `Excel.Range` à l’aide de `someRange.format.font.size = 10;`. Vous n’avez pas besoin de charger la propriété avant de la configurer.

N’oubliez pas que certaines des propriétés sous un objet peuvent avoir le même nom qu’un autre objet. Par exemple, `format` est une propriété sous l’objet `Excel.Range`, `format` est également un objet. Donc, si vous effectuez un appel tel que `range.load("format")`, cela équivaut à `range.format.load()` (une instruction `load()` vide indésirable). Pour éviter cela, votre code devrait charger uniquement les nœuds « terminaux » dans une arborescence d’objets.

#### <a name="calling-load-without-parameters-not-recommended"></a>Appel `load` sans paramètres (non recommandé)

Si vous appelez la méthode `load()` sur un objet (ou une collection) sans spécifier de paramètres, toutes les propriétés scalaires de l’objet ou les objets de la collection sont chargées. Le chargement des données inutiles ralentit votre complément. Vous devez toujours spécifier explicitement les propriétés à charger.

> [!IMPORTANT]
> La quantité de données renvoyées par une `load`instruction sans paramètre peut dépasser les limites de taille du service. Pour réduire les risques pesant sur les compléments plus anciens, certaines propriétés ne sont pas renvoyées par `load` sans en faire la demande explicite. Les propriétés suivantes sont exclues de ces opérations de chargement.
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

Les méthodes utilisées dans les API basées sur la promesse qui renvoient des types possèdent un modèle similaire au modèle `load`/`sync`. Par exemple, `Excel.TableCollection.getCount` obtient le nombre de tableaux dans la collection. `getCount` renvoie un `ClientResult<number>`, ce qui signifie que la propriété `value` dans le [`ClientResult`](/javascript/api/office/officeextension.clientresult) renvoyé est un nombre. Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé.

Le code suivant obtient le nombre total de tableaux dans un feuille de calcul Excel et enregistre ce nombre dans la console.

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a>set()

La définition de propriétés sur un objet avec des propriétés de navigation imbriquées peut être laborieuse. Au lieu de définir des propriétés individuelles à l’aide de chemins de navigation comme décrit ci-dessus, vous pouvez utiliser la méthode `object.set()` disponible sur les objets dans les API JavaScript basées sur une promesse. Grâce à cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant soit un autre objet du même type Office.js, soit un objet JavaScript avec des propriétés structurées comme celles de l’objet sur lequel la méthode est appelée.

L’exemple de code suivant définit plusieurs propriétés de mise en forme d’une plage en appelant la méthode `set()` et en transmettant un objet JavaScript avec des noms et des types de propriétés reflétant la structure des propriétés dans l’objet `Range`. Cet exemple part du principe que des données sont présentes dans la plage **B2:E2**.

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="some-properties-cannot-be-set-directly"></a>Certaines propriétés ne peuvent pas être définies directement

Certaines propriétés ne peuvent pas être définies, même si elles sont accessibles en écriture. Ces propriétés font partie d’une propriété parente qui doit être définie en tant qu’objet unique. En effet, cette propriété parente s’appuie sur les sous-propriétés ayant des relations logiques spécifiques. Ces propriétés parentes doivent être définies à l’aide de la notation littérale de l’objet pour définir l’intégralité de l’objet, plutôt que de définir les sous-propriétés individuelles de cet objet. Un exemple de ce modèle est trouvé dans [PageLayout](/javascript/api/excel/excel.pagelayout). La propriété `zoom` doit être définie avec un objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) unique, comme illustré ici :

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

Dans l’exemple précédent, vous ***ne pouvez pas*** affecter directement une valeur à `zoom` : `sheet.pageLayout.zoom.scale = 200;`. Cette instruction génère une erreur, car `zoom` n’est pas chargé. Même si `zoom` était chargé, l’ensemble d’échelles n’est pas pris en compte. Toutes les opérations de contexte se produisent sur `zoom`, elles actualisent l’objet proxy du complément et remplacement des valeurs définies localement.

Ce comportement diffère des [propriétés de navigation](application-specific-api-model.md#scalar-and-navigation-properties) telles que [Range.format](/javascript/api/excel/excel.range#format). Les propriétés de `format` peuvent être définies à l’aide de la navigation d’objets, comme illustré ici :

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Vous pouvez identifier une propriété qui ne peut pas avoir ses sous-propriétés définies directement en consultant son modificateur en lecture seule. Toutes les propriétés en lecture seule peuvent avoir leurs sous-propriétés sans lecture seule directement définit. Les propriétés disponibles en écriture comme `PageLayout.zoom`, par exemple, doivent être définies avec un objet de ce niveau. En Résumé :

- Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.
- Propriété accessibles en écriture : les sous-propriétés ne peuvent pas être définies via la navigation (elles doivent être définies dans le cadre de l’affectation d’objet parent initiale).



## <a name="42ornullobject-methods-and-properties"></a>Méthodes et propriétés de &#42;OrNullObject

Certaines méthodes et propriétés d’accessoires ajoutent une exception lorsque l’objet souhaité n’existe pas. Par exemple, si vous tentez d’obtenir une feuille de calcul Excel en spécifiant le nom d’une feuille de calcul qui n’existe pas dans le classeur, la méthode `getItem()` renvoie une exception `ItemNotFound`. Les bibliothèques spécifiques de l’application permettent à votre code de tester l’existence d’entités de document sans exiger de code de gestion d’exceptions. Cela est possible à l’aide des variantes `*OrNullObject` de méthodes et de propriétés. Ces variantes renvoient un objet dont la propriété `isNullObject` est définie sur `true`, si l’élément spécifié n’existe pas, plutôt que de renvoyer une exception.

Par exemple, vous pouvez appeler la méthode `getItemOrNullObject()` sur une collection telle que **Feuilles de calcul** pour récupérer un élément de la collection. La méthode `getItemOrNullObject()` renvoie l’élément spécifié s’il existe. sinon, il renvoie un objet dont la propriété `isNullObject` est définie sur `true`. Votre code peut ensuite évaluer cette propriété pour déterminer si l’objet existe.

> [!NOTE]
> Les variantes `*OrNullObject` ne renvoient jamais la valeur JavaScript `null`. Ils renvoient des objets proxy Office ordinaires. Si l’entité que l’objet représente n’existe pas, la propriété `isNullObject` de l’objet est définie sur `true`. Ne testez pas l’objet renvoyé pour nullité ou fausseté. Ce n’est jamais `null`, `false` ou `undefined`.

L’exemple de code suivant tente de récupérer une feuille de calcul Excel nommée « Données » à l’aide de la méthode `getItemOrNullObject()`. Si une feuille de calcul avec ce nom n’existe pas, une nouvelle feuille est créée. Notez que le code ne charge pas la propriété `isNullObject`. Office charge automatiquement cette propriété lorsque `context.sync` est appelé. Vous n’avez donc pas besoin de la charger explicitement avec quelque chose comme `datasheet.load('isNullObject')`.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a>Voir aussi

* [Modèle d’objet API JavaScript courant](office-javascript-api-object-model.md)
* [Limites des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md)
