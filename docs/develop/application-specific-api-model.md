---
title: Utilisation du modèle d’API propre à l’application
description: Découvrez le modèle d’API basée sur la promesse pour les compléments Excel, OneNote et Word.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 0a5068312b8b17f7ceeafcffd5dcea4203314ebf
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294032"
---
# <a name="using-the-application-specific-api-model"></a>Utilisation du modèle d’API propre à l’application

Cet article explique comment utiliser le modèle d’API pour créer des compléments dans Excel, Word et OneNote. Il présente les concepts fondamentaux de l’utilisation des API basées sur les promesses.

> [!NOTE]
> Ce modèle n’est pas pris en charge par les clients Office 2013. Utilisez le [modèle d’API commun](office-javascript-api-object-model.md) pour utiliser ces versions d’Office. Pour obtenir des notes sur la disponibilité complète de la plateforme, consultez la rubrique [Office client Application and Platform Availability for Office Add-ins](../overview/office-add-in-availability.md).

> [!TIP]
> Les exemples de cette page utilisent les API JavaScript Excel, mais les concepts s’appliquent également aux API JavaScript OneNote, Visio et Word.

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Nature asynchrone des API basées sur les promesses

Les compléments Office sont des sites Web qui s’affichent dans un conteneur de navigateur dans les applications Office, telles qu’Excel. Ce conteneur est incorporé dans l’application Office sur les plateformes de bureau, comme Office sur Windows, et s’exécute dans un iFrame HTML dans Office sur le Web. Pour des raisons de performances, les API Office.js ne peuvent pas interagir de façon synchrone avec les applications Office sur toutes les plateformes. Par conséquent, l' `sync()` appel de l’API dans Office.js renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est résolue lorsque l’application Office termine les actions de lecture ou d’écriture demandées. En outre, vous pouvez mettre en file d’attente plusieurs actions, telles que définir des propriétés ou appeler des méthodes, et les exécuter sous la forme d’un lot de commandes avec un seul appel à `sync()` , au lieu d’envoyer une demande distincte pour chaque action. Les sections suivantes décrivent comment effectuer cette procédure à l’aide des `run()` `sync()` API et.

## <a name="run-function"></a>fonction *. Run

`Excel.run`, `Word.run` et `OneNote.run` exécutent une fonction qui spécifie les actions à effectuer sur Excel, Word et OneNote. `*.run` crée automatiquement un contexte de demande que vous pouvez utiliser pour interagir avec les objets Office. `*.run`Une fois l’opération terminée, une promesse est résolue et tous les objets qui ont été alloués lors de l’exécution sont automatiquement publiés.

L’exemple suivant montre comment utiliser `Excel.run` . Le même modèle est également utilisé avec Word et OneNote.

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

L’application Office et votre complément s’exécutent dans deux processus différents. Dans la mesure où ils utilisent des environnements d’exécution différents, les compléments nécessitent un `RequestContext` objet pour connecter votre complément à des objets dans Office, tels que des feuilles de calcul, des plages, des paragraphes et des tableaux. Cet `RequestContext` objet est fourni en tant qu’argument lors de l’appel `*.run` .

## <a name="proxy-objects"></a>Objets de proxy

Les objets JavaScript Office que vous déclarez et utilisez avec les API basées sur les promesses sont des objets proxy. Les méthodes que vous appelez ou les propriétés que vous définissez ou chargez sur les objets proxy sont simplement ajoutées à une file d’attente de commandes en attente. Lorsque vous appelez la `sync()` méthode sur le contexte de la demande (par exemple, `context.sync()` ), les commandes en file d’attente sont envoyées vers l’application Office et exécutées. Ces API sont fondamentalement centrées sur les lots. Vous pouvez mettre en file d’attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la `sync()` méthode pour exécuter le lot de commandes en file d’attente.

Par exemple, l’extrait de code suivant déclare l’objet [Excel. Range](/javascript/api/excel/excel.range) JavaScript local, `selectedRange` , pour faire référence à une plage sélectionnée dans le classeur Excel, puis définit certaines propriétés sur cet objet. L' `selectedRange` objet est un objet proxy, de sorte que les propriétés qui sont définies et la méthode appelée sur cet objet ne sont pas reflétées dans le document Excel tant que votre complément n’a pas été appelé `context.sync()` .

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>Conseil de performance : réduisez le nombre d’objets proxy créés

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

### <a name="sync"></a>Sync()

L’appel de la `sync()` méthode sur le contexte de demande synchronise l’état entre les objets proxy et les objets dans le document Office. La `sync()` méthode exécute toutes les commandes qui sont placées en file d’attente dans le contexte de la demande et récupère des valeurs pour les propriétés qui doivent être chargées sur les objets proxy. La `sync()` méthode s’exécute de façon asynchrone et renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est résolue à la fin de la `sync()` méthode.

L’exemple suivant montre une fonction batch qui définit un objet proxy JavaScript local ( `selectedRange` ), charge une propriété de cet objet, puis utilise le modèle de promet JavaScript pour appeler `context.sync()` pour synchroniser l’état entre les objets proxy et les objets dans le document Excel.

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

Étant donné qu’il `sync()` s’agit d’une opération asynchrone, vous devez toujours retourner l' `Promise` objet pour vous assurer que l' `sync()` opération se termine avant que le script continue à s’exécuter. Si vous utilisez la machine à écrire ou ES6 + JavaScript, vous pouvez `await` `context.sync()` appeler au lieu de renvoyer la promesse.

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>Conseil de performance : réduisez le nombre d’appels de synchronisation

Dans l’API JavaScript Excel, `sync()` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel sur le web. Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez `sync()` et mettre en file d’attente autant de modifications que possible avant d’appeler. Pour plus d’informations sur l’optimisation des performances avec `sync()` , reportez-vous à [la rubrique éviter d’utiliser la méthode Context. Sync dans les boucles](../concepts/correlated-objects-pattern.md).

### <a name="load"></a>load()

Avant de pouvoir lire les propriétés d’un objet proxy, vous devez charger explicitement les propriétés pour remplir l’objet proxy avec les données du document Office, puis appeler `context.sync()` . Par exemple, si vous créez un objet proxy pour référencer une plage sélectionnée, puis que vous souhaitez lire la propriété de la plage sélectionnée `address` , vous devez charger la `address` propriété avant de pouvoir la lire. Pour demander le chargement des propriétés d’un objet proxy, appelez la `load()` méthode sur l’objet et spécifiez les propriétés à charger. L’exemple suivant montre la `Range.address` propriété en cours de chargement `myRange` .

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
> Si vous appelez uniquement des méthodes ou définissez des propriétés sur un objet proxy, il n’est pas nécessaire d’appeler la `load()` méthode. La `load()` méthode n’est requise que si vous souhaitez lire les propriétés sur un objet proxy.

À l’instar des demandes de définition de propriétés ou d’appel de méthodes sur des objets proxy, des demandes de chargement de propriétés sur des objets proxy sont ajoutées à la file d’attente des commandes sur le contexte de demande, qui s’exécutera la prochaine fois que vous appellerez la méthode `sync()`. Vous pouvez mettre en file d’attente autant d’appels `load()` sur le contexte de la demande que nécessaire.

#### <a name="scalar-and-navigation-properties"></a>Propriétés scalaires et de navigation

Il existe deux catégories de propriétés: **scalaire** et **de navigation**. Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON. Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont affectés au lieu d’affecter directement la propriété. Par exemple, `name` les `position` membres de l’objet [Excel. Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que les `protection` Propriétés de `tables` navigation.

Votre complément peut utiliser les propriétés de navigation comme chemin d’accès pour charger des propriétés scalaires spécifiques. Le code suivant met en file d’attente une `load` commande pour le nom de la police utilisée par un `Excel.Range` objet, sans charger aucune autre information.

```js
someRange.load("format/font/name")
```

Vous pouvez également définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès. Par exemple, vous pouvez définir la taille de la police pour un `Excel.Range` à l’aide de `someRange.format.font.size = 10;` . Vous n’avez pas besoin de charger la propriété avant de la définir.

N’oubliez pas que certaines propriétés sous un objet peuvent avoir le même nom qu’un autre objet. Par exemple, `format` est une propriété sous l' `Excel.Range` objet, mais `format` elle est également un objet. Par conséquent, si vous effectuez un appel tel que `range.load("format")` , cela équivaut à `range.format.load()` (une instruction indésirable vide `load()` ). Pour éviter cela, votre code doit uniquement charger les « nœuds feuille » dans une arborescence d’objets.

#### <a name="calling-load-without-parameters-not-recommended"></a>Appel `load` sans paramètres (non recommandé)

Si vous appelez la `load()` méthode sur un objet (ou une collection) sans spécifier de paramètres, toutes les propriétés scalaires de l’objet ou des objets de la collection sont chargées. Le chargement des données inutiles ralentira votre complément. Vous devez toujours spécifier explicitement les propriétés à charger.

> [!IMPORTANT]
> La quantité de données renvoyées par une `load`instruction sans paramètre peut dépasser les limites de taille du service. Pour réduire les risques pesant sur les compléments plus anciens, certaines propriétés ne sont pas renvoyées par `load` sans en faire la demande explicite. Les propriétés suivantes sont exclues des opérations de chargement suivantes :
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

Les méthodes dans les API basées sur la promesse qui retournent des types primitifs ont un modèle similaire pour le `load` / `sync` paradigme. Par exemple, `Excel.TableCollection.getCount` obtient le nombre de tableaux dans la collection. `getCount` renvoie un `ClientResult<number>` , ce qui signifie que la `value` propriété dans le renvoyé [`ClientResult`](/javascript/api/office/officeextension.clientresult) est un nombre. Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé.

Le code suivant obtient le nombre total de tables dans un classeur Excel et enregistre ce nombre dans la console.

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

La définition de propriétés sur un objet avec des propriétés de navigation imbriquées peut être laborieuse. En guise d’alternative à la définition de propriétés individuelles à l’aide de chemins de navigation, comme décrit ci-dessus, vous pouvez utiliser la `object.set()` méthode qui est disponible sur les objets dans les API JavaScript à promesse. Grâce à cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant soit un autre objet du même type Office.js, soit un objet JavaScript avec des propriétés structurées comme celles de l’objet sur lequel la méthode est appelée.

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

## <a name="42ornullobject-methods-and-properties"></a>&#42;des méthodes et des propriétés de OrNullObject

Certaines propriétés et méthodes d’accesseur génèrent une exception lorsque l’objet souhaité n’existe pas. Par exemple, si vous tentez d’obtenir une feuille de calcul Excel en spécifiant un nom de feuille de calcul qui ne se trouve pas dans le classeur, la `getItem()` méthode génère une `ItemNotFound` exception.

Tout `*OrNullObject` Variant vous permet de vérifier un objet sans lever d’exceptions. Ces méthodes et propriétés renvoient un objet null (et non JavaScript `null` ) au lieu de lever une exception si l’élément spécifié n’existe pas. Par exemple, vous pouvez appeler la `getItemOrNullObject()` méthode sur une collection telle que **Worksheets** pour récupérer un élément de la collection. La méthode `getItemOrNullObject()` renvoie l’élément spécifié, s’il existe. Sinon, elle renvoie un objet Null. L’objet Null renvoyé contient la propriété booléenne `isNullObject` que vous pouvez étudier pour déterminer l’existence de l’objet.

L’exemple de code suivant tente de récupérer une feuille de calcul Excel nommée « Data » à l’aide de la `getItemOrNullObject()` méthode. Si la méthode renvoie un objet null, une nouvelle feuille est créée avant que les actions de la feuille ne soient effectuées.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        // If `dataSheet` is a null object, create the worksheet.
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a>Voir aussi

* [Modèle d’objet d’API JavaScript courant](office-javascript-api-object-model.md)
* [Problèmes courants liés au code et comportements de plateforme inattendus](/common-coding-issues.md).
* [Limites des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md)
