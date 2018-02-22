---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 01/29/2018
---

# <a name="work-with-events-using-the-excel-javascript-api"></a>Utilisation d’événements à l’aide de l’API JavaScript pour Excel

Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel. 

> [!IMPORTANT]
> Les API décrites dans cet article sont actuellement disponibles uniquement dans la version d’évaluation publique (bêta) et ne sont pas destinés à être utilisés dans des environnements de production. Pour exécuter les exemples de code contenus dans cet article, vous devez utiliser une version suffisamment récente d’Office et faire référence à la bibliothèque bêta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## <a name="events-in-excel"></a>Événements dans Excel

Chaque fois que certains types de modifications se produisent dans un classeur Excel, une notification d’événement se déclenche. En utilisant l’API JavaScript pour Excel, vous pouvez inscrire les gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit. Les événements suivants sont actuellement pris en charge.

| Événement | Description | Objets pris en charge |
|:---------------|:-------------|:-----------|
| `onAdded` | Événement se produisant lors de l’ajout d’un objet. | **WorksheetCollection** |
| `onDeleted`  | Événement se produisant lorsqu’un objet est supprimé. | **WorksheetCollection** |
| `onActivated` | Événement se produisant lorsqu’un objet est activé. | **WorksheetCollection**, **Worksheet** |
| `onDeactivated` | Événement se produisant lorsqu’un objet est désactivé. | **WorksheetCollection**, **Worksheet** |
| `onDataChanged` | Événement se produisant lorsque les données au sein des cellules sont modifiées. | **Worksheet**, **Table**, **TableCollection**, **Binding** |
| `onSelectionChanged` | Événement se produisant lorsque la cellule active ou la plage sélectionnée est modifiée. | **Worksheet**, **Table**, **Binding** |

### <a name="event-triggers"></a>Déclencheurs d’événements

Événements au sein d’un classeur Excel pouvant être déclenchés par :

- Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur
- Complément (JavaScript) Office modifiant le classeur
- Complément VBA (macro) modifiant le classeur

Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.

### <a name="lifecycle-of-an-event-handler"></a>Cycle de vie d’un gestionnaire d’événements

Un gestionnaire d’événements est créé lorsqu’un complément inscrit le Gestionnaire d’événements et est détruit lorsque le complément désinscrit le Gestionnaire d’événements ou que le complément est fermé. Les gestionnaires d’événements ne sont pas conservées dans le cadre du fichier Excel.

### <a name="events-and-coauthoring"></a>Événements et co-création

Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent collaborer et modifier simultanément le même classeur Excel. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onDataChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Inscription d’un gestionnaire d’événements

L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onDataChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onDataChanged.add(handleDataChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onDataChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Gestion d’un événement

Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console. 

```js
function handleDataChange(event)
{ 
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>Suppression d’un gestionnaire d’événements

L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{ 
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();
        
        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Spécification d’ouverture d’API JavaScript pour Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Présentation des fonctionnalités d’événement Excel (aperçu)](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)