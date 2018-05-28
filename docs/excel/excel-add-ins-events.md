---
title: Utilisation d??v?nements ? l?aide de l?API JavaScript pour Excel
description: ''
ms.date: 01/29/2018
ms.openlocfilehash: 4e04b31e7a130f21d6a9c94d041dc2a122a5890e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Utilisation d??v?nements ? l?aide de l?API JavaScript pour Excel 

Cet article d?crit des concepts importants relatifs ? l?utilisation des ?v?nements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d??v?nements, g?rer des ?v?nements et supprimer des gestionnaires d??v?nements ? l?aide de l?API JavaScript pour Excel. 

> [!IMPORTANT]
> Les API d?crites dans cet article sont actuellement disponibles uniquement dans la version d??valuation publique (b?ta) et ne sont pas destin?es ? ?tre utilis?es dans des environnements de production. Pour ex?cuter les exemples de code contenus dans cet article, vous devez utiliser une version suffisamment r?cente d?Office et faire r?f?rence ? la biblioth?que b?ta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## <a name="events-in-excel"></a>?v?nements dans Excel

Chaque fois que certains types de modifications se produisent dans un classeur Excel, une notification d??v?nement se d?clenche. En utilisant l?API JavaScript pour Excel, vous pouvez inscrire les gestionnaires d??v?nements autorisant votre compl?ment ? ex?cuter automatiquement une fonction d?sign?e lorsqu?un ?v?nement sp?cifique se produit. Les ?v?nements suivants sont actuellement pris en charge.

| ?v?nement | Description | Objets pris en charge |
|:---------------|:-------------|:-----------|
| `onAdded` | ?v?nement se produisant lors de l?ajout d?un objet. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | ?v?nement se produisant lorsqu?un objet est activ?. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md) |
| `onDeactivated` | ?v?nement se produisant lorsqu?un objet est d?sactiv?. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md) |
| `onChanged` | ?v?nement se produisant lorsque les donn?es au sein des cellules sont modifi?es. | [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**Liaison**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md) |
| `onSelectionChanged` | ?v?nement se produisant lorsque la cellule active ou la plage s?lectionn?e est modifi?e. | [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md), [**Liaison**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md) |

### <a name="event-triggers"></a>D?clencheurs d??v?nements

?v?nements au sein d?un classeur Excel pouvant ?tre d?clench?s par :

- Interaction de l?utilisateur via l?interface utilisateur Excel (IU) modifiant le classeur
- Compl?ment (JavaScript) Office modifiant le classeur
- Compl?ment VBA (macro) modifiant le classeur

Toute modification conforme aux comportements par d?faut d?Excel d?clenche les ?v?nements correspondants dans un classeur.

### <a name="lifecycle-of-an-event-handler"></a>Cycle de vie d?un gestionnaire d??v?nements

Un gestionnaire d??v?nements est cr?? lorsqu?un compl?ment inscrit le gestionnaire d??v?nements et est d?truit lorsque le compl?ment d?sinscrit le gestionnaire d??v?nements ou que le compl?ment est ferm?. Les gestionnaires d??v?nements ne persistent pas en tant que partie du fichier Excel.

### <a name="events-and-coauthoring"></a>?v?nements et co-cr?ation

Avec la [co-cr?ation](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le m?me classeur Excel simultan?ment. Pour les ?v?nements pouvant ?tre d?clench?s par un co-auteur, tels que `onChanged`, l?objet **Event** correspondant contient une propri?t? **source** qui indique si l??v?nement a ?t? d?clench? localement par l?utilisateur actuel (`event.source = Local`) ou par le co-auteur ? distance (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Inscription d?un gestionnaire d??v?nements

L?exemple de code suivant inscrit un gestionnaire d??v?nements pour l??v?nement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit ?tre ex?cut?e lorsque les donn?es de la feuille de calcul sont modifi?es.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Gestion d?un ?v?nement

Comme indiqu? dans l?exemple pr?c?dent, lorsque vous inscrivez un gestionnaire d??v?nements, vous indiquez la fonction devant ?tre ex?cut?e lorsque l??v?nement sp?cifi? se produit. Vous pouvez cr?er cette fonction pour effectuer n?importe quelle action n?cessaire ? votre sc?nario. L?exemple de code suivant montre une fonction de gestionnaire d??v?nements qui ?crit simplement des informations sur l??v?nement dans la console. 

```js
function handleChange(event)
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

## <a name="remove-an-event-handler"></a>Suppression d?un gestionnaire d??v?nements

L?exemple de code suivant inscrit un gestionnaire d??v?nements pour l??v?nement `onSelectionChanged` dans la feuille de calcul **Sample** et d?finit la fonction `handleSelectionChange` qui est ex?cut?e lorsqu?un ?v?nement se produit. Il d?finit ?galement la fonction `remove()` pouvant ?tre appel?e par la suite pour supprimer ce gestionnaire d??v?nements.

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

- [Concepts de base de l?API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Sp?cification d?ouverture d?API JavaScript pour Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Pr?sentation des fonctionnalit?s d??v?nement Excel (aper?u)](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)
