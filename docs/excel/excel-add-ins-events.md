---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: Liste d’événements pour les objets JavaScript Excel. Cela inclut des informations sur l’utilisation des gestionnaires d’événements et les modèles associés.
ms.date: 08/18/2020
localization_priority: Normal
ms.openlocfilehash: adb924ff370c49f5b8f6a3175683ecdb959d9329
ms.sourcegitcommit: 7faa0932b953a4983a80af70f49d116c3236d81a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/21/2020
ms.locfileid: "46845520"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Utilisation d’événements à l’aide de l’API JavaScript pour Excel

Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.

## <a name="events-in-excel"></a>Événements dans Excel

Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Événement | Description | Objets pris en charge |
|:---------------|:-------------|:-----------|
| `onActivated` | Se produit lorsqu’un objet est activé. | [**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated) |
| `onAdded` | Se produit lorsqu’un objet est ajouté à la collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded) |
| `onAutoSaveSettingChanged` | Se produit lorsque le paramètre de `autoSave` est modifié dans le classeur. | [**Classeur**](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | Se produit lorsqu’une feuille de calcul a terminé un calcul (ou que toutes les feuilles de calcul de la collection ont terminé). | [**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated) |
| `onChanged` | Se produit lorsque les données au sein des cellules sont modifiées. | [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged) |
| `onColumnSorted` | Se produit lorsqu’une ou plusieurs colonnes ont été triées. Ce problème se produit en raison de l’opération de tri de gauche à droite. | [**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted) |
| `onDataChanged` | Se produit lors de la modification des données ou de la mise en forme dans la liaison. | [**Binding**](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | Se produit lorsqu’un objet est désactivé. | [**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated) |
| `onDeleted` | Se produit lorsqu’un objet est supprimé de la collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted) |
| `onFormatChanged` | Se produit lorsque le format est modifié sur une feuille de calcul. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged) |
| `onRowSorted` | Se produit lorsqu’une ou plusieurs lignes ont été triées. Cela se produit en raison d’une opération de tri de haut en bas. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted) |
| `onSelectionChanged` | Se produit lorsque la cellule active ou la plage sélectionnée est modifiée. | [**Liaison**](/javascript/api/excel/excel.binding#onselectionchanged), [**table**](/javascript/api/excel/excel.table#onselectionchanged), [**classeur**](/javascript/api/excel/excel.workbook#onselectionchanged), [**feuille de calcul**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged) |
| `onRowHiddenChanged` | Se produit lorsque l’état de ligne masquée change sur une feuille de calcul spécifique. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged) |
| `onSettingsChanged` | Se produit lorsque les paramètres dans le document sont modifiés. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | Se produit lorsque l’opération clic gauche/tape se produit dans la feuille de calcul. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked) |

### <a name="events-in-preview"></a>Événements en préversion

> [!NOTE]
> Les événements suivants sont actuellement disponibles uniquement en préversion publique. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Événement | Description | Objets pris en charge |
|:---------------|:-------------|:-----------|
| `onFiltered` | Se produit lorsqu’un filtre est appliqué à un objet. | [**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered) |

### <a name="event-triggers"></a>Déclencheurs d’événements

Événements au sein d’un classeur Excel pouvant être déclenchés par :

- Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur
- Complément (JavaScript) Office modifiant le classeur
- Complément VBA (macro) modifiant le classeur

Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.

### <a name="lifecycle-of-an-event-handler"></a>Cycle de vie d’un gestionnaire d’événements

Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements. Il est détruit lorsque le complément annule l’inscription du gestionnaire d’événements ou lorsque le complément est actualisé, rechargé ou fermé. Les gestionnaires d’événements ne sont pas conservés dans le fichier Excel ou entre des sessions avec Excel sur le web.

> [!CAUTION]
> Lorsqu’un objet dans lequel des événements sont inscrits est supprimé (par exemple, un tableau avec un événement `onChanged`), le gestionnaire d’événements n’est plus déclenché mais reste en mémoire jusqu’à ce que le complément ou la session Excel soit actualisé(e) ou se ferme.

### <a name="events-and-coauthoring"></a>Événements et co-création

Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Inscription d’un gestionnaire d’événements

L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.

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

## <a name="handle-an-event"></a>Gestion d’un événement

Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.

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

## <a name="remove-an-event-handler"></a>Suppression d’un gestionnaire d’événements

L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements. Notez que le `RequestContext` utilisé pour créer le gestionnaire d’événements est nécessaire pour le supprimer. 

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

## <a name="enable-and-disable-events"></a>Activation et désactivation d’événements

La performance d’un complément peut être améliorée en désactivant les événements.
Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.

Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).
La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.

L’exemple de code suivant montre comment activer et désactiver des événements.

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
