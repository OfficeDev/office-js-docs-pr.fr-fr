---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: Liste des événements pour Excel objets JavaScript. Cela inclut des informations sur l’utilisation des handlers d’événements et les modèles associés.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8bc1dcad8bccb51dbcedfee741954fabf6967670
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340581"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Utilisation d’événements à l’aide de l’API JavaScript pour Excel

Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.

## <a name="events-in-excel"></a>Événements dans Excel

Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Événement | Description | Objets pris en charge |
|:---------------|:-------------|:-----------|
| `onActivated` | Se produit lorsqu’un objet est activé. | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member) |
| `onActivated` | Se produit lorsqu’un workbook est activé. | [**Classeur**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) |
| `onAdded` | Se produit lorsqu’un objet est ajouté à la collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member) |
| `onAutoSaveSettingChanged` | Se produit lorsque le paramètre de `autoSave` est modifié dans le classeur. | [**Classeur**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member) |
| `onCalculated` | Se produit lorsqu’une feuille de calcul a terminé un calcul (ou que toutes les feuilles de calcul de la collection ont terminé). | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member) |
| `onChanged` | Se produit lorsque les données de cellules individuelles ou de commentaires ont changé. | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member) |
| `onColumnSorted` | Se produit lorsqu’une ou plusieurs colonnes ont été triées. Ce problème se produit en raison de l’opération de tri de gauche à droite. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member) |
| `onDataChanged` | Se produit lors de la modification des données ou de la mise en forme dans la liaison. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-ondatachanged-member) |
| `onDeactivated` | Se produit lorsqu’un objet est désactivé. | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member) |
| `onDeleted` | Se produit lorsqu’un objet est supprimé de la collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member) |
| `onFormatChanged` | Se produit lorsque le format est modifié sur une feuille de calcul. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member) |
| `onFormulaChanged` | Se produit lorsqu’une formule est modifiée. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member) |
| `onProtectionChanged` | Se produit lorsque l’état de protection de la feuille de calcul est modifié. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member) |
| `onRowHiddenChanged` | Se produit lorsque l’état de ligne masquée change sur une feuille de calcul spécifique. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member) |
| `onRowSorted` | Se produit lorsqu’une ou plusieurs lignes ont été triées. Cela se produit en raison d’une opération de tri de haut en bas. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member) |
| `onSelectionChanged` | Se produit lorsque la cellule active ou la plage sélectionnée est modifiée. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-onselectionchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member), [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member) |
| `onSettingsChanged` | Se produit lorsque les paramètres dans le document sont modifiés. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member) |
| `onSingleClicked` | Se produit lorsque l’opération clic gauche/tape se produit dans la feuille de calcul. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member) |

### <a name="events-in-preview"></a>Événements en préversion

> [!NOTE]
> Les événements suivants sont actuellement disponibles uniquement en préversion publique. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Événement | Description | Objets pris en charge |
|:---------------|:-------------|:-----------|
| `onFiltered` | Se produit lorsqu’un filtre est appliqué à un objet. | [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member) |

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

L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.

```js
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the worksheet.");
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Gestion d’un événement

Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.

```js
async function handleChange(event) {
    await Excel.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>Suppression d’un gestionnaire d’événements

L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements. Notez que l’utilisation `RequestContext` pour créer le handler d’événements est nécessaire pour le supprimer.

```js
let eventResult;

async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    await context.sync();
    console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
  });
}

async function handleSelectionChange(event) {
  await Excel.run(async (context) => {
    await context.sync();
    console.log("Address of current selection: " + event.address);
  });
}

async function remove() {
  await Excel.run(eventResult.context, async (context) => {
    eventResult.remove();
    await context.sync();
    
    eventResult = null;
    console.log("Event handler successfully removed.");
  });
}
```

## <a name="enable-and-disable-events"></a>Activation et désactivation d’événements

La performance d’un complément peut être améliorée en désactivant les événements.
Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.

Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).
La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.

L’exemple de code suivant montre comment activer et désactiver des événements.

```js
await Excel.run(async (context) => {
    context.runtime.load("enableEvents");
    await context.sync();

    let eventBoolean = !context.runtime.enableEvents;
    context.runtime.enableEvents = eventBoolean;
    if (eventBoolean) {
        console.log("Events are currently on.");
    } else {
        console.log("Events are currently off.");
    }
    
    await context.sync();
});
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
