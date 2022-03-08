---
title: Résolution des problèmes Excel des modules
description: Découvrez comment résoudre les erreurs de développement dans Excel des modules.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c6a523354cc938ac9e9ba041ddb09f12142a3a58
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340791"
---
# <a name="troubleshooting-excel-add-ins"></a>Résolution des problèmes Excel des modules

Cet article traite des problèmes de résolution propres aux Excel. Utilisez l’outil de commentaires en bas de la page pour suggérer d’autres problèmes qui peuvent être ajoutés à l’article.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Limitations de l’API lorsque le workbook actif bascule

Les Excel sont destinés à fonctionner sur un seul et même workbook à la fois. Des erreurs peuvent survenir lorsqu’un workbook distinct de celui qui exécute le add-in prend le focus. Cela se produit uniquement lorsque des méthodes particulières sont en cours d’appel lorsque le focus change.

Les API suivantes sont affectées par ce commutateur de workbook.

|sur les API JavaScript pour Excel | Erreur lancée |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> Cela s’applique uniquement à plusieurs Excel de travail ouverts sur Windows mac.

## <a name="coauthoring"></a>Co-édition

Voir [co-Excel des modèles](co-authoring-in-excel-add-ins.md) à utiliser avec des événements dans un environnement de co-auteur. L’article traite également des conflits potentiels de fusion lors de l’utilisation de certaines API, telles que [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)).

## <a name="known-issues"></a>Problèmes connus

### <a name="binding-events-return-temporary-binding-obects"></a>Les événements de liaison retournent desobects `Binding` temporaires

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#excel-excel-bindingdatachangedeventargs-binding-member) et [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#excel-excel-bindingselectionchangedeventargs-binding-member) retournent tous deux un objet temporaire qui contient l’ID `Binding` de l’objet qui a `Binding` élevé l’événement. Utilisez cet ID avec `BindingCollection.getItem(id)` pour récupérer l’objet `Binding` qui a levé l’événement.

L’exemple de code suivant montre comment utiliser cet ID de liaison temporaire pour récupérer l’objet `Binding` associé. Dans l’exemple, un listener d’événement est affecté à une liaison. L’écouteur appelle la `getBindingId` méthode lorsque l’événement `onDataChanged` est déclenché. La `getBindingId` méthode utilise l’ID de l’objet temporaire `Binding` pour `Binding` récupérer l’objet qui a levé l’événement.

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve your binding.
        let binding = context.workbook.bindings.getItemAt(0);
    
        await context.sync();
    
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);
        await context.sync();
    });
}

async function getBindingId(eventArgs) {
    await Excel.run(async (context) => {
        // Get the temporary binding object and load its ID. 
        let tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        let originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Format des `useStandardHeight` cellules et `useStandardWidth` problèmes

La [propriété useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) de `CellPropertiesFormat` ne fonctionne pas correctement dans Excel sur le Web. En raison d’un problème dans l Excel sur le Web’interface utilisateur, `useStandardHeight` `true` la définition de la propriété pour calculer la hauteur de manière imprécise sur cette plateforme. Par exemple, une hauteur standard de **14** est modifiée à **14,25** Excel sur le Web.

Sur toutes les plateformes, les propriétés [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) et [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member) `CellPropertiesFormat` sont uniquement destinées à être définies sur `true`. La définition de ces propriétés n’a `false` aucun effet.

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Méthode Range `getImage` non pris en Excel pour Mac

La méthode [Range getImage](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1)) n’est actuellement pas prise en charge dans Excel pour Mac. [Consultez officeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) pour l’état actuel.

### <a name="range-return-character-limit"></a>Limite de caractères de retour de plage

Les [méthodes Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) et [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1)) ont une limite de chaîne d’adresses de 8 192 caractères. Lorsque cette limite est dépassée, la chaîne d’adresse est tronquée à 8 192 caractères.

## <a name="see-also"></a>Voir aussi

- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/troubleshoot-development-errors.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
