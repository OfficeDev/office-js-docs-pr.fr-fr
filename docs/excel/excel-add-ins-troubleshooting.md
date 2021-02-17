---
title: Résolution des problèmes de l’outil de dépannage des add-ins Excel
description: Découvrez comment résoudre les erreurs de développement dans les add-ins Excel.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270727"
---
# <a name="troubleshooting-excel-add-ins"></a>Résolution des problèmes de l’outil de dépannage des add-ins Excel

Cet article traite des problèmes de résolution des problèmes propres à Excel. Utilisez l’outil de commentaires en bas de la page pour suggérer d’autres problèmes qui peuvent être ajoutés à l’article.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Limitations de l’API lorsque le workbook actif bascule

Les add-ins pour Excel sont destinés à fonctionner sur un seul et même workbook à la fois. Des erreurs peuvent survenir lorsqu’un workbook distinct de celui qui exécute le add-in prend le focus. Cela se produit uniquement lorsque des méthodes particulières sont en cours d’appel lorsque le focus change.

Les API suivantes sont affectées par ce commutateur de workbook :

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
> Cela s’applique uniquement à plusieurs workbooks Excel ouverts sur Windows ou Mac.

## <a name="coauthoring"></a>Co-édition

Voir [Co-auteur dans les add-ins Excel](co-authoring-in-excel-add-ins.md) pour les modèles à utiliser avec des événements dans un environnement de co-auteur. L’article traite également des conflits potentiels de fusion lors de l’utilisation de certaines API, telles que [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .

## <a name="known-issues"></a>Problèmes connus

### <a name="binding-events-return-temporary-binding-obects"></a>Les événements de liaison retournent `Binding` desobects temporaires

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) et [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) retournent tous deux un objet temporaire qui contient l’ID de l’objet qui a élevé l’événement. `Binding` `Binding` Utilisez cet ID avec `BindingCollection.getItem(id)` pour récupérer `Binding` l’objet qui a levé l’événement.

L’exemple de code suivant montre comment utiliser cet ID de liaison temporaire pour récupérer l’objet `Binding` associé. Dans l’exemple, un listener d’événement est affecté à une liaison. L’écouteur appelle `getBindingId` la méthode lorsque `onDataChanged` l’événement est déclenché. La `getBindingId` méthode utilise l’ID de l’objet temporaire pour récupérer `Binding` `Binding` l’objet qui a levé l’événement.

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Format des `useStandardHeight` cellules `useStandardWidth` et problèmes

La [propriété useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) de ne fonctionne pas `CellPropertiesFormat` correctement dans Excel sur le web. En raison d’un problème dans l’interface utilisateur d’Excel sur le web, la définition de la propriété pour calculer la hauteur de manière `useStandardHeight` `true` imprécise sur cette plateforme. Par exemple, une hauteur standard de **14** est modifiée à **14,25** dans Excel sur le web.

Sur toutes les plateformes, les propriétés [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) et [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) sont uniquement destinées `CellPropertiesFormat` à être définies sur `true` . La définition de ces `false` propriétés n’a aucun effet. 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Méthode `getImage` Range non pris en cas de non-traitement dans Excel pour Mac

La méthode [Range getImage](/javascript/api/excel/excel.range#getImage__) n’est actuellement pas prise en charge dans Excel pour Mac. Consultez [la #235 OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues/235) pour l’état actuel.

### <a name="range-return-character-limit"></a>Limite de caractères de retour de plage

Les [méthodes Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) et [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) ont une limite de chaîne d’adresses de 8 192 caractères. Lorsque cette limite est dépassée, la chaîne d’adresse est tronquée à 8 192 caractères.

## <a name="see-also"></a>Voir aussi

- [Résoudre les erreurs de développement avec les add-ins Office](../testing/troubleshoot-development-errors.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
