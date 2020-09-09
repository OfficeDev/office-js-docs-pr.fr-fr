---
title: Résolution des problèmes liés aux compléments Excel
description: Découvrez comment résoudre les problèmes liés aux erreurs de développement dans les compléments Excel.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409388"
---
# <a name="troubleshooting-excel-add-ins"></a>Résolution des problèmes liés aux compléments Excel

Cet article traite de la résolution des problèmes propres à Excel. Veuillez utiliser l’outil de commentaires en bas de la page pour suggérer d’autres problèmes pouvant être ajoutés à l’article.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Limitations de l’API lorsque le classeur actif bascule

Les compléments pour Excel sont conçus pour fonctionner sur un seul classeur à la fois. Des erreurs peuvent se produire lorsqu’un classeur distinct de celui qui exécute le complément obtient le focus. Cela se produit uniquement lorsque des méthodes particulières sont en cours d’appel lorsque le focus est modifié.

Les API suivantes sont affectées par ce commutateur de classeurs :

|sur les API JavaScript pour Excel | Erreur générée |
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
> Cela s’applique uniquement à plusieurs classeurs Excel ouverts sous Windows ou Mac.

## <a name="coauthoring"></a>Co-édition

Consultez la rubrique [co-authoring in Excel Add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with Events in a CoAuthoring Environment. L’article aborde également les conflits de fusion potentiels lors de l’utilisation de certaines API, telles que [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .

## <a name="see-also"></a>Voir aussi

- [Résoudre les erreurs de développement avec les compléments Office](../testing/troubleshoot-development-errors.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
