---
title: Gestion des erreurs avec l Excel API JavaScript
description: En savoir plus sur Excel logique de gestion des erreurs de l’API JavaScript pour prendre en compte les erreurs d’utilisation.
ms.date: 09/20/2021
ms.localizationpriority: medium
ms.openlocfilehash: 24daaa8dcd5256be997c8742016a9ec80b3294df
ms.sourcegitcommit: 43f20d0933d0159dd390da052187b315222b185f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/24/2021
ms.locfileid: "59502730"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Gestion des erreurs avec l Excel API JavaScript

Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.

> [!NOTE]
> Pour plus d’informations sur la méthode et la `sync()` nature asynchrone de l’API JavaScript Excel, voir Excel modèle objet [JavaScript](excel-add-ins-core-concepts.md)dans les Office de recherche.

## <a name="best-practices"></a>Meilleures pratiques

Dans l’ensemble des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une instruction `catch` afin de détecter les erreurs qui se produisent au sein de `Excel.run`. Nous vous recommandons d’utiliser le même modèle lorsque vous développez un complément à l’aide des API JavaScript pour Excel.

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a>Erreurs API

Lorsqu’une Excel’API JavaScript échoue, l’API renvoie un objet d’erreur qui contient les propriétés suivantes.

- **code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas traduits.

- **message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.

- **debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.

> [!NOTE]
> si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finaux ne voient pas ces messages d’erreur dans le volet Des tâches du Office application.

## <a name="error-messages"></a>Messages d’erreur

Le tableau suivant contient la liste des erreurs que l’API peut renvoyer.

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`AccessDenied` |Vous ne pouvez pas effectuer l’opération demandée.| |
|`ActivityLimitReached`|La limite d’activité a été atteinte.| |
|`ApiNotAvailable`|L’API demandée n’est pas disponible.| |
|`ApiNotFound`|L’API que vous essayez d’utiliser est in trouver. Il peut être disponible dans une version plus récente de Excel. Pour plus [d’informations, voir Excel’ensembles de](../reference/requirement-sets/excel-api-requirement-sets.md) conditions requises de l’API JavaScript.| |
|`BadPassword`|Le mot de passe que vous avez fourni est incorrect.| |
|`Conflict`|La demande n’a pas pu être traitée en raison d’un conflit.| |
|`ContentLengthRequired`|Un `Content-length` en-tête HTTP est manquant.| |
|`FilteredRangeConflict`|L’opération tentée provoque un conflit avec une plage filtrée.| |
|`FormulaLengthExceedsLimit`|Le bytecode de la formule appliquée dépasse la limite de longueur maximale. Pour Office sur les ordinateurs 32 bits, la limite de longueur du bytecode est de 1 6384 caractères. Sur les ordinateurs 64 bits, la limite de longueur du bytecode est de 32 768 caractères.| Cette erreur se produit à la fois dans Excel sur le Web et sur le bureau.|
|`GeneralException`|Une erreur interne s’est produite lors du traitement de la demande.| |
|`InactiveWorkbook`|L’opération a échoué car plusieurs workbooks sont ouverts et le workbook appelé par cette API a perdu le focus.| |
|`InsertDeleteConflict`|L’opération d’insertion ou de suppression tentée a créé un conflit.| |
|`InvalidArgument` |L’argument est manquant ou non valide, ou a un format incorrect.| |
|`InvalidBinding` |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.| |
|`InvalidOperation`|L’opération tentée n’est pas valide sur l’objet.| |
|`InvalidOperationInCellEditMode`|L’opération n’est pas disponible Excel est en mode Modifier la cellule. Quittez le mode Édition à l’aide des touches **Entrée** ou **Tabulation,** ou en sélectionnant une autre cellule, puis essayez à nouveau.| |
|`InvalidReference`|Cette référence n’est pas valide pour l’opération en cours.| |
|`InvalidRequest`  |Impossible de traiter la demande.| |
|`InvalidSelection`|La sélection en cours est incorrecte pour cette action.| |
|`ItemAlreadyExists`|La ressource en cours de création existe déjà.| |
|`ItemNotFound` |La ressource demandée n’existe pas.| |
|`MemoryLimitReached`|La limite de mémoire a été atteinte. Votre action n’a pas pu être terminée.| |
|`MergedRangeConflict`|Impossible de terminer l’opération. Une table ne peut pas se chevaucher avec un autre tableau, un rapport de tableau croisé dynamique, des résultats de requête, des cellules fusionnées ou une carte XML.|
|`NonBlankCellOffSheet`|Microsoft Excel ne peut pas insérer de nouvelles cellules, car cela pousse les cellules non vides à la fin de la feuille de calcul. Ces cellules non vides peuvent apparaître vides mais ont des valeurs vides, une mise en forme ou une formule. Supprimez suffisamment de lignes ou de colonnes pour faire de la place à ce que vous souhaitez insérer, puis essayez à nouveau.| |
|`NotImplemented`|La fonctionnalité demandée n’est pas implémentée| |
|`OperationCellsExceedLimit`|L’opération tentée affecte plus que la limite de 33554000 cellules.| Si le déclencheur de cette erreur, confirmez qu’il n’y a pas de données involontaires dans la feuille de calcul mais en `TableColumnCollection.add API` dehors du tableau. En particulier, recherchez les données dans les colonnes les plus à droite de la feuille de calcul. Supprimez les données inattendues pour résoudre cette erreur. Une façon de vérifier le nombre de cellules qu’une opération traite consiste à exécuter le calcul suivant `(number of table rows) x (16383 - (number of table columns))` : Le nombre 16383 est le nombre maximal de colonnes que les Excel prend en charge. <br><br>Cette erreur se produit uniquement dans Excel sur le Web. |
|`PivotTableRangeConflict`|L’opération tentée provoque un conflit avec une plage de tableau croisé dynamique.| |
|`RangeExceedsLimit`|Le nombre de cellules dans la plage a dépassé le nombre maximal pris en charge. Pour plus d’informations, voir les limites de ressources et l’optimisation des performances pour Office’article sur les [modules complémentaires.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)| |
|`RefreshWorkbookLinksBlocked`|L’opération a échoué, car l’utilisateur n’a pas accordé l’autorisation d’actualiser les liens debook externes.| |
|`RequestAborted`|La demande a été interrompue pendant l’exécution.| |
|`RequestPayloadSizeLimitExceeded`|La taille de la charge utile de la demande a dépassé la limite. Pour plus d’informations, voir les limites de ressources et l’optimisation des performances pour Office’article sur les [modules complémentaires.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)| Cette erreur se produit uniquement dans Excel sur le Web.|
|`ResponsePayloadSizeLimitExceeded`|La taille de la charge utile de réponse a dépassé la limite. Pour plus d’informations, voir les limites de ressources et l’optimisation des performances pour Office’article sur les [modules complémentaires.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)|  Cette erreur se produit uniquement dans Excel sur le Web.|
|`ServiceNotAvailable`|Le service n’est pas disponible.| |
|`Unauthenticated` |Les informations d’authentification requises sont manquantes ou incorrectes.| |
|`UnsupportedFeature`|L’opération a échoué car la feuille de calcul source contient une ou plusieurs fonctionnalités non pris en compte.| |
|`UnsupportedOperation`|L’opération tentée n’est pas prise en charge.| |
|`UnsupportedSheet`|Ce type de feuille ne prend pas en charge cette opération, car il s’agit d’une feuille Macro ou Graphique.| |

> [!NOTE]
> Le tableau précédent répertorie les messages d’erreur que vous pouvez rencontrer lors de l’utilisation Excel API JavaScript. Si vous travaillez avec l’API commune au lieu de l’API JavaScript Excel spécifique à l’application, voir Office Codes d’erreur [d’API](../reference/javascript-api-for-office-error-codes.md) courants pour en savoir plus sur les messages d’erreur pertinents.

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Codes d'erreur de l'API commune de l'Office](../reference/javascript-api-for-office-error-codes.md)
