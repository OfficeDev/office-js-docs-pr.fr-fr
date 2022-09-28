---
title: Gestion des erreurs avec les API JavaScript spécifiques à l’application
description: Découvrez Excel, Word, PowerPoint et une autre logique de gestion des erreurs de l’API JavaScript spécifique à l’application pour prendre en compte les erreurs d’exécution.
ms.date: 07/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: b6f25f5740892df4729b72ee5ad87403853f45fb
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092993"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>Gestion des erreurs avec les API JavaScript spécifiques à l’application

Lorsque vous générez un complément à l’aide [des API JavaScript Office spécifiques à l’application, veillez](../develop/application-specific-api-model.md) à inclure une logique de gestion des erreurs pour tenir compte des erreurs d’exécution. Cette opération est essentielle en raison de la nature asynchrone des API.

## <a name="best-practices"></a>Bonnes pratiques

Dans nos [exemples de code](https://github.com/OfficeDev/Office-Add-in-samples) et [Script Lab](../overview/explore-with-script-lab.md) extraits de code, vous remarquerez que chaque appel à `Excel.run`, `PowerPoint.run`ou `Word.run` est accompagné d’une `catch` instruction pour intercepter les erreurs. Nous vous recommandons d’utiliser le même modèle lorsque vous générez un complément à l’aide des API spécifiques à l’application.

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add your Excel JavaScript API calls here.

      // Await the completion of context.sync() before continuing.
    await context.sync();
    console.log("Finished!");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

```

## <a name="api-errors"></a>Erreurs API

Lorsqu’une demande d’API JavaScript Office ne s’exécute pas correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes.

- **code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas traduits.

- **message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.

- **debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.

> [!NOTE]
> Si vous utilisez cette option `console.log()` pour imprimer des messages d’erreur dans la console, ces messages sont visibles uniquement sur le serveur. Les utilisateurs finaux ne voient pas ces messages d’erreur dans le volet Office du complément ou n’importe où dans l’application Office. Pour signaler des erreurs à l’utilisateur, consultez [notifications d’erreur](#error-notifications).

## <a name="error-codes-and-messages"></a>Codes d’erreur et messages

Les tableaux suivants répertorient les erreurs que les API spécifiques à l’application peuvent retourner.

> [!NOTE]
> Le tableau précédent répertorie les messages d’erreur que vous pouvez rencontrer lors de l’utilisation des API spécifiques à l’application. Si vous utilisez l’API Commune, consultez [les codes d’erreur de l’API Commune Office](../reference/javascript-api-for-office-error-codes.md) pour en savoir plus sur les messages d’erreur pertinents.

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`AccessDenied` |Vous ne pouvez pas effectuer l’opération demandée.|*Aucun.* |
|`ActivityLimitReached`|La limite d’activité a été atteinte.|*Aucun.* |
|`ApiNotAvailable`|L’API demandée n’est pas disponible.|*Aucun.* |
|`ApiNotFound`|L’API que vous essayez d’utiliser est introuvable. Il peut être disponible dans une version plus récente d’Excel. Pour plus d’informations, consultez l’article sur les [ensembles de conditions requises de l’API JavaScript Excel](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) .|*Aucun.* |
|`BadPassword`|Le mot de passe que vous avez fourni est incorrect.|*Aucun.* |
|`Conflict`|La demande n’a pas pu être traitée en raison d’un conflit.|*Aucun.* |
|`ContentLengthRequired`|Un `Content-length` en-tête HTTP est manquant.|*Aucun.* |
|`GeneralException`|Une erreur interne s’est produite lors du traitement de la demande.|*Aucun.* |
|`InsertDeleteConflict`|L’opération d’insertion ou de suppression tentée a créé un conflit.|*Aucun.* |
|`InvalidArgument` |L’argument est manquant ou non valide, ou a un format incorrect.|*Aucun.* |
|`InvalidBinding` |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.|*Aucun.* |
|`InvalidOperation`|L’opération tentée n’est pas valide sur l’objet.|*Aucun.* |
|`InvalidOperationInCellEditMode`|L’opération n’est pas disponible tant qu’Excel est en mode Modifier la cellule. Quittez le mode Édition à l’aide des touches **Entrée** ou **Tab** , ou en sélectionnant une autre cellule, puis réessayez.|*Aucun.* |
|`InvalidReference`|Cette référence n’est pas valide pour l’opération en cours.|*Aucun.* |
|`InvalidRequest`  |Impossible de traiter la demande.|*Aucun.* |
|`InvalidSelection`|La sélection en cours est incorrecte pour cette action.|*Aucun.* |
|`ItemAlreadyExists`|La ressource en cours de création existe déjà.|*Aucun.* |
|`ItemNotFound` |La ressource demandée n’existe pas.|*Aucun.* |
|`MemoryLimitReached`|La limite de mémoire a été atteinte. Votre action n’a pas pu être effectuée.|*Aucun.* |
|`NotImplemented`|La fonctionnalité demandée n’est pas implémentée| Cela peut signifier que l’API est en préversion ou uniquement prise en charge sur une plateforme particulière (par exemple, en ligne uniquement). Pour plus d’informations, consultez [la disponibilité des applications clientes et de la plateforme Office pour les compléments Office](/javascript/api/requirement-sets) .|
|`RequestAborted`|La demande a été interrompue pendant l’exécution.|*Aucun.* |
|`RequestPayloadSizeLimitExceeded`|La taille de la charge utile de la demande a dépassé la limite. Pour plus d’informations, consultez l’article Sur [les limites de ressources et l’optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .| Cette erreur se produit uniquement dans Office sur le Web.|
|`ResponsePayloadSizeLimitExceeded`|La taille de la charge utile de la réponse a dépassé la limite. Pour plus d’informations, consultez l’article Sur [les limites de ressources et l’optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .|  Cette erreur se produit uniquement dans Office sur le Web.|
|`ServiceNotAvailable`|Le service n’est pas disponible.|*Aucun.* |
|`Unauthenticated` |Les informations d’authentification requises sont manquantes ou incorrectes.|*Aucun.* |
|`UnsupportedFeature`|L’opération a échoué, car la feuille de calcul source contient une ou plusieurs fonctionnalités non prises en charge.|*Aucun.* |
|`UnsupportedOperation`|L’opération tentée n’est pas prise en charge.|*Aucun.* |

### <a name="excel-specific-error-codes-and-messages"></a>Codes d’erreur et messages spécifiques à Excel

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`EmptyChartSeries`|L’opération tentée a échoué, car la série de graphiques est vide.|*Aucun.* |
|`FilteredRangeConflict`|La tentative d’opération provoque un conflit avec une plage filtrée.|*Aucun.* |
|`FormulaLengthExceedsLimit`|L’octet de la formule appliquée dépasse la limite de longueur maximale. Pour Office sur les machines 32 bits, la limite de longueur d’octet est de 16 384 caractères. Sur les machines 64 bits, la longueur d’octet est de 32 768 caractères.| Cette erreur se produit à la fois dans Excel sur le Web et sur le bureau.|
|`InactiveWorkbook`|L’opération a échoué, car plusieurs classeurs sont ouverts et le classeur appelé par cette API a perdu le focus.|*Aucun.* |
|`MergedRangeConflict`|Impossible de terminer l’opération. Une table ne peut pas chevaucher une autre table, un rapport de tableau croisé dynamique, des résultats de requête, des cellules fusionnées ou une carte XML.|*Aucun.* |
|`NonBlankCellOffSheet`|Microsoft Excel ne peut pas insérer de nouvelles cellules, car il pousserait les cellules non vides à la fin de la feuille de calcul. Ces cellules non vides peuvent apparaître vides, mais ont des valeurs vides, une certaine mise en forme ou une formule. Supprimez suffisamment de lignes ou de colonnes pour faire place à ce que vous voulez insérer, puis réessayez.|*Aucun.* |
|`OperationCellsExceedLimit`|L’opération tentée affecte plus que la limite de 33554000 cellules.| Si l’erreur `TableColumnCollection.add API` se déclenche, vérifiez qu’il n’y a pas de données involontaires dans la feuille de calcul, mais en dehors de la table. En particulier, recherchez les données dans les colonnes les plus à droite de la feuille de calcul. Supprimez les données inattendues pour résoudre cette erreur. Une façon de vérifier le nombre de cellules qu’une opération traite consiste à exécuter le calcul suivant : `(number of table rows) x (16383 - (number of table columns))`. Le nombre 16383 correspond au nombre maximal de colonnes prises en charge par Excel. <br><br>Cette erreur se produit uniquement dans Excel sur le Web. |
|`PivotTableRangeConflict`|La tentative d’opération provoque un conflit avec une plage de tableaux croisés dynamiques.|*Aucun.* |
|`RangeExceedsLimit`|Le nombre de cellules dans la plage a dépassé le nombre maximal pris en charge. Pour plus d’informations, consultez l’article Sur [les limites de ressources et l’optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .|*Aucun.* |
|`RefreshWorkbookLinksBlocked`|L’opération a échoué, car l’utilisateur n’a pas accordé l’autorisation d’actualiser les liens de classeur externes.|*Aucun.* |
|`UnsupportedSheet`|Ce type de feuille ne prend pas en charge cette opération, car il s’agit d’une feuille macro ou graphique.|*Aucun.* |

## <a name="error-notifications"></a>Notifications d’erreur

La façon dont vous signalez des erreurs aux utilisateurs dépend du système d’interface utilisateur que vous utilisez. Si vous utilisez React comme système d’interface utilisateur, utilisez les composants de l’interface utilisateur Fluent et les éléments de conception. Choisissez un contrôle approprié dans cette [page Fluent UI](https://developer.microsoft.com/fluentui#/controls/web). Nous vous recommandons de transmettre les messages d’erreur à l’aide d’une barre de messages, d’une boîte de dialogue ou d’un modal. Si l’erreur se trouve dans l’entrée de l’utilisateur, affichez l’erreur en rouge gras près du contrôle d’entrée. L’exemple [office-complément-microsoft-graphe-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) utilise un élément MessageBar et le modifie pour prendre en compte le menu personnalité dans un volet office de complément.

Si vous n’utilisez pas React pour l’interface utilisateur, envisagez d’utiliser les anciens composants de l’interface utilisateur Fabric implémentés directement en HTML et JavaScript. Certains exemples de modèles se trouvent dans le référentiel [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) . Jetez un coup d’œil en particulier dans les sous-dossiers de dialogue et de navigation. L’exemple [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) utilise une bannière de message.

## <a name="see-also"></a>Voir aussi

- [Objet OfficeExtension.Error (API JavaScript pour Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Codes d'erreur de l'API commune de l'Office](../reference/javascript-api-for-office-error-codes.md)
