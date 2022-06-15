---
title: Gestion des erreurs avec l’API JavaScript Excel
description: Découvrez Excel logique de gestion des erreurs de l’API JavaScript pour prendre en compte les erreurs d’exécution.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6fa5ca0c7ebf9400fcdd83c7bf4eb4b906f2e5b5
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090830"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Gestion des erreurs avec l’API JavaScript Excel

Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.

> [!NOTE]
> Pour plus d’informations sur la `sync()` méthode et la nature asynchrone de Excel’API JavaScript, consultez [Excel modèle objet JavaScript dans Office compléments](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Meilleures pratiques

Dans nos [exemples de code](https://github.com/OfficeDev/Office-Add-in-samples) et [Script Lab](../overview/explore-with-script-lab.md) extraits de code, vous remarquerez que chaque appel `Excel.run` est accompagné d’une `catch` instruction pour intercepter les erreurs qui se produisent dans le `Excel.run`. Nous vous recommandons d’utiliser le même modèle lorsque vous développez un complément à l’aide des API JavaScript pour Excel.

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

Lorsqu’une demande d’API JavaScript Excel ne s’exécute pas correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes.

- **code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas traduits.

- **message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.

- **debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.

> [!NOTE]
> Si vous utilisez cette option `console.log()` pour imprimer des messages d’erreur dans la console, ces messages sont visibles uniquement sur le serveur. Les utilisateurs finaux ne voient pas ces messages d’erreur dans le volet Office du complément ou n’importe où dans l’application Office. Pour signaler des erreurs à l’utilisateur, consultez [notifications d’erreur](#error-notifications).

## <a name="error-messages"></a>Messages d’erreur

Le tableau suivant contient la liste des erreurs que l’API peut renvoyer.

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`AccessDenied` |Vous ne pouvez pas effectuer l’opération demandée.| |
|`ActivityLimitReached`|La limite d’activité a été atteinte.| |
|`ApiNotAvailable`|L’API demandée n’est pas disponible.| |
|`ApiNotFound`|L’API que vous essayez d’utiliser est introuvable. Il peut être disponible dans une version plus récente de Excel. Pour plus d’informations, consultez l’article Excel ensembles de [conditions requises de l’API JavaScript](/javascript/api/requirement-sets/excel/excel-api-requirement-sets).| |
|`BadPassword`|Le mot de passe que vous avez fourni est incorrect.| |
|`Conflict`|La demande n’a pas pu être traitée en raison d’un conflit.| |
|`ContentLengthRequired`|Un `Content-length` en-tête HTTP est manquant.| |
|`EmptyChartSeries`|L’opération tentée a échoué, car la série de graphiques est vide.| |
|`FilteredRangeConflict`|La tentative d’opération provoque un conflit avec une plage filtrée.| |
|`FormulaLengthExceedsLimit`|L’octet de la formule appliquée dépasse la limite de longueur maximale. Pour Office sur les machines 32 bits, la limite de longueur d’octet est de 1 6384 caractères. Sur les machines 64 bits, la longueur d’octet est de 32 768 caractères.| Cette erreur se produit à la fois dans Excel sur le Web et sur le bureau.|
|`GeneralException`|Une erreur interne s’est produite lors du traitement de la demande.| |
|`InactiveWorkbook`|L’opération a échoué, car plusieurs classeurs sont ouverts et le classeur appelé par cette API a perdu le focus.| |
|`InsertDeleteConflict`|L’opération d’insertion ou de suppression tentée a créé un conflit.| |
|`InvalidArgument` |L’argument est manquant ou non valide, ou a un format incorrect.| |
|`InvalidBinding` |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.| |
|`InvalidOperation`|L’opération tentée n’est pas valide sur l’objet.| |
|`InvalidOperationInCellEditMode`|L’opération n’est pas disponible tant que Excel est en mode Modifier la cellule. Quittez le mode Édition à l’aide des touches **Entrée** ou **Tab** , ou en sélectionnant une autre cellule, puis réessayez.| |
|`InvalidReference`|Cette référence n’est pas valide pour l’opération en cours.| |
|`InvalidRequest`  |Impossible de traiter la demande.| |
|`InvalidSelection`|La sélection en cours est incorrecte pour cette action.| |
|`ItemAlreadyExists`|La ressource en cours de création existe déjà.| |
|`ItemNotFound` |La ressource demandée n’existe pas.| |
|`MemoryLimitReached`|La limite de mémoire a été atteinte. Votre action n’a pas pu être effectuée.| |
|`MergedRangeConflict`|Impossible de terminer l’opération. Une table ne peut pas chevaucher une autre table, un rapport de tableau croisé dynamique, des résultats de requête, des cellules fusionnées ou une carte XML.|
|`NonBlankCellOffSheet`|Microsoft Excel ne peut pas insérer de nouvelles cellules, car cela pousserait des cellules non vides à la fin de la feuille de calcul. Ces cellules non vides peuvent apparaître vides, mais ont des valeurs vides, une certaine mise en forme ou une formule. Supprimez suffisamment de lignes ou de colonnes pour faire place à ce que vous voulez insérer, puis réessayez.| |
|`NotImplemented`|La fonctionnalité demandée n’est pas implémentée| |
|`OperationCellsExceedLimit`|L’opération tentée affecte plus que la limite de 33554000 cellules.| Si l’erreur `TableColumnCollection.add API` se déclenche, vérifiez qu’il n’y a pas de données involontaires dans la feuille de calcul, mais en dehors de la table. En particulier, recherchez les données dans les colonnes les plus à droite de la feuille de calcul. Supprimez les données inattendues pour résoudre cette erreur. Une façon de vérifier le nombre de cellules qu’une opération traite consiste à exécuter le calcul suivant : `(number of table rows) x (16383 - (number of table columns))`. Le nombre 16383 est le nombre maximal de colonnes prises en charge par Excel. <br><br>Cette erreur se produit uniquement dans Excel sur le Web. |
|`PivotTableRangeConflict`|La tentative d’opération provoque un conflit avec une plage de tableaux croisés dynamiques.| |
|`RangeExceedsLimit`|Le nombre de cellules dans la plage a dépassé le nombre maximal pris en charge. Pour plus d’informations, consultez l’article sur [les limites de ressources et l’optimisation des performances pour Office compléments](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).| |
|`RefreshWorkbookLinksBlocked`|L’opération a échoué, car l’utilisateur n’a pas accordé l’autorisation d’actualiser les liens de classeur externes.| |
|`RequestAborted`|La demande a été interrompue pendant l’exécution.| |
|`RequestPayloadSizeLimitExceeded`|La taille de la charge utile de la demande a dépassé la limite. Pour plus d’informations, consultez l’article sur [les limites de ressources et l’optimisation des performances pour Office compléments](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).| Cette erreur se produit uniquement dans Excel sur le Web.|
|`ResponsePayloadSizeLimitExceeded`|La taille de la charge utile de la réponse a dépassé la limite. Pour plus d’informations, consultez l’article sur [les limites de ressources et l’optimisation des performances pour Office compléments](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).|  Cette erreur se produit uniquement dans Excel sur le Web.|
|`ServiceNotAvailable`|Le service n’est pas disponible.| |
|`Unauthenticated` |Les informations d’authentification requises sont manquantes ou incorrectes.| |
|`UnsupportedFeature`|L’opération a échoué, car la feuille de calcul source contient une ou plusieurs fonctionnalités non prises en charge.| |
|`UnsupportedOperation`|L’opération tentée n’est pas prise en charge.| |
|`UnsupportedSheet`|Ce type de feuille ne prend pas en charge cette opération, car il s’agit d’une feuille macro ou graphique.| |

> [!NOTE]
> Le tableau précédent répertorie les messages d’erreur que vous pouvez rencontrer lors de l’utilisation de l’API JavaScript Excel. Si vous utilisez l’API Common au lieu de l’API JavaScript spécifique à l’application Excel, consultez [Office codes d’erreur de l’API commune](../reference/javascript-api-for-office-error-codes.md) pour en savoir plus sur les messages d’erreur pertinents.

## <a name="error-notifications"></a>Notifications d’erreur

La façon dont vous signalez des erreurs aux utilisateurs dépend du système d’interface utilisateur que vous utilisez. Si vous utilisez React comme système d’interface utilisateur, utilisez les composants d’interface utilisateur Fluent et les éléments de conception. Choisissez un contrôle approprié dans cette [page d’interface utilisateur Fluent](https://developer.microsoft.com/fluentui#/controls/web). Nous vous recommandons de transmettre les messages d’erreur à l’aide d’une barre de messages, d’une boîte de dialogue ou d’un modal. Si l’erreur se trouve dans l’entrée de l’utilisateur, affichez l’erreur en rouge gras près du contrôle d’entrée. L’exemple [Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) utilise un élément MessageBar et le modifie pour prendre en compte le menu personnalité dans un volet Office de complément.

Si vous n’utilisez pas React pour l’interface utilisateur, envisagez d’utiliser les anciens composants de l’interface utilisateur Fabric implémentés directement en HTML et JavaScript. Certains exemples de modèles se trouvent dans le référentiel [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates). Jetez un coup d’œil en particulier dans les sous-dossiers de dialogue et de navigation. L’exemple [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) utilise une bannière de message.

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Codes d'erreur de l'API commune de l'Office](../reference/javascript-api-for-office-error-codes.md)
