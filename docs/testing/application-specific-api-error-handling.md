---
title: Gestion des erreurs avec les API JavaScript spécifiques à l’application
description: Découvrez Excel, Word, PowerPoint et d’autres logiques de gestion des erreurs propres à l’API JavaScript spécifiques à l’application pour prendre en compte les erreurs d’exécution.
ms.date: 10/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21d8d3eef36f919f95459fd8e0b3037c1d5ae1b1
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767153"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>Gestion des erreurs avec les API JavaScript spécifiques à l’application

Lorsque vous créez un complément à l’aide [des API JavaScript Office spécifiques à l’application](../develop/application-specific-api-model.md), veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Cela est essentiel en raison de la nature asynchrone des API.

## <a name="best-practices"></a>Meilleures pratiques

Dans nos [exemples de code](https://github.com/OfficeDev/Office-Add-in-samples) et [extraits de code Script Lab](../overview/explore-with-script-lab.md), vous remarquerez que chaque appel à `Excel.run`, `PowerPoint.run`ou `Word.run` est accompagné d’une instruction pour intercepter les `catch` erreurs. Nous vous recommandons d’utiliser le même modèle lorsque vous générez un complément à l’aide des API spécifiques à l’application.

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

Lorsqu’une demande d’API JavaScript Office ne s’exécute pas correctement, l’API retourne un objet d’erreur qui contient les propriétés suivantes.

- **code** : la `code` propriété d’un message d’erreur contient une chaîne qui fait partie de `OfficeExtension.ErrorCodes` ou `{application}.ErrorCodes` où *{application}* représente Excel, PowerPoint ou Word. Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas traduits.

- **message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas destiné à être consommé par les utilisateurs finaux ; vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finaux.

- **debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.

> [!NOTE]
> Si vous utilisez `console.log()` pour imprimer des messages d’erreur sur la console, ces messages sont uniquement visibles sur le serveur. Les utilisateurs finaux ne voient pas ces messages d’erreur dans le volet Office du complément ou n’importe où dans l’application Office. Pour signaler des erreurs à l’utilisateur, consultez [Notifications d’erreurs](#error-notifications).

## <a name="error-codes-and-messages"></a>Codes d’erreur et messages

Les tableaux suivants répertorient les erreurs que les API spécifiques à l’application peuvent retourner.

> [!NOTE]
> Les tableaux suivants répertorient les messages d’erreur que vous pouvez rencontrer lors de l’utilisation des API spécifiques à l’application. Si vous utilisez l’API commune, consultez [Codes d’erreur de l’API commune Office](../reference/javascript-api-for-office-error-codes.md) pour en savoir plus sur les messages d’erreur pertinents.

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`AccessDenied` |Vous ne pouvez pas effectuer l’opération demandée.|*Aucun.* |
|`ActivityLimitReached`|La limite d’activité a été atteinte.|*Aucun.* |
|`ApiNotAvailable`|L’API demandée n’est pas disponible.|*Aucun.* |
|`ApiNotFound`|L’API que vous essayez d’utiliser est introuvable. Il peut être disponible dans une version plus récente de l’application Office. Pour plus d’informations, voir Disponibilité des applications [clientes et des plateformes Office pour les compléments Office](/javascript/api/requirement-sets) .|*Aucun.* |
|`BadPassword`|Le mot de passe que vous avez fourni est incorrect.|*Aucun.* |
|`Conflict`|La demande n’a pas pu être traitée en raison d’un conflit.|*Aucun.* |
|`ContentLengthRequired`|Un `Content-length` en-tête HTTP est manquant.|*Aucun.* |
|`GeneralException`|Une erreur interne s’est produite lors du traitement de la demande.|*Aucun.* |
|`InsertDeleteConflict`|L’opération d’insertion ou de suppression tentée a créé un conflit.|*Aucun.* |
|`InvalidArgument` |L’argument est manquant ou non valide, ou a un format incorrect.|*Aucun.* |
|`InvalidBinding` |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.|*Aucun.* |
|`InvalidOperation`|L’opération tentée n’est pas valide sur l’objet.|*Aucun.* |
|`InvalidReference`|Cette référence n’est pas valide pour l’opération en cours.|*Aucun.* |
|`InvalidRequest`  |Impossible de traiter la demande.|*Aucun.* |
|`InvalidSelection`|La sélection en cours est incorrecte pour cette action.|*Aucun.* |
|`ItemAlreadyExists`|La ressource en cours de création existe déjà.|*Aucun.* |
|`ItemNotFound` |La ressource demandée n’existe pas.|*Aucun.* |
|`MemoryLimitReached`|La limite de mémoire a été atteinte. Votre action n’a pas pu être terminée.|*Aucun.* |
|`NotImplemented`|La fonctionnalité demandée n’est pas implémentée| Cela peut signifier que l’API est en préversion ou prise en charge uniquement sur une plateforme particulière (par exemple, en ligne uniquement). Pour plus d’informations, voir Disponibilité des applications [clientes et des plateformes Office pour les compléments Office](/javascript/api/requirement-sets) .|
|`RequestAborted`|La demande a été interrompue pendant l’exécution.|*Aucun.* |
|`RequestPayloadSizeLimitExceeded`|La taille de la charge utile de la requête a dépassé la limite. Pour plus d’informations, consultez l’article [Limites de ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .| Cette erreur se produit uniquement dans Office sur le Web.|
|`ResponsePayloadSizeLimitExceeded`|La taille de la charge utile de réponse a dépassé la limite. Pour plus d’informations, consultez l’article [Limites de ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .|  Cette erreur se produit uniquement dans Office sur le Web.|
|`ServiceNotAvailable`|Le service n’est pas disponible.|*Aucun.* |
|`Unauthenticated` |Les informations d’authentification requises sont manquantes ou incorrectes.|*Aucun.* |
|`UnsupportedFeature`|L’opération a échoué, car la feuille de calcul source contient une ou plusieurs fonctionnalités non prises en charge.|*Aucun.* |
|`UnsupportedOperation`|L’opération tentée n’est pas prise en charge.|*Aucun.* |

### <a name="excel-specific-error-codes-and-messages"></a>Codes d’erreur et messages propres à Excel

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`EmptyChartSeries`|L’opération tentée a échoué, car la série de graphiques est vide.|*Aucun.* |
|`FilteredRangeConflict`|L’opération tentée provoque un conflit avec une plage filtrée.|*Aucun.* |
|`FormulaLengthExceedsLimit`|Le bytecode de la formule appliquée dépasse la limite de longueur maximale. Pour Office sur les ordinateurs 32 bits, la longueur limite du bytecode est de 16 384 caractères. Sur les ordinateurs 64 bits, la limite de longueur du bytecode est de 32 768 caractères.| Cette erreur se produit à la fois dans Excel sur le Web et sur le bureau.|
|`GeneralException`|*Divers.*|Les API de types de données retournent `GeneralException` des erreurs avec des messages d’erreur dynamiques. Ces messages font référence à la cellule qui est la source de l’erreur et au problème à l’origine de l’erreur, par exemple : « La cellule A1 ne contient pas la propriété `type`requise ».|
|`InactiveWorkbook`|L’opération a échoué, car plusieurs classeurs sont ouverts et le classeur appelé par cette API a perdu le focus.|*Aucun.* |
|`InvalidOperationInCellEditMode`|L’opération n’est pas disponible lorsqu’Excel est en mode Modifier la cellule. Quittez le mode Édition en utilisant les **touches Entrée** ou **Tab** , ou en sélectionnant une autre cellule, puis réessayez.|*Aucun.* |
|`MergedRangeConflict`|Impossible de terminer l’opération. Une table ne peut pas chevaucher une autre table, un rapport de tableau croisé dynamique, des résultats de requête, des cellules fusionnées ou un mappage XML.|*Aucun.* |
|`NonBlankCellOffSheet`|Microsoft Excel ne peut pas insérer de nouvelles cellules, car il pousserait les cellules non vides hors de la fin de la feuille de calcul. Ces cellules non vides peuvent apparaître vides, mais avoir des valeurs vides, une mise en forme ou une formule. Supprimez suffisamment de lignes ou de colonnes pour libérer de l’espace pour ce que vous souhaitez insérer, puis réessayez.|*Aucun.* |
|`OperationCellsExceedLimit`|L’opération tentée affecte plus que la limite de 33554000 cellules.| Si le `TableColumnCollection.add API` déclenche cette erreur, vérifiez qu’il n’y a pas de données involontaires dans la feuille de calcul, mais en dehors de la table. En particulier, recherchez les données dans les colonnes les plus à droite de la feuille de calcul. Supprimez les données involontaires pour résoudre cette erreur. Une façon de vérifier le nombre de cellules qu’une opération traite consiste à exécuter le calcul suivant : `(number of table rows) x (16383 - (number of table columns))`. Le nombre 16383 est le nombre maximal de colonnes pris en charge par Excel. <br><br>Cette erreur se produit uniquement dans Excel sur le Web. |
|`PivotTableRangeConflict`|L’opération tentée provoque un conflit avec une plage de tableau croisé dynamique.|*Aucun.* |
|`RangeExceedsLimit`|Le nombre de cellules dans la plage a dépassé le nombre maximal pris en charge. Pour plus d’informations, consultez l’article [Limites de ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .|*Aucun.* |
|`RefreshWorkbookLinksBlocked`|L’opération a échoué, car l’utilisateur n’a pas accordé l’autorisation d’actualiser les liens de classeur externe.|*Aucun.* |
|`UnsupportedSheet`|Ce type de feuille ne prend pas en charge cette opération, car il s’agit d’une feuille Macro ou Graphique.|*Aucun.* |

### <a name="word-specific-error-codes-and-messages"></a>Codes d’erreur et messages spécifiques à Word

|Code d’erreur | Message d’erreur | Remarques |
|:----------|:--------------|:------|
|`SearchDialogIsOpen`|La boîte de dialogue de recherche est ouverte.|*Aucun.* |
|`SearchStringInvalidOrTooLong`|La chaîne de recherche n’est pas valide ou est trop longue.| La chaîne de recherche maximale est de 255 caractères. |

## <a name="error-notifications"></a>Notifications d’erreurs

La façon dont vous signalez les erreurs aux utilisateurs dépend du système d’interface utilisateur que vous utilisez. Si vous utilisez React comme système d’interface utilisateur, utilisez les composants de l’interface utilisateur Fluent et les éléments de conception. Choisissez un contrôle approprié dans cette [page de l’interface utilisateur Fluent](https://developer.microsoft.com/fluentui#/controls/web). Nous vous recommandons de transmettre les messages d’erreur à l’aide d’une barre de messages, d’une boîte de dialogue ou d’une fenêtre modale. Si l’erreur se trouve dans l’entrée de l’utilisateur, affichez l’erreur en gras rouge près du contrôle d’entrée. [L’exemple Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) utilise un élément MessageBar et le modifie pour prendre en compte le menu de personnalité dans le volet Office d’un complément.

Si vous n’utilisez pas React pour l’interface utilisateur, envisagez d’utiliser les anciens composants de l’interface utilisateur Fabric implémentés directement en HTML et JavaScript. Certains exemples de modèles se trouvent dans le référentiel [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) . Examinez en particulier la boîte de dialogue et les sous-dossiers de navigation. L’exemple [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) utilise une bannière de message.

## <a name="see-also"></a>Voir aussi

- [Objet OfficeExtension.Error](/javascript/api/office/officeextension.error)
- [Codes d'erreur de l'API commune de l'Office](../reference/javascript-api-for-office-error-codes.md)
