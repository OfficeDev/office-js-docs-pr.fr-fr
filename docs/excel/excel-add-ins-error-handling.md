---
title: Gestion des erreurs avec l’API JavaScript pour Excel
description: En savoir plus sur la logique de gestion des erreurs de l’API JavaScript Excel pour prendre en compte les erreurs d’exécution.
ms.date: 10/22/2020
localization_priority: Normal
ms.openlocfilehash: a3b1bbfa7daba1b856bce35aa075d5b625bd9769
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740818"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Gestion des erreurs avec l’API JavaScript pour Excel

Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.

> [!NOTE]
> Pour plus d’informations sur la `sync()` méthode et la nature asynchrone de l’API JavaScript pour Excel, reportez-vous à la rubrique [Excel JavaScript Object Model in Office Add-ins](excel-add-ins-core-concepts.md).

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

Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes :

- **code** :  la propriété `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur « InvalidReference » indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas traduits.

- **message** : la propriété `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas conçu pour être utilisé par l’utilisateur final. Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finals.

- **debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause principale de l’erreur.

> [!NOTE]
> si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finals ne verront pas ces messages d’erreur dans le volet Office du complément ni n’importe où dans l’application Office.

## <a name="error-messages"></a>Messages d’erreur

Le tableau suivant contient la liste des erreurs que l’API peut renvoyer.

|Code d’erreur | Message d’erreur |
|:----------|:--------------|
|`AccessDenied` |Vous ne pouvez pas effectuer l’opération demandée.|
|`ActivityLimitReached`|La limite d’activité a été atteinte.|
|`ApiNotAvailable`|L’API demandée n’est pas disponible.|
|`ApiNotFound`|L’API que vous tentez d’utiliser est introuvable. Il est peut-être disponible dans une version plus récente d’Excel. Pour plus d’informations, consultez l’article [ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md) .|
|`BadPassword`|Le mot de passe que vous avez entré est incorrect.|
|`Conflict`|La demande n’a pas pu être traitée en raison d’un conflit.|
|`ContentLengthRequired`|Un `Content-length` en-tête http est manquant.|
|`GeneralException`|Une erreur interne s’est produite lors du traitement de la demande.|
|`InsertDeleteConflict`|L’opération d’insertion ou de suppression tentée a créé un conflit.|
|`InvalidArgument` |L’argument est manquant ou non valide, ou a un format incorrect.|
|`InvalidBinding`  |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.|
|`InvalidOperation`|L’opération tentée n’est pas valide sur l’objet.|
|`InvalidReference`|Cette référence n’est pas valide pour l’opération en cours.|
|`InvalidRequest`  |Impossible de traiter la demande.|
|`InvalidSelection`|La sélection en cours est incorrecte pour cette action.|
|`ItemAlreadyExists`|La ressource en cours de création existe déjà.|
|`ItemNotFound` |La ressource demandée n’existe pas.|
|`NonBlankCellOffSheet`|La demande d’insertion de nouvelles cellules ne peut pas être effectuée, car elle replacerait les cellules non vides à la fin de la feuille de calcul. Ces cellules non vides peuvent apparaître vides mais avoir des valeurs vides, une mise en forme ou une formule. Supprimez suffisamment de lignes ou de colonnes pour libérer de l’espace pour ce que vous voulez insérer, puis réessayez.|
|`NotImplemented`|La fonctionnalité demandée n’est pas implémentée|
|`RangeExceedsLimit`|Le nombre de cellules dans la plage a dépassé le nombre maximal pris en charge. Pour plus d’informations, reportez-vous à l’article [limitations des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .|
|`RequestAborted`|La demande a été interrompue pendant l’exécution.|
|`RequestPayloadSizeLimitExceeded`|La taille de la charge utile de la demande a dépassé la limite. Pour plus d’informations, reportez-vous à l’article [limitations des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) . <br><br>Cette erreur se produit uniquement dans Excel sur le Web.|
|`ResponsePayloadSizeLimitExceeded`|La taille de la charge utile de la réponse a dépassé la limite. Pour plus d’informations, reportez-vous à l’article [limitations des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .  <br><br>Cette erreur se produit uniquement dans Excel sur le Web.|
|`ServiceNotAvailable`|Le service n’est pas disponible.|
|`Unauthenticated` |Les informations d’authentification requises sont manquantes ou incorrectes.|
|`UnsupportedOperation`|L’opération tentée n’est pas prise en charge.|
|`UnsupportedSheet`|Ce type de feuille ne prend pas en charge cette opération, car il s’agit d’une macro ou d’une feuille de graphique.|

## <a name="see-also"></a>Voir aussi

- [Modèle objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
