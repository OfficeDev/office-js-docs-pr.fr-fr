---
title: Gestion des erreurs
description: En savoir plus sur la logique de gestion des erreurs de l’API JavaScript Excel pour prendre en compte les erreurs d’exécution.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: 85fcd580828a2db95cd8e021dec3611ca6591e1c
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43225727"
---
# <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.

> [!NOTE]
> Pour plus d’informations sur `sync()` la méthode et la nature asynchrone de l’API JavaScript pour Excel, voir [concepts de programmation fondamentaux avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md).

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
> Si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finals ne verront pas ces messages d’erreur dans le volet Office du complément ni n’importe où dans l’application hôte.

## <a name="error-messages"></a>Messages d’erreur

Le tableau suivant contient la liste des erreurs que l’API peut renvoyer.

|error.code | error.message |
|:----------|:--------------|
|`AccessDenied` |Vous ne pouvez pas effectuer l’opération demandée.|
|`ActivityLimitReached`|La limite d’activité a été atteinte.|
|`ApiNotAvailable`|L’API demandée n’est pas disponible.|
|`Conflict`|La demande n’a pas pu être traitée en raison d’un conflit.|
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
|`NotImplemented`  |La fonctionnalité demandée n’est pas implémentée|
|`RequestAborted`|La demande a été interrompue pendant l’exécution.|
|`ServiceNotAvailable`|Le service n’est pas disponible.|
|`Unauthenticated` |Les informations d’authentification requises sont manquantes ou incorrectes.|
|`UnsupportedOperation`|L’opération tentée n’est pas prise en charge.|

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](/javascript/api/office/officeextension.error)
