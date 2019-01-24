---
title: Gestion des erreurs
description: ''
ms.date: 10/16/2018
localization_priority: Normal
ms.openlocfilehash: 8c6de5d2a22fdb4614742ddfb7fbf566780c0f0f
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388961"
---
# <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.

> [!NOTE]
> Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript pour Excel, reportez-vous à la rubrique [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md).

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
|InvalidArgument |L’argument est manquant ou non valide, ou a un format incorrect.|
|InvalidRequest  |Impossible de traiter la demande.|
|InvalidReference|Cette référence n’est pas valide pour l’opération en cours.|
|InvalidBinding  |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.|
|InvalidSelection|La sélection en cours est incorrecte pour cette action.|
|Unauthenticated |Les informations d’authentification requises sont manquantes ou incorrectes.|
|AccessDenied |Vous ne pouvez pas effectuer l’opération demandée.|
|ItemNotFound |La ressource demandée n’existe pas.|
|ActivityLimitReached|La limite d’activité a été atteinte.|
|GeneralException|Une erreur interne s’est produite lors du traitement de la demande.|
|NotImplemented  |La fonctionnalité demandée n’est pas implémentée|
|ServiceNotAvailable|Le service n’est pas disponible.|
|Conflict|La demande n’a pas pu être traitée en raison d’un conflit.|
|ItemAlreadyExists|La ressource en cours de création existe déjà.|
|UnsupportedOperation|L’opération tentée n’est pas prise en charge.|
|RequestAborted|La demande a été interrompue pendant l’exécution.|
|ApiNotAvailable|L’API demandée n’est pas disponible.|
|InsertDeleteConflict|L’opération d’insertion ou de suppression tentée a créé un conflit.|
|InvalidOperation|L’opération tentée n’est pas valide sur l’objet.|

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error)
