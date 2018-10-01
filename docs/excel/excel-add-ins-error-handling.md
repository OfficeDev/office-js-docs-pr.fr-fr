---
title: Gestion des erreurs
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348085"
---
# <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément à l’aide de l’API JavaScript pour Excel, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. Il s’agit d’une étape essentielle en raison de la nature asynchrone de l’API.

> [!NOTE]
> Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript pour Excel, reportez-vous à la rubrique [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md).

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
> Si vous utilisez `console.log()` pour imprimer les messages d’erreur de la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finaux ne verront pas ces messages d’erreur dans le volet Office du complément ou n’importe où dans l’application hôte.

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
