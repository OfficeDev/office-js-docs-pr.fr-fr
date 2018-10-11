---
title: Gestion des erreurs
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b07012516cbe15374d0707c157738117a9c8fe96
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459230"
---
# <a name="error-handling"></a>Gestion des erreurs

Lorsque vous créez un complément à l'aide de l'API JavaScript Excel, veillez à inclure une logique de traitement des erreurs afin de prendre en compte les erreurs d'exécution. Cela est essentiel en raison de la nature asynchrone de l'API.

> [!NOTE]
> Pour plus d’informations sur la méthode **sync()** et la nature asynchrone de l’API JavaScript d'Excel, voir [Concepts de programmation fondamentaux avec l’API JavaScript d'Excel](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Meilleures pratiques

Tout au long des exemples de code de cette documentation, vous remarquerez que chaque appel à `Excel.run` est accompagné d’une `catch` instruction permettant de détecter les erreurs éventuelles dans le fichier `Excel.run`. Nous vous recommandons d'utiliser le même modèle lorsque vous créez un complément à l'aide des API JavaScript Excel.

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

Quand une demande d’API JavaScript pour Excel ne parvient pas à s’exécuter correctement, l’API renvoie un objet d’erreur qui contient les propriétés suivantes : 

- **code**: la propriété  `code` d’un message d’erreur contient une chaîne qui fait partie de la liste `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Par exemple, le code d’erreur «InvalidReference» indique que la référence n’est pas valide pour l’opération spécifiée. Les codes d’erreur ne sont pas localisés. 

- **message**: la propriété  `message` d’un message d’erreur contient un résumé de l’erreur dans la chaîne localisée. Le message d’erreur n’est pas destiné à la consommation par les utilisateurs finaux ; Vous devez utiliser le code d’erreur et la logique métier appropriée pour déterminer le message d’erreur que votre complément affiche aux utilisateurs finaux.

- **debugInfo** : le cas échéant, la propriété `debugInfo` du message d’erreur fournit des informations supplémentaires que vous pouvez utiliser pour comprendre la cause première de l’erreur. 

> [!NOTE]
> Si vous utilisez `console.log()` pour imprimer des messages d’erreur sur la console, ces messages ne seront visibles que sur le serveur. Les utilisateurs finaux ne verront pas ces messages d’erreur dans le panneau de tâches du complément ou ailleurs dans l'application hôte.

## <a name="see-also"></a>Voir aussi

- [Concepts  de programmation fondamentaux avec l’API JavaScript d'Excel](excel-add-ins-core-concepts.md)
- [Objet OfficeExtension.Error (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
