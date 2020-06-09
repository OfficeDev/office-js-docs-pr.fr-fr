---
title: Élément FunctionFile dans le fichier manifest
description: Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: db447a904c04d07d51119f1eac2556af536a647c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611840"
---
# <a name="functionfile-element"></a>Élément FunctionFile

Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur. L' `FunctionFile` élément est un élément enfant de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md). L' `resid` attribut de l' `FunctionFile` élément est défini sur la valeur de l' `id` attribut d’un `Url` élément dans l' `Resources` élément qui contient l’URL d’un fichier HTML qui contient ou charge toutes les fonctions JavaScript utilisées par les boutons de commande de complément sans interface utilisateur, comme défini par l' [élément Control](control.md).

Voici un exemple de l' `FunctionFile` élément.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

Le code JavaScript dans le fichier HTML indiqué par l' `FunctionFile` élément doit appeler `Office.initialize` et définir des fonctions nommées qui prennent un seul paramètre : `event` . Les fonctions doivent utiliser l’`item.notificationMessages` API pour indiquer l’avancement, réussite ou Échec de l’utilisateur. Elle doit également appeler `event.completed` lorsqu’il a fini d’exécution. Le nom des fonctions est utilisé dans l' `FunctionName` élément pour les boutons sans interface utilisateur.

Voici un exemple de fichier HTML définissant une `trackMessage` fonction.

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

Le code suivant montre comment implémenter la fonction utilisée par `FunctionName` .

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> L’appel de `event.completed` signale que vous avez réussi à gérer l’événement. Lorsqu’une fonction est appelée à plusieurs reprises, par exemple lorsque l’utilisateur clique plusieurs fois sur une même commande de complément, tous les événements sont automatiquement mis en file d’attente. Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente. Lorsque votre fonction appelle `event.completed` , l’appel suivant de cette fonction s’exécute. Vous devez appeler `event.completed` , sinon votre fonction ne s’exécutera pas.
