---
title: Initialiser votre complément Office
description: Découvrez comment initialiser votre complément Office.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: ca7b21c35fc82011c673cb83d077a89cc29f56eb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718936"
---
# <a name="initialize-your-office-add-in"></a>Initialiser votre complément Office

Les compléments Office ont souvent une logique de démarrage pour effectuer des actions telles que :

- Vérifiez que la version de l’utilisateur d’Office prend en charge toutes les API Office que votre code appelle.

- Vérifier l’existence de certains artefacts, tels qu’une feuille de calcul avec un nom spécifique.

- Inviter l’utilisateur à sélectionner certaines cellules dans Excel, puis insérer un graphique initialisé avec ces valeurs sélectionnées.

- Établir des liaisons.

- Utiliser l’API de boîte de dialogue Office pour inviter l’utilisateur à entrer les valeurs des paramètres de complément par défaut.

Toutefois, un complément Office ne peut pas appeler d’API JavaScript Office tant que la bibliothèque n’a pas été chargée. Cet article décrit les deux façons dont votre code peut s’assurer que la bibliothèque a été chargée :

- Initialiser avec `Office.onReady()`.
- Initialiser avec `Office.initialize`.

> [!TIP]
> Au lieu de `Office.initialize`, nous vous recommandons d’utiliser `Office.onReady()`. Bien `Office.initialize` que est toujours pris `Office.onReady()` en charge, offre davantage de flexibilité. Vous ne pouvez attribuer qu’un seul `Office.initialize` gestionnaire à et il n’est appelé qu’une seule fois par l’infrastructure Office. Vous pouvez appeler `Office.onReady()` à différents endroits de votre code et utiliser des rappels différents.
> 
> Pour plus d’informations sur les différences entre ces techniques, reportez-vous à la rubrique [Différences majeures entre Office.initialize et Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et environnement d’exécution](loading-the-dom-and-runtime-environment.md).

## <a name="initialize-with-officeonready"></a>Initialiser avec Office.onReady()

`Office.onReady()`est une méthode asynchrone qui renvoie un objet [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) pendant qu’il vérifie si la bibliothèque Office. js est chargée. Uniquement lorsque la bibliothèque est chargée, cela résout la promesse sous forme d’objet qui spécifie l’application Office hôte avec une`Office.HostType` valeur enum (`Excel`, `Word`, etc.) et la plateforme avec une`Office.PlatformType` valeur enum (`PC`, `Mac`, `OfficeOnline`, etc..). L’objet Promise se résout immédiatement si la bibliothèque est déjà chargée quand `Office.onReady()` est appelée.

Une méthode pour appeler `Office.onReady()` consiste à transmettre une méthode de rappel. Voici un exemple :

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

Par ailleurs, vous pouvez mettre en chaîne une`then()` méthode permettant d’appeler `Office.onReady()`, au lieu de spécifier un rappel. Par exemple, le code suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Voici le même exemple utilisant les mots clés `async` et `await` dans TypeScript :

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être*habituellement* placées dans la réponse à`Office.onReady()`. Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Toutefois, il existe des exceptions à cette pratique. Par exemple, supposons que vous voulez ouvrir votre complément dans un navigateur (au lieu de le charger dans un hôte Office) afin de déboguer votre interface utilisateur avec les outils de navigateur. Étant donné que Office.js ne sera pas chargé dans le navigateur, `onReady` ne s’exécutera pas et le `$(document).ready` ne s’exécutera pas si cette opération est appelée à l’intérieur d’Office `onReady`. 

Il est également possible d’afficher un indicateur de progression dans le volet Office pendant le chargement du complément. Dans ce scénario, votre code doit appeler jQuery `ready` et utiliser son rappel pour afficher l’indicateur de progression. Puis le rappel `onReady` Office peut remplacer l’indicateur de progression par l’interface utilisateur final. 

## <a name="initialize-with-officeinitialize"></a>Initialiser avec Office.initialize

Un événement initialisé se déclenche lorsque la bibliothèque Office.js est chargée et prête pour une interaction avec l’utilisateur. Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente votre logique d’initialisation. L’exemple suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, ceux-ci doivent *généralement* être placés au `Office.initialize` sein de l’événement (les exceptions décrites dans la section **Initialize with Office. onReady ()** ci-dessus s’appliquent également dans ce cas). Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Pour les compléments de tâches et de contenu, `Office.initialize` fournit un paramètre_raison_ supplémentaire. Ce paramètre peut être utilisé pour savoir comment un complément a été ajouté au document actif. Vous pouvez l’utiliser pour fournir une logique différente quand un complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document.

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

Pour plus d’informations, consultez les pages relatives à l’[événement Office.initialize](/javascript/api/office) et à l’[énumération InitializationReason](/javascript/api/office/office.initializationreason).

## <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principales différences entre Office.initialize et Office.onReady

- Vous ne pouvez assigner qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure d’Office, mais vous pouvez appeler `Office.onReady()`à plusieurs endroits dans votre code et utiliser des rappels différents. Par exemple, votre code pourrait appeler `Office.onReady()` dès que votre script personnalisé charge avec un rappel qui exécute la logique d’initialisation ; et votre code peut également comporter un bouton dans le volet Office dont le script appelle `Office.onReady()` avec un rappel différent. Si c’est le cas, le deuxième rappel s’exécute quand l’utilisateur clique sur le bouton.

- L’événement`Office.initialize` se déclenche à la fin du processus interne dans lequel Office.js s’initialise lui-même. Et il se déclenche *immédiatement* après la fin du processus interne. Si le code dans lequel vous attribuez un gestionnaire à l’événement s’exécute trop longtemps après le déclenchement de l’événement, votre gestionnaire ne s’exécutera pas. Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du complément pour charger les fichiers polyfill une fois que le serveur charge Office.js mais avant que le serveur ne charge votre code JavaScript personnalisé. Le temps que votre script se charge et affecte le Gestionnaire, l’événement initialisé s’est déjà produit. Mais il n’est jamais « trop tard » pour appeler `Office.onReady()`. Si l’événement initialisé s’est déjà produit, le rappel s’exécute immédiatement.

> [!NOTE]
> Même si vous n’avez aucune logique de démarrage, appelez `Office.onReady()` ou attribuez une fonction vide à `Office.initialize` lorsque votre complément JavaScript se charge. Certaines combinaisons de plateforme et d’hôte Office ne chargeront pas le volet Office tant que l’une de ces situations se produisent. Les exemples suivants présentent ces deux approches.
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md)