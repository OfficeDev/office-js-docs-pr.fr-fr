---
title: Initialiser votre complément Office
description: Découvrez comment initialiser votre complément Office.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52e75770dc4852ac3905256b6ea4230552df48ca
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797595"
---
# <a name="initialize-your-office-add-in"></a>Initialiser votre complément Office

Les compléments Office ont souvent une logique de démarrage pour effectuer des actions telles que :

- Vérifiez que la version d’Office de l’utilisateur prend en charge toutes les API Office que votre code appelle.

- Vérifiez l’existence de certains artefacts, tels qu’une feuille de calcul portant un nom spécifique.

- Invitez l’utilisateur à sélectionner certaines cellules dans Excel, puis insérez un graphique initialisé avec les valeurs sélectionnées.

- Établir des liaisons.

- Utilisez l’API de boîte de dialogue Office pour demander à l’utilisateur les valeurs de paramètres de complément par défaut.

Toutefois, un complément Office ne peut pas appeler correctement les API JavaScript Office tant que la bibliothèque n’a pas été chargée. Cet article décrit les deux façons dont votre code peut s’assurer que la bibliothèque a été chargée.

- Initialiser avec `Office.onReady()`.
- Initialiser avec `Office.initialize`.

> [!TIP]
> Au lieu de `Office.initialize`, nous vous recommandons d’utiliser `Office.onReady()`. Bien qu’elle soit `Office.initialize` toujours prise en charge, `Office.onReady()` elle offre plus de flexibilité. Vous ne pouvez affecter qu’un seul gestionnaire et `Office.initialize` il n’est appelé qu’une seule fois par l’infrastructure Office. Vous pouvez appeler `Office.onReady()` à différents endroits dans votre code et utiliser différents rappels.
> 
> Pour plus d’informations sur les différences entre ces techniques, reportez-vous à la rubrique [Différences majeures entre Office.initialize et Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et environnement d’exécution](loading-the-dom-and-runtime-environment.md).

## <a name="initialize-with-officeonready"></a>Initialiser avec Office.onReady()

`Office.onReady()` est une méthode asynchrone qui retourne un objet [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) pendant qu’il vérifie si la bibliothèque Office.js est chargée. Lorsque la bibliothèque est chargée, elle résout la promesse en tant qu’objet qui spécifie l’application cliente Office avec une `Office.HostType` valeur d’énumération (`Excel`, `Word`etc.) et la plateforme avec une `Office.PlatformType` valeur d’énumération (`PC`, `Mac`, `OfficeOnline`etc.). L’objet Promise se résout immédiatement si la bibliothèque est déjà chargée quand `Office.onReady()` est appelée.

Une méthode pour appeler `Office.onReady()` consiste à transmettre une méthode de rappel. Voici un exemple.

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

Voici le même exemple d’utilisation des mots clés et `await` des `async` mots clés dans TypeScript.

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être *habituellement* placées dans la réponse à`Office.onReady()`. Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Toutefois, il existe des exceptions à cette pratique. Par exemple, supposons que vous souhaitez ouvrir votre complément dans un navigateur (au lieu de le charger de manière indépendante dans une application Office) afin de déboguer votre interface utilisateur avec les outils de navigateur. Dans ce scénario, une fois que Office.js détermine qu’elle s’exécute en dehors d’une application hôte Office, elle appelle le rappel et résout la promesse avec `null` laquelle elle s’exécute à la fois pour l’hôte et pour la plateforme.

Une autre exception serait si vous souhaitez qu’un indicateur de progression apparaisse dans le volet Office pendant le chargement du complément. Dans ce scénario, votre code doit appeler jQuery `ready` et utiliser son rappel pour afficher l’indicateur de progression. Ensuite, le `Office.onReady` rappel peut remplacer l’indicateur de progression par l’interface utilisateur finale.

## <a name="initialize-with-officeinitialize"></a>Initialiser avec Office.initialize

Un événement initialisé se déclenche lorsque la bibliothèque Office.js est chargée et prête pour une interaction avec l’utilisateur. Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente votre logique d’initialisation. L’exemple suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Si vous utilisez des frameworks JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, ceux-ci doivent *généralement* être placés dans l’événement `Office.initialize` (les exceptions décrites dans la section **Initialiser avec Office.onReady()** s’appliquent également dans ce cas). Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Pour les compléments de tâches et de contenu, `Office.initialize` fournit un paramètre _raison_ supplémentaire. Ce paramètre peut être utilisé pour savoir comment un complément a été ajouté au document actif. Vous pouvez l’utiliser pour fournir une logique différente quand un complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document.

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
> Même si vous n’avez aucune logique de démarrage, appelez `Office.onReady()` ou attribuez une fonction vide à `Office.initialize` lorsque votre complément JavaScript se charge. Certaines combinaisons d’applications et de plateformes Office ne chargent pas le volet Office tant que l’une de ces combinaisons n’a pas lieu. Les exemples suivants présentent ces deux approches.
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="debug-initialization"></a>Initialisation du débogage

Pour plus d’informations sur le débogage des méthodes et `Office.onReady()` des `Office.initialize` méthodes, consultez [Déboguer les méthodes initialize et onReady](../testing/debug-initialize-onready.md).

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md)