---
title: Présentation de l’interface API JavaScript pour Office
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 58829c623c06225bcc7d15925fb02a082df039c6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640091"
---
# <a name="understanding-the-javascript-api-for-office"></a>Présentation de l’interface API JavaScript pour Office

Cet article fournit des informations sur l’interface API JavaScript pour Office et son utilisation. Pour obtenir des informations de référence, voir [Interface API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). Pour plus d’informations sur la mise à jour des fichiers de projet Visual Studio vers la version la plus récente de l’interface API JavaScript pour Office, voir [Mettez à jour la version de votre interface API JavaScript pour Office et les fichiers de schéma manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de contrôle AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Par exemple, pour réussir le contrôle, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page Hôte et disponibilité du complément Office](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Référencement de la  bibliothèque de l’interface API JavaScript pour Office dans votre complément

La bibliothèque de l’[interface API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) comprend le fichier Office.js et des fichiers associés .js propres  à l’application hôte, comme Excel-15.js et  Outlook-15.js. La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Cette opération permet de télécharger et de mettre en cache les fichiers de l’interface API JavaScript pour Office lors du premier chargement de votre complément, pour garantir que l’interface utilise l’implémentation d’Office.js la plus récente et les fichiers .js qui lui sont associés pour la version indiquée.

Pour en savoir plus sur le CDN Office.js, y compris sur la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de diffusion de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Initialisation de votre complément

**S’applique à :** tous les types de complément

Les compléments Office ont souvent une logique de démarrage pour effectuer des tâches telles que :

- Vérifier que la version utilisateur d’Office prendra en charge toutes les API Office appelées par votre code.

- Vérifier l’existence de certains artifacts, tels qu’une feuille de calcul avec un nom spécifique.

- Inviter l’utilisateur à sélectionner des cellules dans Excel, puis d’insérer un graphique initialisé avec ces valeurs sélectionnées.

- Établir des liaisons.

- Depuis la boîte de dialogue API Office, inviter l’utilisateur de paramétrer les valeurs de paramètres de module complémentaire par défaut.

Mais votre code de démarrage ne doit pas appeler d’API Office.js avant que la bibliothèque ne soit entièrement chargée. Il existe deux manières de lancer le chargement de la bibliothèque dans votre code. Elles sont décrites dans les sections suivantes : 

- [Initialiser avec Office.onReady()](#initialize-with-officeonready)
- [Initialiser avec Office.initialize](#initialize-with-officeinitialize)

Pour plus d’informations sur les différences entre ces techniques, reportez-vous à l’article [Principales différences entre Office.initialize et Office.onReady()](#major-differences-between-officeinitialize-and-officeonready). Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>Initialiser avec Office.onReady()

`Office.onReady()` est une méthode asynchrone qui renvoie un objet  Promise pendant qu’il vérifie si la bibliothèque Office.js est entièrement chargée. Lorsque, et uniquement lorsque la bibliothèque est chargée, elle résout Promise en tant qu’objet qui spécifie l’application hôte Office avec une valeur enum `Office.HostType` (`Excel`, `Word`, etc.) et la plateforme avec une valeur enum `Office.PlatformType`  (`PC`, `Mac`, `OfficeOnline`, etc.). Si la bibliothèque est déjà chargée lorsque `Office.onReady()` est appelée, Promise est immédiatement résolu.

Une façon d’appeler `Office.onReady()` consiste à lui transmettre une méthode callback. Voici un exemple :

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

Sinon, vous pouvez chaîner une méthode `then()`  à l’appel de `Office.onReady()`, au lieu de passer un rappel. Par exemple, le code suivant vérifie que la version d’Excel de l’utilisateur prend en charge toutes les API que le complément peut appeler.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Voici le même exemple utilisant les mots-clés `async` et `await` dans TypeScript :

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou des tests, ceux-ci doivent *généralement* être placés dans la réponse à `Office.onReady()`. Par exemple, la fonction de [JQuery’s](https://jquery.com), `$(document).ready()`, serait référencée comme suit :

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Toutefois, il existe des exceptions à cette pratique. Par exemple, supposons que vous souhaitiez ouvrir votre complément dans un navigateur (au lieu d’en charger une version test dans un hôte Office) pour déboguer votre interface utilisateur avec les outils du navigateur. Puisque Office.js ne se charge pas dans le navigateur, `onReady` ne s’exécutera pas et le `$(document).ready` ne s’exécutera pas, s’il est appelé dans Office `onReady`. Une autre exception : vous voulez qu’un indicateur de progression apparaisse dans le volet Office pendant le chargement du complément. Dans ce scénario, votre code devrait appeler le jQuery `ready` et utiliser son rappel pour afficher l’indicateur de progression. Puis le rappel `onReady`d’Office peut remplacer l’indicateur de progression par l’interface utilisateur finale. 

### <a name="initialize-with-officeinitialize"></a>Initialiser avec Office.initialize

Un événement d’initialisation se déclenche lorsque la bibliothèque Office.js est entièrement chargée et prête pour l’interaction utilisateur. Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente votre logique d’initialisation. L’exemple suivant vérifie si la version d’Excel de l’utilisateur prend en charge toutes les API que le complément peut appeler.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou des tests, ceux-ci doivent *généralement* être placés dans l’événement `Office.initialize` . (Mais les exceptions décrites précédemment dans la section **Initialiser avec Office.onReady()** s’appliquent également dans ce cas.) Par exemple, la fonction de [JQuery](https://jquery.com), `$(document).ready()`, serait référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Pour les compléments de contenu et du volet Office, `Office.initialize` fournit un paramètre _reason_ supplémentaire. Ce paramètre spécifie comment un complément a été ajouté au document actif. Vous pouvez l’utiliser pour fournir une logique différente pour le moment où un complément est inséré pour la première fois par rapport au moment où il existait déjà dans le document.

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

Pour plus d’informations, consultez les pages relatives à l’[événement Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) et à l’[énumération InitializationReason](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js)

> [!NOTE]
> Actuellement, vous devez définir `Office.Initialize`, indépendamment du fait que `Office.onReady()` soit également appelée. Si vous n’avez pas besoin de `Office.Initialize`, vous pouvez le définir sur une fonction vide comme indiqué dans l’exemple suivant.
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principales différences entre Office.initialize et Office.onReady

- Vous ne pouvez assigner qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure Office ; mais vous pouvez appeler `Office.onReady()` à différents endroits dans votre code et utiliser différents rappels. Par exemple, votre code pourrait appeler `Office.onReady()` dès que votre script personnalisé se charge avec un rappel qui exécute la logique d’initialisation ; et votre code pourrait aussi avoir un bouton dans le volet Office, dont le script appelle `Office.onReady()` avec un rappel différent. Si tel est le cas, le second rappel s’exécute lorsque l’utilisateur clique sur le bouton.

- L’événement `Office.initialize` est déclenché à la fin du processus interne au cours duquel Office.js s’initialise. Et il se déclenche *immédiatement* après la fin du processus interne. Si le code dans lequel vous affectez un gestionnaire à l’événement s’exécute trop longtemps après que l’événement se soit déclenché, votre gestionnaire ne s’exécute pas.Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du module complémentaire pour charger les fichiers polyfill après le chargement des Office.js mais avant de charger votre code JavaScript personnalisé. Au moment où votre script est chargé et affecte le gestionnaire, l’événement Initialiser a déjà été exécuté. Mais il n’est jamais « trop tard » pour appeler `Office.onReady()`. Si l’événement Initialiser a déjà eu lieu, le rappel s’exécute immédiatement.

> [!NOTE]
> Même si vous n’avez aucune logique de démarrage, vous devez affecter une fonction vide à `Office.initialize` lorsque votre complément JavaScript se charge, comme indiqué dans l’exemple suivant. Certaines combinaisons d’hôte et la plateforme Office ne chargent pas le volet Office tant que l’événement initialiser ne se déclenche pas et que la fonction de gestionnaire d’événements spécifiée ne s’exécute pas.
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modèle objet JavaScript Office

Une fois initialisé, le complément peut interagir avec l’hôte (par exemple, Excel, Outlook). La page [Modèle d’objet API JavaScript pour Office](office-javascript-api-object-model.md) a plus de détails sur les modèles d’utilisations spécifiques. Il existe également une documentation de référence détaillée pour les [API partagées](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) et les hôtes spécifiques.

## <a name="api-support-matrix"></a>Matrice de prise en charge d’API

Ce tableau récapitule l’API et les fonctionnalités prises en charge dans les types de complément (contenu, volet Office et Outlook), ainsi que les applications Office qui peuvent les héberger lorsque vous indiquez les applications hôte Office prises en charge par votre complément, à l’aide du [schéma de manifeste de complément 1.1 et des fonctionnalités prises en charge par la version 1.1 de l’interface API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nom d’hôte**|Base de données|Classeur|Boîte aux lettres|Présentation|Document|Projet|
||**Applications hôtes** **prises en charge**|Applications Web Access|Excel,<br/>Excel Online|Outlook,<br/>Application web Outlook,<br/>OWA pour les appareils|PowerPoint,<br/>PowerPoint Online|Word|Projet|
|**Types de compléments pris en charge**|Contenu|v|v||v|||
||Volet Office||v||v|v|v|
||Outlook|||v||||
|**Fonctionnalités d’API prises en charge**|Texte en lecture et en écriture||v||v|v|v<br/>(En lecture seule)|
||Matrice en lecture et en écriture||v|||v||
||Tableau en lecture et en écriture||v|||v||
||HTML en lecture et en écriture|||||v||
||En lecture et en écriture<br/>Office Open XML|||||v||
||Lecture des propriétés de tâche, de ressource, de vue et de champ||||||v|
||Sélection des événements modifiés||v|||v||
||Obtention de l’ensemble du document||||v|v||
||Liaisons et événements de liaison|v<br/>(Liaisons de tableau complètes et partielles uniquement)|v|||v||
||Parties XML personnalisées en lecture et en écriture|||||v||
||Faire persister les données d’état de complément (paramètres)|v<br/>(Par complément hôte)|v<br/>(Par document)|v<br/>(Par boîte aux lettres)|v<br/>(Par document)|v<br/>(Par document)||
||Paramètres des événements modifiés|v|v||v|v||
||Obtention du mode de vue active<br/>et affichage des événements modifiés||||v|||
||Accès à des emplacements<br/>dans le document||v||v|v||
||Activation en fonction du contexte<br/>à l’aide de règles et de RegEx|||v||||
||Lecture des propriétés d’élément|||v||||
||Lecture de profil utilisateur|||v||||
||Obtention des pièces jointes|||v||||
||Obtention du jeton d’identité d’utilisateur|||v||||
||Appel des services web Exchange|||v||||
