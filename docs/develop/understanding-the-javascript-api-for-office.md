---
title: Présentation de l’API JavaScript pour Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 12e7d9030ec37746f84e3fc725cddda2a5675761
ms.sourcegitcommit: 5bef9828f047da03ecf2f43c6eb5b8514eff28ce
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/31/2018
ms.locfileid: "23782793"
---
# <a name="understanding-the-javascript-api-for-office"></a>Présentation de l’API JavaScript pour Office

Cet article fournit des informations sur l’API JavaScript pour Office et son utilisation. Pour obtenir des informations de référence, voir [API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). Pour plus d’informations sur la mise à jour des fichiers de projet Visual Studio vers la version la plus récente de l’API JavaScript pour Office, voir [Mettre à jour la version de votre API JavaScript pour Office et les fichiers de schéma manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Référence à la bibliothèque de l’interface API JavaScript pour Office dans votre complément

La bibliothèque de l’[interface API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Cette opération permet de télécharger et de mettre en cache les fichiers de l’interface API JavaScript pour Office lors du premier chargement de votre complément pour garantir qu’elle utilise l’implémentation d’Office.js la plus récente et les fichiers .js qui lui sont associés pour la version indiquée.

Pour en savoir plus sur le CDN Office.js, y compris sur la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de la bibliothèque de l'interface API JavaScript pour Office à partir de son réseau de diffusion de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Initialisation de votre complément

**S’applique à :** tous les types de complément

Les modules complémentaires Office requièrent souvent de suivre une logique de démarrage telle que :

- Vérifier que la version utilisateur d’Office prendra en charge toutes les API Office appelées par votre code.

- Vérifier l’existence de certains artifacts, tels qu'une feuille de calcul avec un nom spécifique.

- Inviter l’utilisateur à sélectionner des cellules dans Excel, puis d’insérer un graphique initialisé avec ces valeurs sélectionnées.

- Établir des liaisons.

- Depuis la boîte de dialogue API Office, inviter l’utilisateur de paramétrer les valeurs de paramètres de module complémentaire par défaut.

Mais votre code de démarrage ne doit pas appeler d'API Office.js avant que la bibliothèque ne soit entièrement chargée. Il existe deux manières de lancer le chargement de la bibliothèque dans votre code. Elles sont décrites dans les sections ci-dessous. Nous vous conseillons d’utiliser la technique plus récente, plus souple, qui appelle `Office.onReady()`. Mais la technique antérieure, qui affecte un gestionnaire à `Office.initialize`, est toujours prise en charge. Voir aussi les [principales différences entre Office.initialize et Office.onReady()](#major-differences-between-office-initialize-and-office-onready).

Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>Initialisation avec Office.onReady()

`Office.onReady()` est une méthode asynchrone qui renvoie un objet Promesse tout en vérifiant si la bibliothèque Office.js est entièrement chargée. Si, et uniquement si, la bibliothèque est chargée, l'objet Promesse est réalisé en tant qu’objet qui spécifie l’application Office hôte avec une valeur enum `Office.HostType` (`Excel`, `Word`, etc.) et la plateforme avec une valeur enum `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline`, etc.). Si la bibliothèque est déjà chargée lors de l'appel de `Office.onReady()` et l'objet Promesse est résolu immédiatement.

Une manière d'appeler `Office.onReady()` consiste à passer une méthode de rappel. Voici un exemple :

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

Sinon, vous pouvez enchaîner une méthode `then()` à l’appel de `Office.onReady()`, au lieu de passer un rappel. Par exemple, le code suivant vérifie que la version utilisateur d’Excel prend en charge toutes les API que le complément peut appeler.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Voici le même exemple utilisant les mots-clés `async` et `await` dans TypeScript :

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

Si vous utilisez des infrastructures JavaScript supplémentaires incluant leurs propres tests ou gestionnaire d’initialisation, il convient *généralement* de les placer dans la réponse à `Office.onReady()`. Par exemple, la fonction de [JQuery's](https://jquery.com) `$(document).ready()` sera référencée comme suit :

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Toutefois, il existe des exceptions à cette pratique. Par exemple, supposons que vous vouliez ouvrir votre module complémentaire dans un navigateur (au lieu d'en charger une version dans un hôte Office) pour pouvoir déboguer votre interface utilisateur avec les outils de navigateur. Dans la mesure où Office.js ne se charge pas dans le navigateur, `onReady` et le `$(document).ready` ne seront pas exécutés si il est appelé dans le `onReady` Office. Une autre exception : vous voulez qu'un indicateur de progression apparaisse dans le volet Office pendant le chargement du module complémentaire. Dans ce scénario, votre code doit appeler la jQuery `ready` et utiliser son rappel pour afficher l’indicateur de progression. Puis le rappel du `onReady` Office peut remplacer l’indicateur de progression avec l’interface utilisateur final. 

### <a name="initialize-with-officeinitialize"></a>Initialiser avec Office.initialize

Un événement Initialiser se déclenche lorsque la bibliothèque Office.js est entièrement chargée et prête pour l’interaction utilisateur. Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémentera votre logique d’initialisation. L'exemple suivant montre comment vérifier que la version utilisateur d’Excel prend en charge toutes les API que le module complémentaire peut appeler.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Si vous utilisez des infrastructures JavaScript supplémentaires incluant leurs propres tests ou gestionnaire d’initialisation, il convient *généralement* de les placer dans l'événement `Office.initialize`. (Mais les exceptions précédemment décrites dans la section **initialisation avec Office.onReady()** s’appliquent dans ce cas également.) Par exemple, la fonction `$(document).ready()` [de JQuery](https://jquery.com) serait référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Pour le volet Office et le contenu des compléments, `Office.initialize` fournit un paramètre _reason_ supplémentaire. Ce paramètre spécifie la manière dont un module complémentaire a été ajouté au document actif. Vous pouvez utiliser cette méthode pour appliquer des logiques différentes lorsqu’un complément est inséré pour la première fois et lorsqu'il existait déjà dans le document.

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

Pour plus d’informations, consultez les sections [événement Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) et [énumération InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principales différences entre Office.initialize et Office.onReady

- Vous pouvez attribuer qu’un seul gestionnaire à `Office.initialize` et il n'est appelé qu’une seule fois par l’infrastructure Office ; mais vous pouvez appeler `Office.onReady()` à différents emplacements dans votre code et utiliser différents rappels. Par exemple, votre code peut appeler `Office.onReady()` dès que votre script personnalisé se charge avec un rappel qui exécute une logique d’initialisation ; et votre code peut avoir également un bouton dans le volet Office, dont le script appelle `Office.onReady()` avec un autre rappel. Dans ce cas, le deuxième rappel s’exécute lorsque le bouton est activé.

- L'événement `Office.initialize` est déclenché à la fin du processus interne au cours duquel Office.js s'initialise. Et il se déclenche *immédiatement* après la fin du processus interne. Si le code dans lequel vous affectez un gestionnaire à l’événement s’exécute trop longtemps après que l’événement se soit déclenché, votre gestionnaire ne s’exécute pas. Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du module complémentaire pour charger les fichiers polyfill après le chargement des Office.js mais avant de charger votre code JavaScript personnalisé. Au moment où votre script est chargé et affecte le gestionnaire, l’événement Initialiser a déjà été exécuté. Mais il n'est jamais « trop tard » pour appeler `Office.onReady()`. Si l’événement Initialiser a déjà eu lieu, le rappel s’exécute immédiatement.

> [!NOTE]
> Même si vous n’avez aucune logique de démarrage, il est conseillé d’appeler `Office.onReady()` ou d'assigner une fonction vide à `Office.initialize` lors du chargement de votre complément JavaScript, car certaines combinaisons d’hôte et de plateforme Office ne chargent pas le volet Office tant que l'un de ces deux événements ne s'est pas produit. Les lignes suivantes illustrent les deux méthodes pour ce faire :
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modèle objet JavaScript Office

Une fois initialisé, le complémentpeut interagir avec l'hôte (par exemple Excel, Outlook). La page sur le [modèle objet de l'API JavaScript Office](office-javascript-api-object-model.md) contient plus de détails sur les habitudes d'utilisation spécifiques. Il existe également une documentation de référence détaillée à la fois pour les [API partagées](https://dev.office.com/reference/add-ins/javascript-api-for-office) et les hôtes spécifiques.

## <a name="api-support-matrix"></a>Matrice de prise en charge d’API

Ce tableau récapitule l’API et les fonctionnalités prises en charge dans les types de complément (contenu, volet Office et Outlook), ainsi que les applications Office qui peuvent les héberger lorsque vous indiquez les applications hôte Office prises en charge par votre complément à l’aide du [schéma de manifeste de complément 1.1 et des fonctionnalités prises en charge par la version 1.1 de l’interface API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nom de l’hôte**|Base de données|Classeur|Boîte aux lettres|Présentation|Document|Projet|
||**Applications hôtes** **prises en charge**|Applications web Access|Excel,<br/>Excel Online|Outlook,<br/>Application web Outlook,<br/>OWA pour les périphériques|PowerPoint,<br/>PowerPoint Online|Word|Projet|
|**Types de compléments pris en charge**|Contenu|v|v||v|||
||Volet de tâches||v||v|v|v|
||Outlook|||O||||
|**Fonctionnalités d’API prises en charge**|Lecture/écriture de texte||v||v|v|v<br/>(En lecture seule)|
||Lecture/Écriture de matrice||v|||v||
||Lecture/écriture de tableau||v|||v||
||Lecture/écriture HTML|||||v||
||Lecture/Écriture<br/>Office Open XML|||||v||
||Lecture des propriétés de tâche, de ressource, de vue et de champ||||||v|
||Événements modifiés de sélection||v|||v||
||Obtention de l’ensemble du document||||v|v||
||Liaisons et événements de liaison|v<br/>(Liaisons de tableau complètes et partielles uniquement)|v|||v||
||Lecture/écriture des parties XML personnalisées|||||v||
||Faire persister les données d’état de complément (paramètres)|v<br/>(Par complément hôte)|v<br/>(Par document)|v<br/>(Par boîte aux lettres)|v<br/>(Par document)|v<br/>(Par document)||
||Événements modifiés de paramètres|v|v||v|v||
||Obtention du mode de vue active<br/>et affichage des événements modifiés||||v|||
||Accès à des emplacements<br/>dans le document||v||v|v||
||Activation en fonction du contexte<br/>à l’aide de règles et de RegEx|||v||||
||Lecture des propriétés d’élément|||v||||
||Lecture de profil utilisateur|||v||||
||Obtention des pièces jointes|||v||||
||Obtention du jeton d’identité d’utilisateur|||v||||
||Appel des services web Exchange|||v||||
