---
title: Présentation de l’API JavaScript pour Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e9d9efdda5e237ab076d22d50b1f7ded5e075845
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505947"
---
# <a name="understanding-the-javascript-api-for-office"></a>Présentation de l’API JavaScript pour Office

Cet article fournit des informations sur l’API JavaScript pour Office et son utilisation. Pour obtenir des informations de référence, voir [API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). Pour plus d’informations sur la mise à jour des fichiers de projet Visual Studio vers la version la plus récente de l’API JavaScript pour Office, voir [Mettre à jour la version de votre API JavaScript pour Office et les fichiers de schéma manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Référence à la bibliothèque de l’interface API JavaScript pour Office dans votre complément

La bibliothèque de l’[interface API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Cette opération permet de télécharger et de mettre en cache les fichiers de l’interface API JavaScript pour Office lors du premier chargement de votre complément pour garantir qu’elle utilise l’implémentation d’Office.js la plus récente et les fichiers .js qui lui sont associés pour la version indiquée.

Pour en savoir plus sur le CDN Office.js, y compris sur la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de la bibliothèque de l'interface API JavaScript pour Office à partir de son réseau de diffusion de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Initialisation de votre complément

**S’applique à :** tous les types de complément

Les compléments Office ont souvent une logique de démarrage pour effectuer des tâches telles que :

- Vérifier que la version utilisateur d’Office prendra en charge toutes les API Office appelées par votre code.

- Vérifier l’existence de certains artifacts, tels qu'une feuille de calcul avec un nom spécifique.

- Inviter l’utilisateur à sélectionner des cellules dans Excel, puis d’insérer un graphique initialisé avec ces valeurs sélectionnées.

- Établir des liaisons.

- Depuis la boîte de dialogue API Office, inviter l’utilisateur de paramétrer les valeurs de paramètres de module complémentaire par défaut.

Mais votre code de démarrage ne doit pas appeler n’importe quel APIs Office.js jusqu'à ce que la bibliothèque est entièrement chargée. Il existe deux manières que votre code peut faire en sorte que la bibliothèque est chargée. Ils sont décrits dans les sections ci-dessous. Nous vous conseillons d’utiliser la technique plus récente, plus souple, l’appel `Office.onReady()`. La technique antérieure, affectation d’un gestionnaire à `Office.initialize`, est toujours prise en charge. Voir aussi les [principales différences entre Office.initialize et Office.onReady()](#major-differences-between-office-initialize-and-office-onready).

Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>Initialisation avec Office.onReady()

`Office.onReady()` est une méthode asynchrone qui renvoie un objet promesse pendant qu’il vérifie si la bibliothèque Office.js est entièrement chargée. Lorsque, et uniquement lorsque la bibliothèque est chargée, elle résout promesse en tant qu’objet qui spécifie l’application Office hôte avec une `Office.HostType` valeur enum (`Excel`, `Word`, etc.) et la plateforme avec un `Office.PlatformType` valeur enum (`PC`, `Mac`, `OfficeOnline`, etc.). Si la bibliothèque est déjà chargée lorsque `Office.onReady()` est appelée, promesse résout immédiatement.

Pour appeler `Office.onReady()` consiste à passer une méthode de rappel. Voici un exemple :

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

Sinon, vous pouvez enchaîner une `then()` l’appel de méthode `Office.onReady()`, au lieu de passer d’un rappel. Par exemple, le code suivant vérifie que la version utilisateur d’Excel prend en charge toutes les API que le complément peut appeler.

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

|||UNTRANSLATED_CONTENT_START|||If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:|||UNTRANSLATED_CONTENT_END|||

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Toutefois, il existe des exceptions à cette application pratique. Par exemple, supposons que vous voulez ouvrir votre complément dans un navigateur (au lieu de sideload dans un hôte Office) pour pouvoir déboguer votre interface utilisateur avec les outils de navigateur. Dans la mesure où Office.js ne se charge pas dans le navigateur, `onReady` ne seront pas exécutés et le `$(document).ready` ne seront pas exécutés si elle est appelée à l’intérieur du bureau `onReady`. Une autre exception : vous voulez un indicateur de progression apparaissent dans le volet de tâches pendant le charge de la macro complémentaire. Dans ce scénario, votre code doit appeler la jQuery `ready` et de rappel pour afficher l’indicateur de progression. Puis le Office `onReady`du rappel peut remplacer l’indicateur de progression avec l’interface utilisateur final. 

### <a name="initialize-with-officeinitialize"></a>Initialiser avec Office.initialize

Un événement initialize se déclenche lorsque la bibliothèque Office.js est entièrement chargé et prêt pour l’interaction utilisateur. Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente la logique d’initialisation. Voici un exemple qui montre comment pour vérifier que la version utilisateur d’Excel prend en charge toutes les API que le complément peut appeler.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Si vous utilisez des infrastructures JavaScript supplémentaires incluant les tests ou leur propre gestionnaire d’initialisation, il doivent *généralement* être placé dans le `Office.initialize` événement. (Mais les exceptions décrites dans la section **initialiser avec Office.onReady()** précédemment s’appliquent dans ce cas également.) Par exemple, [de JQuery](https://jquery.com) `$(document).ready()` fonction serait référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

|||UNTRANSLATED_CONTENT_START|||For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter specifies how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.|||UNTRANSLATED_CONTENT_END|||

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

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principales différences entre Office.initialize et Office.onReady

- Vous pouvez attribuer qu’un seul gestionnaire à `Office.initialize` et elle est appelée, qu’une seule fois par l’infrastructure Office ; mais vous pouvez appeler `Office.onReady()` à différents emplacements dans votre code et utiliser les différents rappels. Par exemple, votre code peut appeler `Office.onReady()` dès que la charge de votre script personnalisé avec un rappel qui s’exécute une logique d’initialisation ; et votre code peut avoir également un bouton dans le volet Office, dont le script appelle `Office.onReady()` avec un autre rappel. Dans ce cas, le deuxième rappel s’exécute lorsque le bouton est activé.

- Le `Office.initialize` événement est déclenché à la fin du processus interne qui initialise Office.js lui-même. Et il déclenche *immédiatement* après la fin du processus interne. Si le code dans lequel vous affectez un gestionnaire à l’événement s’exécute trop long après l’événement se déclenche, votre gestionnaire ne s’exécute. Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du complément pour charger les fichiers polyfill après le chargement des Office.js mais avant de charger votre code JavaScript personnalisé. Au moment où votre script charge et affecte le gestionnaire, l’événement initialize a déjà eu lieu. Mais il est jamais « trop en retard » pour appeler `Office.onReady()`. Si l’événement initialize a déjà eu lieu, le rappel s’exécute immédiatement.

> [!NOTE]
> Même si vous n’avez aucune logique de démarrage, il est conseillé d’appeler `Office.onReady()` ou d'assigner une fonction vide à `Office.initialize` lors du chargement de votre complément JavaScript, car certaines combinaisons d’hôte et de plateforme Office ne chargent pas le volet Office tant que l'un de ces deux événements ne s'est pas produit.
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modèle objet JavaScript Office

Une fois initialisé, le complément peut interagir avec l’hôte (par exemple, Excel, Outlook). La page [Modèle d’objet API JavaScript pour Office](office-javascript-api-object-model.md) a plus de détails sur les modèles d’utilisations spécifiques. Il existe également une documentation de référence détaillée pour les [API partagées](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) et les hôtes spécifiques.

## <a name="api-support-matrix"></a>Matrice de prise en charge d’API

Ce tableau récapitule l’API et les fonctionnalités prises en charge dans les types de complément (contenu, volet Office et Outlook), ainsi que les applications Office qui peuvent les héberger lorsque vous indiquez les applications hôte Office prises en charge par votre complément à l’aide du [schéma de manifeste de complément 1.1 et des fonctionnalités prises en charge par la version 1.1 de l’interface API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nom de l’hôte**|Base de données|Manuel|Boîte aux lettres|Présentation|Document|Projet|
||**Applications hôtes** **prises en charge**|applications web Access|Excel,<br/>Excel Online|Outlook,<br/>Application web Outlook,<br/>OWA pour les périphériques|PowerPoint,<br/>PowerPoint Online|Word|Projet|
|**Types de compléments pris en charge**|Contenu|v|v||v|||
||Volet Office||v||v|v|v|
||Outlook|||v||||
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
