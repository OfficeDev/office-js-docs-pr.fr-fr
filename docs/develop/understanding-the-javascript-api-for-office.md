---
title: Présentation de l’API JavaScript pour Office
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 14de5d8bab791d0954179c21163ba0a08824b834
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458103"
---
# <a name="understanding-the-javascript-api-for-office"></a>Présentation de l’API JavaScript pour Office

Cet article fournit des informations sur l’API JavaScript pour Office et son utilisation. Pour obtenir des informations de référence, voir [API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office). Pour plus d’informations sur la mise à jour des fichiers de projet Visual Studio vers la version la plus récente de l’API JavaScript pour Office, voir [Mettre à jour la version de votre API JavaScript pour Office et les fichiers de schéma manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Référencer la bibliothèque de l’interface API JavaScript pour Office dans votre complément

La bibliothèque de l’[interface API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Cette opération permet de télécharger et de mettre en cache les fichiers de l’interface API JavaScript pour Office lors du premier chargement de votre complément pour garantir qu’elle utilise l’implémentation d’Office.js la plus récente et les fichiers .js qui lui sont associés pour la version indiquée.

Pour obtenir plus d’informations sur le CDN Office.js et la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Initialisation de votre complément

**S’applique à :** tous les types de complément

Les compléments Office ont souvent une logique de démarrage pour effectuer des actions telles que :

- Vérifiez que version de l’utilisateur d’Office prendra en charge tous les API Office que votre code appelle.

- Vérifiez l’existence de certains artefacts tels que des feuille de calcul avec un nom spécifique.

- Avertir l’utilisateur pour sélectionner certaines cellules dans Excel, puis insérer un graphique initialisé avec ces valeurs sélectionnées.

- Établir des liaisons.

- Utilisez la boîte de dialogue Office API pour inviter l’utilisateur pour les valeurs de paramètres des compléments par défaut.

Mais votre code démarrage ne doit pas appeler n’importe quel APIs Office.js jusqu'à ce que la bibliothèque ne soit entièrement chargée. Il existe deux manières pour votre code de s’assurer que la bibliothèque est chargée. Ceci est décrit en détail dans les sections ci-après : 

- [Initialiser avec Office.onReady()](#initialize-with-officeonready)
- [Initialiser avec Office.initialize](#initialize-with-officeinitialize)

Pour plus d’informations sur les différences entre ces techniques, voir [Différences majeures entre Office.initialize et Office.onReady()](#major-differences-between-officeinitialize-and-officeonready). Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et environnement d’exécution](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>Initialiser avec Office.onReady()

`Office.onReady()` est une méthode asynchrone qui renvoie un objet Promise tandis qu’il vérifie si la bibliothèque Office.js est entièrement chargée. Uniquement lorsque la bibliothèque est chargée, cela résout la promesse sous forme d’objet qui spécifie l’application Office hôte avec une`Office.HostType` valeur enum (`Excel`, `Word`, etc.) et la plateforme avec une`Office.PlatformType` valeur enum (`PC`, `Mac`, `OfficeOnline`, etc..). Si la bibliothèque est déjà chargée quand `Office.onReady()` est appelée, la promesse se résout immédiatement.

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

Toutefois, il existe des exceptions à cette pratique. Par exemple, supposons que vous voulez ouvrir votre complément dans un navigateur (au lieu de le charger dans un hôte Office) afin de déboguer votre interface utilisateur avec les outils de navigateur. Étant donné que Office.js ne sera pas chargé dans le navigateur, `onReady` ne s’exécutera pas et le `$(document).ready` ne s’exécutera pas si cette opération est appelée à l’intérieur d’Office `onReady`. Une autre exception : vous souhaitez qu’un indicateur de progression s’affiche dans le volet Office tandis que le complément se charge. Dans ce scénario, votre code doit appeler la jQuery `ready` et utiliser le rappel pour afficher l’indicateur de progression. Puis le rappel `onReady` Office peut remplacer l’indicateur de progression par l’interface utilisateur final. 

### <a name="initialize-with-officeinitialize"></a>Initialiser avec Office.initialize

Un événement initialisé se déclenche lorsque la bibliothèque Office.js est entièrement chargée et prête pour une interaction avec l’utilisateur. Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente votre logique d’initialisation. L’exemple suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être*habituellement* placées dans l’événement`Office.initialize`. (Mais les exceptions décrites dans la section **initialiser avec Office.onReady()** précédente s’appliquent dans ce cas également.) Par exemple, la fonction[ JQuery](https://jquery.com) `$(document).ready()` serait référencée comme suit :

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

Pour plus d’informations, consultez les pages relatives à l’[Événement Office.initialize](https://docs.microsoft.com/javascript/api/office) et à l’[Énumération InitializationReason](https://docs.microsoft.com/javascript/api/office/office.initializationreason).

> [!NOTE]
> Pour l’instant, vous devez définir `Office.Initialize`, peu importe si `Office.onReady()` est également appelé. Si vous ne vous servez pas de `Office.Initialize`, vous pouvez le définir sur une fonction vide comme illustré dans l’exemple suivant.
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Principales différences entre Office.initialize et Office.onReady

- Vous ne pouvez assigner qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure d’Office, mais vous pouvez appeler `Office.onReady()`à plusieurs endroits dans votre code et utiliser des rappels différents. Par exemple, votre code pourrait appeler `Office.onReady()` dès que votre script personnalisé charge avec un rappel qui exécute la logique d’initialisation ; et votre code peut également comporter un bouton dans le volet Office dont le script appelle `Office.onReady()` avec un rappel différent. Si c’est le cas, le deuxième rappel s’exécute quand l’utilisateur clique sur le bouton.

- L’événement`Office.initialize` se déclenche à la fin du processus interne dans lequel Office.js s’initialise lui-même. Et il se déclenche *immédiatement* après la fin du processus interne. Si le code dans lequel vous attribuez un gestionnaire à l’événement s’exécute trop longtemps après le déclenchement de l’événement, votre gestionnaire ne s’exécutera pas. Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du complément pour charger les fichiers polyfill une fois que le serveur charge Office.js mais avant que le serveur ne charge votre code JavaScript personnalisé. Le temps que votre script se charge et affecte le Gestionnaire, l’événement initialisé s’est déjà produit. Mais il n’est jamais « trop tard » pour appeler `Office.onReady()`. Si l’événement initialisé s’est déjà produit, le rappel s’exécute immédiatement.

> [!NOTE]
> Même si vous n’avez aucune logique de démarrage, attribuez une fonction vide à `Office.initialize` lorsque votre complément JavaScript se charge, comme illustré dans l’exemple suivant. Certaines combinaisons de plateforme et hôte Office ne chargent pas le volet Office jusqu'à ce que l’événement initialisé ne se déclenche et que la fonction gestionnaire d’événements spécifiée ne s’exécute.
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Modèle d’objet API JavaScript Office

Une fois initialisé, le complément peut interagir avec l’hôte (par exemple, Excel, Outlook). La page [Modèle objet API JavaScript Office](office-javascript-api-object-model.md) comporte plus d’informations sur les modèles d’utilisation spécifiques. Il existe également une documentation de référence détaillée pour les deux[ APIs Communes](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) et spécifiques.

## <a name="api-support-matrix"></a>Matrice de prise en charge d’API

Ce tableau récapitule l’API et les fonctionnalités prises en charge dans les types de complément (contenu, volet Office et Outlook), ainsi que les applications Office qui peuvent les héberger lorsque vous indiquez les applications hôte Office prises en charge par votre complément à l’aide du [schéma de manifeste de complément 1.1 et des fonctionnalités prises en charge par la version 1.1 de l’interface API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nom de l’hôte**|Base de données|Classeur|Boîte aux lettres|Présentation|Document|Projet|
||**Applications hôtes** **prises en charge**|Applications web Access|Excel,<br/>Excel Online|Outlook,<br/>Outlook Web App,<br/>OWA pour les périphériques|PowerPoint,<br/>PowerPoint Online|Word|Project|
|**Types de compléments pris en charge**|Contenu|v|v||v|||
||Volet de tâches||v||v|v|v|
||Outlook|||O||||
|**Fonctionnalités d’API prises en charge**|Lecture/écriture de texte||v||v|v|v<br/>(En lecture seule)|
||Lecture/Écriture de matrice||v|||v||
||Lecture/écriture de tableau||v|||v||
||Lecture/écriture HTML|||||v||
||Lecture/Écriture<br/>Office Open XML|||||v||
||Lecture des propriétés de tâche, de ressource, de vue et de champ||||||v|
||Événements modifiés de sélection||v|||v||
||Obtention de l’ensemble du document||||v|v||
||Liaisons et événements de liaison|v<br/>(Liaisons de tableau complètes et partielles uniquement)|v|||v||
||Lecture/écriture des parties XML personnalisées|||||v||
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
