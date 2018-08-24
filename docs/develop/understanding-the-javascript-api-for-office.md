---
title: Présentation de l’API JavaScript pour Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: a9e1e26d4ba94a933ecb98250c19afee90750f5d
ms.sourcegitcommit: 28fc652bded31205e393df9dec3a9dedb4169d78
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/23/2018
ms.locfileid: "22928034"
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

Pour en savoir plus sur le CDN Office.js, y compris sur la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de l’API JavaScript pour la bibliothèque Office à partir de son réseau de distribution de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Initialisation de votre complément

**S’applique à :** tous les types de complément

Office.js fournit un événement d’initialisation qui se déclenche lorsque l’API est entièrement chargée et prête à interagir avec l’utilisateur. Vous pouvez utiliser le gestionnaire d’événements **initialize** afin de mettre en œuvre des scénarios d’initialisation de complément courants, comme inviter l’utilisateur à sélectionner des cellules dans Excel, puis insérer un graphique initialisé avec les valeurs sélectionnées. Vous pouvez également utiliser le gestionnaire d’événements initialize pour initialiser d’autres logiques personnalisées pour votre complément, telles que l’établissement de liaisons, la demande de valeurs de paramètres de complément par défaut, et ainsi de suite.

Voici à quoi ressemblerait l’événement initialize :     

```js
Office.initialize = function () { };
```
Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être placées dans l’événement Office.initialize. Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Toutes les pages au sein d’un complément Office sont nécessaires pour attribuer un gestionnaire d’événements à l’événement initialize, **Office.initialize**. Si vous ne parvenez pas à attribuer un gestionnaire d’événements, votre complément peut générer une erreur lors de son démarrage. En outre, si un utilisateur essaie d’utiliser votre complément avec un client web Office Online, notamment Excel Online, PowerPoint Online ou Outlook Web App, l’exécution du complément échouera. Si vous n’avez pas besoin de code d’initialisation, le corps de la fonction attribuée à **Office.initialize** peut être vide, comme dans le premier exemple ci-dessus.

Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Initialisation du paramètre Reason
Pour les compléments de contenu et du volet Office, Office.initialize fournit un paramètre _reason_ supplémentaire. Ce paramètre peut être utilisé pour savoir comment un complément a été ajouté au document actif. Vous pouvez l’utiliser pour fournir une logique différente quand un complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document. 

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
Pour plus d’informations, consultez les pages relatives à l’[événement Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) et à l’[énumération InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration) 

## <a name="office-javascript-api-object-model"></a>Modèle objet JavaScript Office

Une fois initialisé, le complément peut interagir avec l’hôte (par exemple, Excel, Outlook). La page [Modèle d’objet API JavaScript pour Office](office-javascript-api-object-model.md) a plus de détails sur les modèles d’utilisations spécifiques. Il existe également une documentation de référence détaillée pour les [API partagées](https://dev.office.com/reference/add-ins/javascript-api-for-office) et les hôtes spécifiques.

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
