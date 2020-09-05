---
title: Exigences spéciales pour les compléments sur l’iPad
description: Découvrez les conditions requises pour la création d’un complément Office qui s’exécute sur un iPad.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 25ac5767db3301352e1921411af833957c4644d0
ms.sourcegitcommit: 10463841a977e9b8415362a3ae91b0ae5eebbf89
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/04/2020
ms.locfileid: "47399570"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Exigences spéciales pour les compléments sur l’iPad

Si votre complément utilise uniquement les API Office prises en charge sur l’iPad, les clients peuvent l’installer sur iPad. (Pour plus d’informations, voir [spécifier les applications Office et les conditions requises](specify-office-hosts-and-api-requirements.md) pour les API.) *Si le complément est commercialisé via [AppSource](https://appsource.microsoft.com)*, vous devez suivre certaines pratiques pour les compléments qui peuvent être installés sur iPad, ainsi [que les meilleures pratiques qui s’appliquent à tous les compléments Office](../concepts/add-in-development-best-practices.md).

Le tableau suivant répertorie les tâches à effectuer.

> [!NOTE]
> Pour plus d’informations sur la conception de compléments Outlook qui s’affichent correctement et fonctionnent bien sur Outlook Mobile, consultez la rubrique [compléments pour Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tâche|Description|Ressources|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Mettre à jour la version du manifeste et de l’API](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Appliquer les meilleures pratiques de conception d’iOS.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.| Consultez la remarque ci-dessous. |
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de certification 1120,2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Rendez votre commerce de complément gratuit sur l’iPad.|Lorsqu’il est en cours d’exécution sur l’iPad, votre complément doit être exempt d’achats dans l’application, d’offres d’essai, d’interface utilisateur qui visent à proposer une promotion à une version non gratuite ou de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres types de contenu, d’applications ou de compléments. Vos pages politique de confidentialité et conditions d’utilisation doivent également être dépourvues de liens d’interface utilisateur Commerce Server ou de liens AppSource.|[Stratégie de certification 1100,3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Votre complément peut toujours avoir du commerce sur d’autres plateformes. Pour ce faire, testez la propriété [Office. Context. commerceAllowed](/javascript/api/office/office.context#commerceallowed) et supprimez tous les commerciaux quand elle renvoie `false` .|
|Envoyez votre complément à AppSource.|Dans le centre de partenaires, dans la page de **configuration du produit** , activez la case à cocher **rendre mon produit disponible sur iOS et Android (le cas échéant)** , puis indiquez votre ID de développeur Apple dans paramètres du compte. Consultez le [contrat de fournisseur d’applications](https://go.microsoft.com/fwlink/?linkid=715691) pour vous assurer que vous comprenez les termes.|[Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Votre complément peut servir une autre interface utilisateur basée sur l’appareil sur lequel il s’exécute. Pour détecter si votre complément est en cours d’exécution sur un iPad, vous pouvez utiliser les API suivantes.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> Sur un iPad, `touchEnabled` renvoie `true` et `commerceAllowed` renvoie `false` .
>
> Pour plus d’informations sur les meilleures pratiques de conception de l’interface utilisateur pour iPad, consultez la rubrique [Designing for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/).

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Meilleures pratiques pour le développement de compléments Office pouvant s’exécuter sur iPad

Appliquez les meilleures pratiques suivantes pour développer des compléments qui s’exécutent sur iPad.

-  **Développez et déboguez le complément sur Windows ou Mac et chargement-le sur un iPad.**

    Vous ne pouvez pas développer le complément directement sur un iPad, mais vous pouvez le développer et le déboguer sur un ordinateur Windows ou Mac et l’chargement sur un iPad pour le tester. Étant donné qu’un complément exécuté dans Office sur iOS ou Mac prend en charge les mêmes API qu’un complément s’exécutant dans Office sur Windows, le code de votre complément doit s’exécuter de la même manière sur ces plateformes. Pour plus d’informations, consultez la rubrique [tester et déboguer des compléments Office](../testing/test-debug-office-add-ins.md) et [chargement des compléments Office sur iPad et Mac à des fins de test](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**

    Lorsque vous spécifiez les conditions requises de l’API dans le manifeste de votre complément, Office détermine si l’application cliente Office prend en charge ces membres d’API. Si les membres de l’API sont disponibles dans l’application, votre complément sera disponible. Vous pouvez également effectuer une vérification à l’exécution pour déterminer si une méthode est disponible dans l’application avant de l’utiliser dans votre complément. Les vérifications à l’exécution garantissent que votre complément est toujours disponible dans l’application et qu’il fournit des fonctionnalités supplémentaires si les méthodes sont disponibles. Pour plus d’informations, voir [spécifier les applications Office et les conditions requises](specify-office-hosts-and-api-requirements.md)de l’API.
