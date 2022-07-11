---
title: Conditions particulières pour les compléments sur iPad
description: Découvrez quelques exigences pour la création d’un complément Office qui s’exécute sur un iPad.
ms.date: 09/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 17df8855a987bd44e657f6ddfdec9925a979449a
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712992"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Conditions particulières pour les compléments sur iPad

Si votre complément utilise uniquement les API Office prises en charge sur l’iPad, les clients peuvent l’installer sur des iPad. (Pour plus d’informations, consultez [Spécifier les applications Office et les exigences d’API](specify-office-hosts-and-api-requirements.md) .) *Si le complément est commercialisé via [AppSource](https://appsource.microsoft.com)*, vous devez suivre certaines pratiques pour les compléments qui peuvent être installés sur des iPad, en plus des [meilleures pratiques qui s’appliquent à tous les compléments Office](../concepts/add-in-development-best-practices.md).

Le tableau suivant répertorie les tâches à effectuer.

> [!NOTE]
> Pour plus d’informations sur la conception de compléments Outlook qui s’affichent correctement et fonctionnent bien sur Outlook Mobile, consultez [Compléments pour Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tâche|Description|Ressources|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Mettre à jour la version du manifeste et de l’API](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Appliquer les meilleures pratiques de conception iOS.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.| Voir la note ci-dessous. |
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de certification 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Rendez votre commerce de complément gratuit sur l’iPad.|Lorsqu’il s’exécute sur l’iPad, votre complément doit être exempt d’achats intégrés à l’application, d’offres d’essai, d’interface utilisateur qui vise à proposer une version non gratuite ou de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres contenus, applications ou compléments. Vos pages Politique de confidentialité et Conditions d’utilisation doivent également être exemptes de liens d’interface utilisateur commerciale ou AppSource.|[Stratégie de certification 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Votre complément peut toujours avoir du commerce sur d’autres plateformes. Pour ce faire, testez la propriété [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member) et supprimez tout commerce lorsqu’elle retourne `false`.|
|Envoyez votre complément à AppSource.|Dans l’Espace partenaires, dans la page **d’installation du produit** , activez la case à cocher **Rendre mon produit disponible sur iOS et Android (le cas échéant)** et indiquez votre ID de développeur Apple dans les paramètres du compte. Passez en revue le [Contrat du fournisseur d’applications](https://go.microsoft.com/fwlink/?linkid=715691) pour vous assurer que vous comprenez les termes.|[Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Votre complément peut servir une autre interface utilisateur en fonction de l’appareil sur lequel il s’exécute. Pour détecter si votre complément s’exécute sur un iPad, vous pouvez utiliser les API suivantes.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member)
>
> Sur un iPad, `touchEnabled` retourne `true` et `commerceAllowed` retourne `false`.
>
> Pour plus d’informations sur les meilleures pratiques de conception de l’interface utilisateur pour iPad, consultez [Conception pour iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/).

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Meilleures pratiques pour le développement de compléments Office qui peuvent s’exécuter sur iPad

Appliquez les meilleures pratiques suivantes pour le développement de compléments qui s’exécutent sur iPad.

-  **Développez et déboguez le complément sur Windows ou Mac et chargez-le sur un iPad.**

    Vous ne pouvez pas développer le complément directement sur un iPad, mais vous pouvez le développer et le déboguer sur un ordinateur Windows ou Mac et le charger sur un iPad à des fins de test. Étant donné qu’un complément qui s’exécute dans Office sur iOS ou Mac prend en charge les mêmes API qu’un complément s’exécutant dans Office sur Windows, le code de votre complément doit s’exécuter de la même façon sur ces plateformes. Pour plus d’informations, consultez [Test et débogage des compléments Office](../testing/test-debug-office-add-ins.md) et [chargement indépendant des compléments Office sur iPad à des fins de test](../testing/sideload-an-office-add-in-on-ipad.md).

-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**

    Lorsque vous spécifiez des exigences d’API dans le manifeste de votre complément, Office détermine si l’application cliente Office prend en charge ces membres d’API. Si les membres de l’API sont disponibles dans l’application, votre complément sera disponible. Vous pouvez également effectuer une vérification du runtime pour déterminer si une méthode est disponible dans l’application avant de l’utiliser dans votre complément. Les vérifications d’exécution vérifient que votre complément est toujours disponible dans l’application et fournissent des fonctionnalités supplémentaires si les méthodes sont disponibles. Pour plus d’informations, consultez [Spécifier les applications Office et les exigences de l’API](specify-office-hosts-and-api-requirements.md).
