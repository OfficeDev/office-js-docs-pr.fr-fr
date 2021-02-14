---
title: Conditions particulières pour les compléments sur iPad
description: Découvrez quelques conditions requises pour la création d’un add-in Office qui s’exécute sur un iPad.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: fdb402f4302e7e81589d586fa1ecd5b30d4e515d
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237853"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Conditions particulières pour les compléments sur iPad

Si votre application utilise uniquement des API Office qui sont pris en charge sur l’iPad, les clients peuvent l’installer sur iPad. (Pour plus [d’informations, voir Spécifier](specify-office-hosts-and-api-requirements.md) les applications Office et les conditions requises pour l’API.) Si le complément sera commercialisé via *[AppSource,](https://appsource.microsoft.com)* vous devez suivre certaines pratiques pour les compléments qui peuvent être installés sur iPad, en plus des meilleures pratiques qui s’appliquent à tous les [compléments Office.](../concepts/add-in-development-best-practices.md)

Le tableau suivant répertorie les tâches à effectuer.

> [!NOTE]
> Pour plus d’informations sur la conception de add-ins Outlook qui s’lookent bien et fonctionnent bien sur Outlook Mobile, voir Les [add-ins pour Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tâche|Description|Ressources|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Mettre à jour la version du manifeste et de l’API](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Appliquez les meilleures pratiques en matière de conception iOS.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.| Voir la remarque ci-dessous. |
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de certification 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Rendez votre commerce de votre application gratuit sur iPad.|Lorsqu’il est en cours d’exécution sur iPad, votre application ne doit pas avoir besoin d’achats in-app, d’offres d’essai, d’une interface utilisateur qui vise à la vente à une version non gratuite ou de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres contenus, applications ou modules. Vos pages Politique de confidentialité et Conditions d’utilisation ne doivent pas non plus être des liens vers l’interface utilisateur commerciale ou AppSource.|[Stratégie de certification 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Votre add-in peut toujours avoir des échanges commerciaux sur d’autres plateformes. Pour ce faire, testez [la propriété Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed) et supprimez tout commerce lors de son `false` retour.|
|Soumettez votre add-in dans AppSource.|Dans l’Partner Center, dans la **page** Configuration du produit, cochez la case Rendre mon produit disponible sur iOS et Android (le cas **échéant)** et fournissez votre ID de développeur Apple dans les paramètres du compte. Examinez [le contrat du fournisseur d’applications](https://go.microsoft.com/fwlink/?linkid=715691) pour vous assurer que vous comprenez les termes.|[Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Votre add-in peut servir une autre interface utilisateur en fonction de l’appareil sur qui il s’exécute. Pour détecter si votre application est en cours d’exécution sur un iPad, vous pouvez utiliser les API suivantes.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> Sur un iPad, `touchEnabled` renvoie `true` et renvoie `commerceAllowed` `false` .
>
> Pour plus d’informations sur les meilleures pratiques de conception d’interface utilisateur pour iPad, voir [Conception pour iOS.](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Meilleures pratiques pour développer des applications Office qui peuvent s’exécuter sur iPad

Appliquez les meilleures pratiques suivantes pour développer des applications qui s’exécutent sur iPad.

-  **Développez et déboguer le add-in sur Windows ou Mac et chargez-le de nouveau sur un iPad.**

    Vous ne pouvez pas développer le add-in directement sur un iPad, mais vous pouvez le développer et le déboguer sur un ordinateur Windows ou Mac et le recharger de manière test sur un iPad. Étant donné qu’un application qui s’exécute dans Office sur iOS ou Mac prend en charge les mêmes API qu’un application qui s’exécute dans Office sur Windows, le code de votre application doit s’exécuter de la même manière sur ces plateformes. Pour plus d’informations, voir Tester et [déboguer](../testing/test-debug-office-add-ins.md) des applications Office et chargez une version test des macros supplémentaires Office sur iPad et [Mac.](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**

    Lorsque vous spécifiez des conditions requises pour l’API dans le manifeste de votre application, Office détermine si l’application cliente Office prend en charge ces membres d’API. Si les membres de l’API sont disponibles dans l’application, votre application sera disponible. Vous pouvez également effectuer une vérification à l’runtime pour déterminer si une méthode est disponible dans l’application avant de l’utiliser dans votre application. Les vérifications à l’runtime garantissent que votre complément est toujours disponible dans l’application et fournissent des fonctionnalités supplémentaires si les méthodes sont disponibles. Pour plus d’informations, voir [Spécifier les applications Office et les conditions requises pour les API.](specify-office-hosts-and-api-requirements.md)
