---
title: Conditions particulières pour les compléments sur iPad
description: Découvrez quelques conditions requises pour créer un Office qui s’exécute sur une iPad.
ms.date: 09/03/2020
ms.localizationpriority: medium
---


# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Conditions particulières pour les compléments sur iPad

Si votre application utilise uniquement Office API qui sont pris en charge sur le iPad, les clients peuvent l’installer sur iPad. (Pour plus [d’informations, voir Spécifier Office applications et les exigences d’API](specify-office-hosts-and-api-requirements.md).) Si le complément est commercialisé via *[AppSource](https://appsource.microsoft.com)*, vous devez suivre certaines pratiques pour les compléments qui peuvent être installés sur iPad, en plus des meilleures pratiques qui s’appliquent à tous les [compléments Office](../concepts/add-in-development-best-practices.md).

Le tableau suivant répertorie les tâches à effectuer.

> [!NOTE]
> Pour plus d’informations sur Outlook des Outlook qui s’lookent et fonctionnent bien sur Outlook Mobile, voir Les Outlook [Mobile](../outlook/outlook-mobile-addins.md).

|Tâche|Description|Ressources|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Mettre à jour la version du manifeste et de l’API](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Appliquez les meilleures pratiques en matière de conception iOS.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.| Voir la remarque ci-dessous. |
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de certification 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Rendez gratuitement votre commerce de iPad.|Lorsqu’il est en cours d’exécution sur le iPad, votre add-in ne doit pas avoir besoin d’achats in-app, d’offres d’essai, d’interface utilisateur qui vise à la vente vers une version non gratuite ou de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres contenus, applications ou modules. Vos pages Politique de confidentialité et Conditions d’utilisation ne doivent pas non plus être des liens vers l’interface utilisateur commerciale ou AppSource.|[Stratégie de certification 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Votre add-in peut toujours avoir des échanges commerciaux sur d’autres plateformes. Pour ce faire, testez [la propriété Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member) et supprimez tout commerce lors de son retour`false`.|
|Soumettez votre add-in dans AppSource.|Dans l’Partner Center, dans **la page Configuration** du produit, cochez la case Rendre mon produit disponible sur **iOS et Android (** le cas échéant) et fournissez votre ID de développeur Apple dans les paramètres du compte. Examinez [le contrat du fournisseur d’applications](https://go.microsoft.com/fwlink/?linkid=715691) pour vous assurer que vous comprenez les termes.|[Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Votre add-in peut servir une autre interface utilisateur en fonction de l’appareil sur qui il s’exécute. Pour détecter si votre iPad est en cours d’exécution, vous pouvez utiliser les API suivantes.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member)
>
> Sur une iPad, renvoie `touchEnabled` et `true` `commerceAllowed` renvoie `false`.
>
> Pour plus d’informations sur les meilleures pratiques de conception d’interface utilisateur iPad, voir [Conception pour iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/).

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Meilleures pratiques pour le développement de Office des modules qui peuvent s’exécuter sur iPad

Appliquez les meilleures pratiques suivantes pour le développement de iPad.

-  **Développez et déboguer le add-in sur Windows mac et chargez-le sur un iPad.**

    Vous ne pouvez pas développer le add-in directement sur un iPad, mais vous pouvez le développer et le déboguer sur un ordinateur Windows ou Mac et le recharger de manière test sur un iPad. Étant donné qu’un add-in qui s’exécute dans Office sur iOS ou Mac prend en charge les mêmes API qu’un module de Office sur Windows, le code de votre add-in doit s’exécuter de la même manière sur ces plateformes. Pour plus d’informations, voir Tester et déboguer [des](../testing/test-debug-office-add-ins.md) Office et chargement de version test des Office sur [iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md) pour les tests.

-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**

    Lorsque vous spécifiez des conditions requises pour l’API dans le manifeste de votre Office, l’application cliente Office prend en charge ces membres d’API. Si les membres de l’API sont disponibles dans l’application, votre application sera disponible. Vous pouvez également effectuer une vérification à l’runtime pour déterminer si une méthode est disponible dans l’application avant de l’utiliser dans votre application. Les vérifications à l’runtime garantissent que votre complément est toujours disponible dans l’application et fournissent des fonctionnalités supplémentaires si les méthodes sont disponibles. Pour plus d’informations, voir [Spécifier les Office applications et les api requises](specify-office-hosts-and-api-requirements.md).
