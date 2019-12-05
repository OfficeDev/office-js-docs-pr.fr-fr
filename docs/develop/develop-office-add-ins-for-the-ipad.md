---
title: Développer des compléments Office pour iPad
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 3fbe065e111519f81c39d2255b452eab9491fa9d
ms.sourcegitcommit: 960ceaf6776ec3ed41a8f5b7bf70b3c95c43386a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2019
ms.locfileid: "39830187"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>Développer des compléments Office pour iPad


Le tableau suivant répertorie les tâches à effectuer pour développer un complément Office à exécuter dans Office sur iPad.


|**Tâche**|**Description**|**Ressources**|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Mettre à jour la version du manifeste et de l’API](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Appliquez les méthodes recommandées pour concevoir une interface utilisateur.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.|[Concevoir pour iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Appliquez les méthodes recommandées pour concevoir un complément.|Assurez-vous que votre complément offre une valeur claire, une expérience conviviale et des performances optimales.|[Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)|
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de validation 10.8](/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|Proposez un commerce de complément gratuit.|Votre complément ne doit pas comporter de services payants, d’offres d’essai, une interface utilisateur destinée à inciter à la vente, ni de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres contenus, applications ou compléments. Vos pages Politique de confidentialité et Conditions d’utilisation ne doivent pas non plus comporter de liens vers une interface utilisateur commerciale ou AppSource.|[Stratégie de validation 3.4](/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|Envoyez à nouveau votre complément à AppSource.|Dans le centre de partenaires, dans la page de **configuration du produit** , activez la case à cocher **rendre mon produit disponible sur iOS et Android (le cas échéant)** , puis indiquez votre ID de développeur Apple dans paramètres du compte. Consultez le [contrat de fournisseur d’applications](https://go.microsoft.com/fwlink/?linkid=715691) pour vous assurer que vous comprenez les termes.|[Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-appsource-via-partner-center)|

Votre complément peut rester en l’état pour les applications Office exécutées sur d’autres plateformes. Vous pouvez également proposer une interface utilisateur différente en fonction du navigateur ou de l’appareil qui utilise votre complément. Pour savoir si votre complément est exécuté sur un iPad, vous pouvez utiliser les API suivantes :
- var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)


## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Meilleures pratiques en matière de développement de compléments Office pour iOS et Mac

Appliquez les meilleures pratiques suivantes pour développer des compléments pour iOS :


-  **Utilisez Visual Studio pour développer votre complément.**

    Si vous développez votre complément avec Visual Studio, vous pouvez [définir des points d’arrêt et déboguer son code](../develop/create-and-debug-office-add-ins-in-visual-studio.md) dans une application hôte Office s’exécutant sous Windows, avant de charger votre complément sur iPad ou Mac. Étant donné qu’un complément exécuté dans Office sur iOS ou Mac prend en charge les mêmes API qu’un complément s’exécutant dans Office sur Windows, le code de votre complément doit s’exécuter de la même manière sur les deux plateformes.

-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**

    Lorsque vous spécifiez des conditions requises d’API dans le manifeste de votre complément, Office détermine si l’application hôte prend en charge ces membres de l’API. Si les membres de l’API sont disponibles dans l’hôte, votre complément sera alors disponible dans cette application hôte. Par ailleurs, vous pouvez effectuer une vérification à l’exécution pour déterminer si une méthode est disponible dans l’hôte avant de l’utiliser dans votre complément. Les vérifications à l’exécution garantissent que votre complément est toujours disponible dans l’hôte et qu’il fournit des fonctionnalités supplémentaires si les méthodes sont disponibles. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](specify-office-hosts-and-api-requirements.md).

Pour plus d’informations sur des pratiques plus générales en matière de développement de compléments, consultez la rubrique [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md).


## <a name="see-also"></a>Voir aussi

- [Charger une version test d’un complément Office sur iPad ou Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Débogage des compléments Office sur iPad et Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
