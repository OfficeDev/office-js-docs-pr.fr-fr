---
title: Développer des compléments Office pour iPad
description: ''
ms.date: 01/23/2018
localization_priority: Normal
ms.openlocfilehash: 7aaae96377c58656d82802e44c964f651e7accbc
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388114"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>Développer des compléments Office pour iPad


Le tableau suivant répertorie les tâches à effectuer pour développer un complément Office à exécuter dans Office pour iPad.


|**Tâche**|**Description**|**Ressources**|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Nouveautés de l’API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office)|
|Appliquez les méthodes recommandées pour concevoir une interface utilisateur.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.|[Concevoir pour iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Appliquez les méthodes recommandées pour concevoir un complément.|Assurez-vous que votre complément offre une valeur claire, une expérience conviviale et des performances optimales.|[Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)|
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de validation 10.8](https://docs.microsoft.com/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|Proposez un commerce de complément gratuit.|Votre complément ne doit pas comporter de services payants, d’offres d’essai, une interface utilisateur destinée à inciter à la vente, ni de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres contenus, applications ou compléments. Vos pages Politique de confidentialité et Conditions d’utilisation ne doivent pas non plus comporter de liens vers une interface utilisateur commerciale ou AppSource.|[Stratégie de validation 3.4](https://docs.microsoft.com/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|Envoyez à nouveau votre complément à AppSource.|Dans le tableau de bord vendeur, cochez la case **Ajouter ce complément au catalogue des compléments Office sur iPad**. Indiquez votre ID de développeur Apple dans la case Identifiant Apple. Lisez le [contrat de fournisseur d’application](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm) pour connaître les termes du contrat.|[Mise à disposition de vos solutions sur AppSource et dans Office](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|

Votre complément peut rester en l’état pour les applications Office exécutées sur d’autres plateformes. Vous pouvez également proposer une interface utilisateur différente en fonction du navigateur ou de l’appareil qui utilise votre complément. Pour savoir si votre complément est exécuté sur un iPad, vous pouvez utiliser les API suivantes :
- var isTouchEnabled = [Office.context.touchEnabled](https://docs.microsoft.com/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](https://docs.microsoft.com/javascript/api/office/office.context#commerceallowed)
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Meilleures pratiques en matière de développement de compléments Office pour iOS et Mac

Appliquez les meilleures pratiques suivantes pour développer des compléments pour iOS :


-  **Utilisez Visual Studio pour développer votre complément.**
    
    Si vous développez votre complément avec Visual Studio, vous pouvez [définir des points d’arrêt et déboguer son code](../develop/create-and-debug-office-add-ins-in-visual-studio.md) dans une application hôte Office s’exécutant sous Windows, avant de charger votre complément sur iPad ou Mac. Étant donné qu’un complément qui s’exécute dans Office pour iOS ou dans Office pour Mac prend en charge les mêmes API qu’un complément qui s’exécute dans Office pour Windows, le code de votre complément doit s’exécuter de la même façon sur ces deux plateformes.
    
-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**
    
    Lorsque vous spécifiez des conditions requises d’API dans le manifeste de votre complément, Office détermine si l’application hôte prend en charge ces membres de l’API. Si les membres de l’API sont disponibles dans l’hôte, votre complément sera alors disponible dans cette application hôte. Par ailleurs, vous pouvez effectuer une vérification à l’exécution pour déterminer si une méthode est disponible dans l’hôte avant de l’utiliser dans votre complément. Les vérifications à l’exécution garantissent que votre complément est toujours disponible dans l’hôte et qu’il fournit des fonctionnalités supplémentaires si les méthodes sont disponibles. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](specify-office-hosts-and-api-requirements.md).
    
Pour plus d’informations sur des pratiques plus générales en matière de développement de compléments, consultez la rubrique [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md).


## <a name="see-also"></a>Voir aussi

- [Charger une version test d’un complément Office sur iPad ou Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Débogage des compléments Office sur iPad et Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
