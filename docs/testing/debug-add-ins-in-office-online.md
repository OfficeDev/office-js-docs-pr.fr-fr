---
title: Débogage de compléments dans Office Online
description: Découvrez comment utiliser Office Online pour tester et déboguer vos compléments.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ff77f3d8b3e332288d4ccb3e2d2305d1b1c4a825
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871149"
---
# <a name="debug-add-ins-in-office-online"></a>Débogage de compléments dans Office Online


Vous pouvez créer et déboguer des compléments sur un ordinateur n’exécutant pas Windows, ou le client de bureau Office 2013 ou Office 2016 (par exemple, si vous développez sur un Mac). Cet article décrit la procédure d’utilisation d’Office Online dans le but de tester et de déboguer vos compléments. Cet article décrit comment utiliser Office Online pour tester et déboguer vos compléments. 

## <a name="prerequisites"></a>Conditions requises

Mise en route :

- Si vous n’en avez pas encore, créez un compte de développeur Office 365, ou accédez à un site SharePoint.
    
  > [!NOTE]
  > Pour vous inscrire et obtenir gratuitement un abonnement Office 365 Développeur, participez à notre [programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program). Consultez la [documentation relative au programme pour les développeurs Office 365](/office/developer-program/office-365-developer-program) pour obtenir des instructions détaillées sur la manière de rejoindre le programme, de vous inscrire et de configurer votre abonnement.
     
- Configurez un catalogue de compléments sur Office 365 (SharePoint Online). Un catalogue de compléments est une collection de sites dédiée dans SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office. Si vous disposez de votre propre site SharePoint, vous pouvez configurer une bibliothèque de document de catalogue de compléments. Pour plus d’informations, voir [Publier des compléments de contenu et du volet Office dans un catalogue de compléments sur SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Débogage de compléments à partir d’Excel Online ou de Word Online

Pour déboguer votre complément à l’aide d’Office Online, procédez comme suit :

1. Déployez votre complément vers un serveur prenant en charge le protocole SSL.
    
    > [!NOTE]
    > Nous vous recommandons d’utiliser le [générateur Yeoman](https://github.com/OfficeDev/generator-office) pour créer et héberger votre complément.
     
2. Dans le [fichier manifeste de votre complément](../develop/add-in-manifests.md), mettez à jour la valeur de l’élément **SourceLocation** afin d’inclure un URI absolu, plutôt que relatif. Par exemple :
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. Téléchargez le manifeste dans la bibliothèque de compléments Office du catalogue de compléments sur SharePoint.
    
4. Lancez Excel Online ou Word Online à partir du lanceur d’applications dans Office 365, puis ouvrez un nouveau document.
    
5. Sur l’onglet Insérer, sélectionnez  **Mes compléments** ou **Compléments Office** pour insérer votre complément et le tester dans l’application.
    
6. Utilisez l’outil de débogage de votre navigateur préféré pour déboguer votre complément.

## <a name="potential-issues"></a>Problèmes potentiels    

Voici certains problèmes que vous pouvez rencontrer lorsque vous effectuez des opérations de débogage :
    
- Certaines erreurs JavaScript peuvent provenir d’Office Online.
      
- Le navigateur peut afficher une erreur liée à un certificat non valide que vous devrez contourner.
      
- Si vous définissez des points d’arrêt dans votre code, Office Online peut générer une erreur indiquant qu’il ne peut pas effectuer d’enregistrement.

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
- [Stratégies de validation AppSource](/office/dev/store/validation-policies)  
- [Création d’applications et de compléments AppSource efficaces](/office/dev/store/create-effective-office-store-listings)  
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
    
