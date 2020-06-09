---
title: Débogage de compléments dans Office sur le web
description: Découvrez comment utiliser Office sur le web pour tester et déboguer vos compléments.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4f2aa498e5d9f49fcdf306ac2c4c80ea6fbd496c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611238"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Débogage de compléments dans Office sur le web


Vous pouvez créer et déboguer des compléments sur un ordinateur n’exécutant pas Windows, ou le client de bureau Office 2013 ou Office 2016 (par exemple, si vous développez sur un Mac). Cet article décrit la procédure d’utilisation d’Office Online dans le but de tester et de déboguer vos compléments. Cet article décrit comment utiliser Office sur le web pour tester et déboguer vos compléments. 

## <a name="prerequisites"></a>Conditions préalables

Mise en route :

- Si vous n’en avez pas encore, créez un compte de développeur Office 365, ou accédez à un site SharePoint.

  > [!NOTE]
  > Pour obtenir gratuitement un abonnement renouvelable de 90 jours à Office 365 Développeur, participez à notre [programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program). Consultez la [documentation relative au programme pour les développeurs Office 365](/office/developer-program/office-365-developer-program) pour obtenir des instructions détaillées sur la manière de rejoindre le programme et de configurer votre abonnement.

- Configurez un catalogue d’applications sur Office 365 (SharePoint Online). Un catalogue d’applications est une collection de sites dédiée dans SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office. Si vous disposez de votre propre site SharePoint, vous pouvez configurer une bibliothèque de document de catalogue d’applications. Pour plus d’informations, voir [Publier des compléments de contenu et du volet Office dans un catalogue d’applications sur SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a>Débogage de compléments à partir d’Excel ou de Word sur le web

Pour déboguer votre complément à l’aide d’Office sur le web, procédez comme suit :

1. Déployez votre complément vers un serveur prenant en charge le protocole SSL.

    > [!NOTE]
    > Nous vous recommandons d’utiliser le [générateur Yeoman](https://github.com/OfficeDev/generator-office) pour créer et héberger votre complément.

2. Dans le [fichier manifeste de votre complément](../develop/add-in-manifests.md), mettez à jour la valeur de l’élément **SourceLocation** afin d’inclure un URI absolu, plutôt que relatif. Par exemple :

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. Téléchargez le manifeste dans la bibliothèque de compléments Office du catalogue d’applications sur SharePoint.

4. Lancez Excel ou Word sur le web à partir du lanceur d’applications dans Office 365, puis ouvrez un nouveau document.

5. Sur l’onglet Insérer, sélectionnez **Mes compléments** ou **Compléments Office** pour insérer votre complément et le tester dans l’application.

6. Utilisez l’outil de débogage de votre navigateur préféré pour déboguer votre complément.

## <a name="potential-issues"></a>Problèmes potentiels

Voici certains problèmes que vous pouvez rencontrer lorsque vous effectuez des opérations de débogage :

- Certaines erreurs JavaScript peuvent provenir d’Office sur le web.

- Le navigateur peut afficher une erreur relative à un certificat non valide que vous devrez contourner. Le processus d’exécution de cette opération varie en fonction du navigateur et des interfaces utilisateur des différents navigateurs permettant d’effectuer cette modification régulièrement. Vous devez effectuer une recherche dans l’aide du navigateur ou rechercher des instructions en ligne. (Par exemple, recherchez « Avertissement de certificat Microsoft Edge non valide ».) La plupart des navigateurs, sur la page d’avertissement, comportent un lien qui vous permet d’accéder à la page du complément. Par exemple, Microsoft Edge comporte un lien « Accéder à la page web (non recommandé) ». En général, vous devez passer par ce lien chaque fois que le complément est rechargé. Pour un contournement plus long, consultez l’aide comme suggéré.

- Si vous définissez des points d’arrêt dans votre code, Office sur le web peut générer une erreur indiquant qu’il ne peut pas effectuer d’enregistrement.

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
- [Stratégies de validation AppSource](/legal/marketplace/certification-policies)  
- [Création d’applications et de compléments AppSource efficaces](/office/dev/store/create-effective-office-store-listings)  
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
