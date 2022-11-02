---
title: Compléments Office de contenu
description: Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Excel ou PowerPoint. Ils permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c10893d60f64d875d92aec979a5700630b2cf96c
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810245"
---
# <a name="content-office-add-ins"></a>Compléments Office de contenu

Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Excel ou PowerPoint. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document.  

*Figure 1. Mise en page type pour les compléments de contenu*

![Disposition classique des compléments de contenu dans une application Office.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Meilleures pratiques

- Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.
- Inclure un élément de la marque tel que BrandBar en bas de votre complément (s’applique aux compléments Excel et PowerPoint uniquement).

## <a name="variants"></a>Variantes

Les tailles de complément de contenu pour Excel et PowerPoint dans le bureau Office et dans un navigateur web sont spécifiées par l’utilisateur.

## <a name="personality-menu"></a>Menu Caractéristique

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 2. Menu Caractéristique sur Windows*

![Menu de personnalité de 12 x 32 pixels sur le bureau Windows.](../images/personality-menu-win.png)

Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace occupé à 34 x 32 pixels, comme illustré.

*Figure 3. Menu Caractéristique sur Mac*

![Menu de personnalité de 34 x 32 pixels sur le bureau Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un complément de contenu, reportez-vous à [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.

## <a name="support-considerations"></a>Considérations relatives à la prise en charge

- Vérifiez si votre complément Office fonctionnera sur une [application ou une plateforme Office spécifique](/javascript/api/requirement-sets).
- Certains compléments de contenu peuvent exiger que l’utilisateur accepte que le complément lise et écrive dans Excel ou PowerPoint. Vous pouvez déclarer le [niveau des autorisations](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) que vous souhaitez attribuer à votre utilisateur dans le manifeste du complément.  
- Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.

## <a name="see-also"></a>Voir aussi

- [Application cliente Office et disponibilité de la plateforme pour les compléments Office](/javascript/api/requirement-sets)
- [Cœur de fabric dans les modules](fabric-core.md)
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/ux-design-pattern-templates.md)
- [Demande d’autorisations d’utilisation de l’API dans des compléments](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
