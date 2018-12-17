---
title: Compléments Office de contenu
description: Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Excel ou PowerPoint. Ils permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données.
ms.date: 12/13/2018
ms.openlocfilehash: efeef65381acb62f877975652d90d962a86a6b0a
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270648"
---
# <a name="content-office-add-ins"></a>Compléments Office de contenu

Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Excel ou PowerPoint. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document.  

*Figure 1. Mise en page type pour les compléments de contenu*

![Exemple d’image affichant une mise en page typique pour des compléments de contenu.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Meilleures pratiques

- Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.
- Inclure un élément de la marque tel que BrandBar en bas de votre complément (s’applique aux compléments Excel et PowerPoint uniquement).

## <a name="variants"></a>Variantes

Les tailles des compléments de contenu pour Excel et PowerPoint dans le bureau Office et Office 365 sont spécifiées par l’utilisateur.

## <a name="personality-menu"></a>Menu Caractéristique

Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac.

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 2. Menu Caractéristique sur Windows* 

![Image illustrant le menu Caractéristique sur le bureau Windows](../images/personality-menu-win.png)


Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace occupé à 34 x 32 pixels, comme illustré.

*Figure 3. Menu Caractéristique sur Mac*

![Image illustrant le menu Caractéristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un complément de contenu, reportez-vous à [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.

## <a name="support-considerations"></a>Considérations relatives à la prise en charge
- Vérifiez si votre complément Office fonctionne sur une [plateforme hôte Office spécifique](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability). 
- Certains compléments de contenu peuvent exiger que l’utilisateur accepte que le complément lise et écrive dans Excel ou PowerPoint. Vous pouvez déclarer le [niveau des autorisations](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) que vous souhaitez attribuer à votre utilisateur dans le manifeste du complément.  
- Les compléments de contenu sont pris en charge dans Excel et PowerPoint dans Office 2013 et versions ultérieures. Si vous ouvrez un complément dans une version d’Office qui ne prend pas en charge les compléments web Office, le complément s’affichera comme une image.

## <a name="see-also"></a>Voir aussi
- [Disponibilité des compléments Office sur les plateformes et les hôtes](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [Office UI Fabric dans des compléments Office](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [Demande d’autorisations d’utilisation de l’API dans des compléments de contenu et de volet des tâches](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
