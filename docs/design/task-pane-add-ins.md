---
title: Volets des tâches dans les compléments Office
description: Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093755"
---
# <a name="task-panes-in-office-add-ins"></a>Volets des tâches dans les compléments Office
 
Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Figure 1. Mise en page type du volet Office*

![Image affichant une disposition du volet des tâches](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:-----|:--------|
|<ul><li>Inclure le nom de votre complément dans le titre.</li></ul>|<ul><li>Ne pas ajouter le nom de votre société au titre.</li></ul>|
|<ul><li>Utiliser des noms descriptifs courts dans le titre.</li></ul>|<ul><li>N’ajoutez pas de chaînes telles que « complément », « pour Word » ou « pour Office » au titre de votre complément.</li></ul>|
|<ul><li>Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.</li></ul>||
|<ul><li>Inclure un élément de la marque tel que le BrandBar en bas de votre complément, sauf si votre complément doit être utilisé dans Outlook.</li></ul>||


## <a name="variants"></a>Variantes

The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*Figure 2. Tailles de volet des tâches du bureau Office 2016*

![Image affichant les tailles de volet des tâches du bureau à une résolution de 1 366 x 768](../images/office-2016-taskpane-sizes.png)

- Excel - 320 x 455
- PowerPoint - 320 x 531
- Word - 320 x 531
- Outlook - 348 x 535

<br/>

*Figure 3. Tailles des volets Office*

![Image affichant les tailles de volet des tâches du bureau à une résolution de 1 366 x 768](../images/office-365-taskpane-sizes.png)

- Excel - 350 x 378
- PowerPoint - 348 x 391
- Word - 329 x 445
- Outlook (sur le web) - 320 x 570

## <a name="personality-menu"></a>Menu Caractéristique

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 4. Menu Caractéristique sur Windows*

![Image illustrant le menu Caractéristique sur le bureau Windows](../images/personality-menu-win.png)

Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace à 34 x 32 pixels, comme illustré.

*figure 5. Menu Caractéristique sur Mac*

![Image illustrant le menu Caractéristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un volet des tâches, reportez-vous à [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) sur GitHub. 


## <a name="see-also"></a>Voir aussi

- [Office UI Fabric dans des compléments Office](office-ui-fabric.md) 
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/ux-design-pattern-templates.md)

