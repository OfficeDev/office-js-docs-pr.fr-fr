---
title: Volets des tâches dans les compléments Office
description: Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: d911101a7df1f1ad8aa01b8e0006bd93d994a193
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092916"
---
# <a name="task-panes-in-office-add-ins"></a>Volets des tâches dans les compléments Office

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Figure 1. Mise en page type du volet Office*

![Illustration affichant une disposition typique du volet Office avec des onglets de section en haut, le logo de l’entreprise et le nom de l’entreprise en bas à gauche, et une icône de paramètres en bas à droite.](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Bonnes pratiques

|À faire|À ne pas faire|
|:-----|:--------|
|Inclure le nom de votre complément dans le titre.|Ne pas ajouter le nom de votre société au titre.|
|Utiliser des noms descriptifs courts dans le titre.|N’ajoutez pas de chaînes telles que « complément », « pour Word » ou « pour Office » au titre de votre complément.|
|Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.|*Aucun.*|
|Inclure un élément de la marque tel que le BrandBar en bas de votre complément, sauf si votre complément doit être utilisé dans Outlook.|*Aucun.*|

## <a name="variants"></a>Variantes

Les images suivantes montrent les différentes tailles du volet Office avec le ruban de l’application Office à une résolution 1366 x 768. Pour Excel, l’espace vertical supplémentaire est requis pour s’adapter à la barre de formule.  

*Figure 2. Tailles de volet des tâches du bureau Office 2016*

![Diagramme affichant les tailles du volet Office de bureau avec une résolution de 1366 x 768.](../images/office-2016-taskpane-sizes.png)

- Excel - 320 x 455 pixels
- PowerPoint - 320 x 531 pixels
- Word - 320 x 531 pixels
- Outlook - 348 x 535 pixels

<br/>

*Figure 3. Tailles du volet Office*

![Diagramme affichant les tailles du volet Office avec une résolution de 1366 x 768.](../images/office-365-taskpane-sizes.png)

- Excel - 350 x 378 pixels
- PowerPoint - 348 x 391 pixels
- Word - 329 x 445 pixels
- Outlook (sur le web) - 320 x 570 pixels

## <a name="personality-menu"></a>Menu Caractéristique

Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac. (Le menu personnalité n’est pas pris en charge dans Outlook.)

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 4. Menu Caractéristique sur Windows*

![Diagramme montrant le menu personnalité sur le bureau Windows.](../images/personality-menu-win.png)

Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace à 34 x 32 pixels, comme illustré.

*figure 5. Menu Caractéristique sur Mac*

![Diagramme montrant le menu personnalité sur le bureau Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un volet des tâches, reportez-vous à [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) sur GitHub.

## <a name="see-also"></a>Voir aussi

- [Cœur de fabric dans les modules](fabric-core.md)
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/ux-design-pattern-templates.md)
