---
title: Volets des tâches dans les compléments Office
description: Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cd8d9386fa9f154d611926add12e21f545e36351
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076303"
---
# <a name="task-panes-in-office-add-ins"></a>Volets des tâches dans les compléments Office

Les volets des tâches sont des surfaces d’interface qui s’affichent généralement sur le côté droit de la fenêtre dans Word, PowerPoint, Excel et Outlook. Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données. Utilisez les volets des tâches lorsque vous n’avez pas besoin d’incorporer des fonctionnalités directement dans le document.

*Figure 1. Mise en page type du volet Office*

![Illustration affichant une disposition classique du volet Des tâches avec des onglets de section en haut, le logo de la société et le nom de la société en bas à gauche, et une icône de paramètres en bas à droite.](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:-----|:--------|
|<ul><li>Inclure le nom de votre complément dans le titre.</li></ul>|<ul><li>Ne pas ajouter le nom de votre société au titre.</li></ul>|
|<ul><li>Utiliser des noms descriptifs courts dans le titre.</li></ul>|<ul><li>N’ajoutez pas de chaînes telles que « add-in », « for Word » ou « for Office » au titre de votre application.</li></ul>|
|<ul><li>Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.</li></ul>||
|<ul><li>Inclure un élément de la marque tel que le BrandBar en bas de votre complément, sauf si votre complément doit être utilisé dans Outlook.</li></ul>||

## <a name="variants"></a>Variantes

Les images suivantes montrent les différentes tailles de volet de tâches avec application Office ruban avec une résolution de 1 366 x 768. Pour Excel, l’espace vertical supplémentaire est requis pour s’adapter à la barre de formule.  

*Figure 2. Tailles de volet des tâches du bureau Office 2016*

![Diagramme affichant les tailles du volet Des tâches du bureau à une résolution de 1 366 x 768.](../images/office-2016-taskpane-sizes.png)

- Excel - 320 x 455 pixels
- PowerPoint - 320 x 531 pixels
- Word - 320 x 531 pixels
- Outlook - 348 x 535 pixels

<br/>

*Figure 3. Office tailles de volet de tâches*

![Diagramme affichant les tailles du volet Des tâches à une résolution de 1 366 x 768.](../images/office-365-taskpane-sizes.png)

- Excel - 350 x 378 pixels
- PowerPoint - 348 x 391 pixels
- Word - 329 x 445 pixels
- Outlook (sur le web) - 320 x 570 pixels

## <a name="personality-menu"></a>Menu Caractéristique

Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac.

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 4. Menu Caractéristique sur Windows*

![Diagramme montrant le menu De la personnalité sur Windows bureau.](../images/personality-menu-win.png)

Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace à 34 x 32 pixels, comme illustré.

*figure 5. Menu Caractéristique sur Mac*

![Diagramme montrant le menu Personnalité sur le bureau Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un volet des tâches, reportez-vous à [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) sur GitHub.

## <a name="see-also"></a>Voir aussi

- [Cœur de fabric dans les modules](fabric-core.md)
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/ux-design-pattern-templates.md)
