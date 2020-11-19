---
title: Volets des tâches dans les compléments Office
description: Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: ed3f3b8fdf7cf62b6016fe8b03393de0d56dfb33
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132017"
---
# <a name="task-panes-in-office-add-ins"></a>Volets des tâches dans les compléments Office

Les volets des tâches sont des surfaces d’interface qui s’affichent généralement sur le côté droit de la fenêtre dans Word, PowerPoint, Excel et Outlook. Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données. Utilisez les volets des tâches lorsque vous n’avez pas besoin d’incorporer des fonctionnalités directement dans le document.

*Figure 1. Mise en page type du volet Office*

![Illustration d’une disposition de volet des tâches par défaut avec des onglets de section en haut, le logo de société et le nom de la société en bas à gauche, et une icône de paramètres dans le coin inférieur droit](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:-----|:--------|
|<ul><li>Inclure le nom de votre complément dans le titre.</li></ul>|<ul><li>Ne pas ajouter le nom de votre société au titre.</li></ul>|
|<ul><li>Utiliser des noms descriptifs courts dans le titre.</li></ul>|<ul><li>N’ajoutez pas de chaînes telles que « complément », « pour Word » ou « pour Office » au titre de votre complément.</li></ul>|
|<ul><li>Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.</li></ul>||
|<ul><li>Inclure un élément de la marque tel que le BrandBar en bas de votre complément, sauf si votre complément doit être utilisé dans Outlook.</li></ul>||

## <a name="variants"></a>Variantes

Les images suivantes présentent les différentes tailles de volet de tâches avec le ruban d’application Office à une résolution 1366x768. Pour Excel, un espace vertical supplémentaire est nécessaire pour accueillir la barre de formule.  

*Figure 2. Tailles de volet des tâches du bureau Office 2016*

![Diagramme affichant les tailles de volet des tâches du Bureau à la résolution 1366x768](../images/office-2016-taskpane-sizes.png)

- Excel-320x455 pixels
- PowerPoint-320 x 531 pixels
- Word-320 x 531 pixels
- Outlook-348x535 pixels

<br/>

*Figure 3. Tailles des volets Office*

![Diagramme affichant la taille des volets des tâches à la résolution 1366x768](../images/office-365-taskpane-sizes.png)

- Excel-350x378 pixels
- PowerPoint-348x391 pixels
- Word-329x445 pixels
- Outlook (sur le Web)-320x570 pixels

## <a name="personality-menu"></a>Menu Caractéristique

Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac.

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 4. Menu Caractéristique sur Windows*

![Diagramme illustrant le menu de personnalité sur le bureau Windows](../images/personality-menu-win.png)

Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace à 34 x 32 pixels, comme illustré.

*figure 5. Menu Caractéristique sur Mac*

![Diagramme illustrant le menu de personnalité sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un volet des tâches, reportez-vous à [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) sur GitHub.

## <a name="see-also"></a>Voir aussi

- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/ux-design-pattern-templates.md)
