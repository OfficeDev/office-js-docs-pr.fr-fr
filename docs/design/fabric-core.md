---
title: Cœur de fabric dans les modules
description: Obtenez une vue d’ensemble de l’utilisation de Fabric Core et des composants de l’interface utilisateur fabric dans Office des composants.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: cd534809bb443134e2df06de478e8283a3452aac
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150119"
---
# <a name="fabric-core-in-office-add-ins"></a>Cœur de fabric dans les modules

Fabric Core est une collection open source de classes CSS et de mixins SASS conçus pour être utilisés dans des React *Office* non utilisés. Fabric Core contient des éléments de base du Fluent de conception de l’interface utilisateur, tels que les icônes, les couleurs, les polices et les grilles. Fabric Core est indépendant de l’infrastructure, il peut donc être utilisé avec n’importe quelle application à page unique ou n’importe quelle infrastructure d’interface utilisateur web côté serveur. (Il est appelé « Fabric Core » au lieu de « Fluent Core » pour des raisons historiques.)

Si l’interface utilisateur de votre React n’est pas basée sur React, vous pouvez également utiliser un ensemble de composants React non utilisés. Voir [Utiliser Office composants JS UI Fabric.](#use-office-ui-fabric-js-components)

> [!NOTE]
> Cet article décrit l’utilisation de Fabric Core dans le contexte de Office des modules. Mais il est également utilisé dans un large éventail d’applications Microsoft 365 et d’extensions. Pour plus d’informations, [voir Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) et le repo open source Office [UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).

## <a name="use-fabric-core-icons-fonts-colors"></a>Utiliser Fabric Core : icônes, polices, couleurs

1. Ajoutez la référence CDN au code HTML sur votre page.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Utilisez les polices et les icônes Fabric Core.

    Pour utiliser une icône Fabric Core, incluez l’élément « i » sur votre page, puis référencez les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police. Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078d7).

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Pour obtenir des instructions plus détaillées, [voir Fluent’interface utilisateur.](https://developer.microsoft.com/fluentui#/styles/web/icons) Pour trouver d’autres icônes disponibles dans Fabric Core, utilisez la fonctionnalité de recherche sur cette page. Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`.

    Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Fabric Core, voir [Typographie](https://developer.microsoft.com/fluentui#/styles/web/typography) et la table des matières **Couleurs** dans [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).

Des exemples sont inclus dans [les exemples plus](#samples) loin dans cet article.

## <a name="use-office-ui-fabric-js-components"></a>Utiliser Office composants JS UI Fabric

Les applications avec des interfaces utilisateur non React peuvent également utiliser l’un des nombreux composants de [Office UI Fabric JS,](https://github.com/OfficeDev/office-ui-fabric-js)y compris les boutons, les boîtes de dialogue, les suceurs et bien plus encore. Consultez le lisez-moi du repo pour obtenir des instructions.

Des exemples sont inclus dans [les exemples plus](#samples) loin dans cet article.

## <a name="samples"></a>Exemples

Les exemples de composants suivants utilisent Fabric Core et/ou Office composants JS UI Fabric. Certains de ces dépôts sont archivés, ce qui signifie qu’ils ne sont plus mis à jour avec des correctifs de bogue ou de sécurité, mais vous pouvez toujours les utiliser pour apprendre à utiliser les composants d’interface utilisateur Fabric Core et Fabric.

- [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel Add-in SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel Tendances des dépenses de WoodGrove du add-in](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office Exemple d’interface utilisateur de la structure de la structure de la add-in](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook Add-in GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint Add-in Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
