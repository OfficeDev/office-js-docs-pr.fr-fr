---
title: Cœur de fabric dans les modules
description: Obtenez une vue d’ensemble de l’utilisation de Fabric Core et des composants de l’interface utilisateur fabric dans Office des composants.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3d10cc5d8f33c8dd66f4f988fdd5a082580b1aca
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074202"
---
# <a name="fabric-core-in-office-add-ins"></a>Cœur de fabric dans les modules

Fabric Core est une collection open source de classes CSS et de mixins SASS conçus pour être utilisés dans des React *Office* non utilisés. Fabric Core contient des éléments de base du Fluent de conception de l’interface utilisateur, tels que les icônes, les couleurs, les polices et les grilles. Fabric Core est indépendant de l’infrastructure, il peut donc être utilisé avec n’importe quelle application à page unique ou n’importe quelle infrastructure d’interface utilisateur web côté serveur. (Il est appelé « Fabric Core » au lieu de « Fluent Core » pour des raisons historiques.)

Si l’interface utilisateur de votre React n’est pas basée sur un React, vous pouvez également utiliser un ensemble de composants non React de données. Voir [Utiliser Office composants JS UI Fabric.](#use-office-ui-fabric-js-components)

> [!NOTE]
> Cet article décrit l’utilisation de Fabric Core dans le contexte de Office des modules. Mais il est également utilisé dans un large éventail d’applications Microsoft 365 et d’extensions. Pour plus d’informations, [voir Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) et le repo open source Office [UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).

## <a name="use-fabric-core-icons-fonts-colors"></a>Utiliser Fabric Core : icônes, polices, couleurs

1. Ajoutez la référence de réseau de distribution de contenu (CDN) au code HTML sur votre page.

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

## <a name="samples"></a>Échantillons

Les exemples de modules suivants utilisent Fabric Core et/ou Office composants JS UI Fabric. Certains de ces dépôts sont archivés, ce qui signifie qu’ils ne sont plus mis à jour avec des correctifs de bogue ou de sécurité, mais vous pouvez toujours les utiliser pour apprendre à utiliser les composants d’interface utilisateur Fabric Core et Fabric.

- [Excel javaScript SalesTracker pour le add-in](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel De salesLeads de Excel](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel des dépenses de woodgrove de la boutique de produits](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel de contenu de l’assurance Humongous](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office de l’interface utilisateur de la structure de la structure du Office](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook GifMe de l’ajout](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint de microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
