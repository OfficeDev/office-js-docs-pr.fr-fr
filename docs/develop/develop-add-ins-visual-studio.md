---
title: Développement de compléments Office avec Visual Studio
description: Comment développer un complément Office à l’aide de Visual Studio
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 2acafc08dcbb8e1fde439f2be289afff3185272c
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915055"
---
# <a name="develop-office-add-ins-with-visual-studio"></a>Développement de compléments Office avec Visual Studio

Cet article explique comment utiliser Visual Studio pour développer votre complément Office. Si votre complément est déjà créé, vous pouvez passer directement à la section [Développer le complément à l’aide de Visual Studio](#develop-the-add-in-using-visual-studio).

> [!NOTE]
> À la place de Visual Studio, vous pouvez choisir d’utiliser le générateur Yeoman pour compléments Office et le code VS afin de créer un complément. Pour en savoir plus sur cette option, voir [Création d’un complément Office](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).

## <a name="create-the-add-in-project-using-visual-studio"></a>Créer un projet de complément Office à l’aide de Visual Studio

Visual Studio peut être utilisé pour créer des compléments Office pour Excel, Outlook, Word et PowerPoint. Un projet de complément Office est créé dans le cadre d’une solution Visual Studio et utilise HTML, CSS et JavaScript. Pour créer un complément Office avec Visual Studio, suivez les instructions dans le démarrage rapide qui correspond au complément que vous souhaitez créer :

- [Démarrage rapide Excel](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Démarrage rapide Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context&tabs=visualstudio)
- [Démarrage rapide Word](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [Démarrage rapide PowerPoint](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

Visual Studio ne prend pas en charge la création de compléments Office pour OneNote ou Project. Pour créer des compléments Office pour l’un de ces hôtes, vous devez utiliser le générateur Yeoman pour compléments Office, comme décrit dans le [Démarrage rapide OneNote](../quickstarts/onenote-quickstart.md) ou le [Démarrage rapide Project](../quickstarts/project-quickstart.md).

## <a name="develop-the-add-in-using-visual-studio"></a>Développer votre complément à l’aide de Visual Studio

Visual Studio crée un complément de base avec une fonctionnalité limitée. Vous pouvez personnaliser le complément en modifiant le [manifeste](add-in-manifests.md), HTML, JavaScript et des fichiers CSS dans Visual Studio. Pour obtenir une description de haut niveau de la structure de projet et des fichiers dans le projet de complément créé par Visual Studio, consultez les guides Visual Studio dans le guide de démarrage rapide que vous avez achevé pour créer votre complément. 

> [!TIP]
> Un complément Office étant une application web, vous devez maîtriser les compétences de base en matière de développement web pour personnaliser votre complément. Si vous débutez avec JavaScript, nous vous conseillons de consulter le didacticiel [Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

Pour personnaliser votre complément, vous devez comprendre les concepts décrits dans les [Principaux concepts > Développer](develop-overview.md) de la documentation, ainsi que ceux décrits dans la documentation spécifique de l’hôte qui correspond au complément que vous créez (par exemple, [Excel](../excel/index.md)). 

## <a name="test-and-debug-the-add-in"></a>Tester et déboguer le complément

Les méthodes de test, de débogage et de résolution des problèmes liés aux compléments Office varient selon la plateforme. Pour plus d’informations, voir [Déboguer des compléments Office dans Visual Studio](debug-office-add-ins-in-visual-studio.md) et [Tester et déboguer les compléments Office](../testing/test-debug-office-add-ins.md).

## <a name="publish-the-add-in"></a>Publier le complément

Un complément Office comprend une application Web et un fichier manifeste. L’application Web définit l’interface utilisateur et les fonctionnalités du complément, tandis que le manifeste spécifie l’emplacement de l’application Web et définit les paramètres et fonctionnalités du complément.

Lorsque vous développez votre complément dans Visual Studio, celui-ci est exécuté sur votre serveur web local (`localhost`). Lorsque votre complément fonctionne comme vous le souhaitez et que vous êtes prêt à le publier pour permettre à d’autres utilisateurs d’y accéder, vous devez procéder comme suit :

1. Déployer l’application web sur un serveur web ou un service d’hébergement web (par exemple, Microsoft Azure).
2. Mettre à jour le manifeste pour préciser l’URL de l’application déployée. 
3. Choisir la méthode que vous voulez utiliser pour [déployer votre complément Office](../publish/publish.md), puis suivre les instructions pour publier le fichier manifeste.

## <a name="see-also"></a>Voir aussi

- [Création de compléments Office](../overview/office-add-ins-fundamentals.md)
- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)