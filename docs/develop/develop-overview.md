---
title: Développement de compléments Office
description: Présentation du développement de compléments Office.
ms.date: 05/25/2022
ms.localizationpriority: high
ms.openlocfilehash: 82573d90f9fa22cb524da01226995e861c258b81
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810021"
---
# <a name="develop-office-add-ins"></a>Développement de compléments Office

> [!TIP]
> Avant de lire cet article, nous vous invitons à consulter [Vue d’ensemble de la plateforme de compléments pour Office](../overview/office-add-ins.md).

Tous les compléments Office sont basés sur la plateforme de compléments Office. Pour les compléments que vous créez, vous devrez comprendre les concepts importants tels que la disponibilité de l’application et de la plateforme, les modèles de programmation de l’API JavaScript Office, la spécification des paramètres et fonctionnalités d’un complément dans le fichier manifeste, la conception de l’interface utilisateur et de l’expérience utilisateur et bien plus encore. Les concepts principaux de développement tels que ceux-ci sont abordés dans la section **Cycle de vie de développement** > **Développer** de la documentation. Consultez les informations ci-dessous avant d’explorer la documentation propre à l’application qui correspond au complément que vous créez (par exemple, [Excel](../excel/index.yml)).

## <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer un complément Office à l’aide du [générateur Yeoman pour les compléments Office](yeoman-generator-overview.md) ou de Visual Studio.

### <a name="yeoman-generator"></a>Générateur Yeoman

The Yeoman generator for Office Add-ins can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor. The generator can create Office Add-ins for any of the following:

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Fonctions personnalisées dans Excel

Créez votre projet en utilisant HTML, CSS et JavaScript (ou TypeScript), ou en utilisant Angular ou React. Pour l’infrastructure de votre choix, vous pouvez également choisir entre JavaScript et Typescript . Pour plus d’informations sur la création de compléments avec le générateur, voir [Générateur Yeoman pour compléments Office](yeoman-generator-overview.md).

### <a name="visual-studio"></a>Visual Studio

Visual Studio peut être utilisé pour créer des compléments Office pour Excel, Outlook, Word, et PowerPoint. Un projet de complément Office est créé dans le cadre d’une solution Visual Studio et utilise HTML, CSS et JavaScript. Pour en savoir plus sur la création de compléments avec Visual Studio, consultez [Développez des compléments Office avec Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understand-the-two-parts-of-an-office-add-in"></a>Comprendre les deux parties d’un complément Office

Un complément Office se compose de deux parties.

- Le manifeste de complément est un fichier XML qui définit les paramètres et les fonctionnalités du complément.

- L'application web qui définit l'interface utilisateur et les fonctionnalités des composants additionnels tels que les volets Office, les compléments de contenu et les boîtes de dialogue.

The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running. Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.

### <a name="define-an-add-ins-settings-and-capabilities"></a>Définir les paramètres et les fonctionnalités d’un complément

Un manifeste de complément Office (fichier XML) définit les paramètres et les fonctionnalités du complément. Vous allez configurer le manifeste pour spécifier des éléments tels que :

- Métadonnées décrivant le complément (par exemple, ID, version, description, nom complet, paramètres régionaux par défaut).
- Les applications Office dans lesquelles le complément s’exécute.
- Autorisations nécessaires au complément.
- Comment le complément s’intègre à Office, y compris toute interface utilisateur personnalisée créée par le complément (par exemple, des onglets personnalisés ou des boutons du ruban personnalisés).
- L’emplacement des images que le complément utilise pour la personnalisation et l'iconographie des commandes.
- Dimensions du complément (par exemple, dimensions pour les compléments de contenu, la hauteur demandée pour des compléments Outlook).
- Règles qui spécifient le moment où le complément est activé dans le contexte d’un message ou d’un rendez-vous (pour les compléments Outlook uniquement).

Si vous souhaitez en savoir plus sur le manifeste, veuillez consulter l’article sur le [manifeste XML de compléments Office](add-in-manifests.md).

### <a name="interact-with-content-in-an-office-document"></a>Interagir avec du contenu dans un document Office

Un complément Office peut utiliser l’API JavaScript Office pour interagir avec le contenu du document Office dans lequel le complément est exécuté.

#### <a name="access-the-office-javascript-api-library"></a>Accéder à la bibliothèque d’API JavaScript Office

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a>Modèles API

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a>Ensembles de conditions requises de l’API

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="explore-apis-with-script-lab"></a>Explorer les API avec Script Lab

Script Lab est un complément qui vous permet d’explorer l’API JavaScript Office et d’exécuter des extraits de code lorsque vous travaillez dans un programme Office tel qu’Excel ou Word. Il est disponible gratuitement via AppSource, il s’agit d’un outil utile pour inclure votre kit de ressources de développement pendant que vous projetez et vérifiez les fonctionnalités de votre complément. Dans Script Lab, vous pouvez accéder à une bibliothèque d'exemples intégrés pour essayer rapidement des API ou même utiliser un exemple comme point de départ pour votre propre code.

La vidéo d’une minute suivante illustre Script Lab en action.

[![Vidéo d’aperçu montrant l’exécution de Script Lab dans Excel, Word et PowerPoint.](../images/screenshot-wide-youtube.png 'Vidéo de la version préliminaire de Script Lab')](https://aka.ms/scriptlabvideo)

Si vous souhaitez en savoir plus sur Script Lab, veuillez consulter[Axplorer les API Office JavaScript à l’aide d’un Script Lab](../overview/explore-with-script-lab.md).

## <a name="extend-the-office-ui"></a>Étendre l’interface utilisateur d’Office

Un complément Office peut étendre l'interface utilisateur d'Office à l’aide de commandes de complément et de conteneurs HTML tels que les volets de tâches, les compléments de contenu ou les boîtes de dialogue.

- [Commandes de complément](../design/add-in-commands.md) peuvent être utilisé pour ajouter des onglets, boutons et menus personnalisés au ruban par défaut dans Office, ou développer le menu contextuel par défaut qui apparaît lorsque les utilisateurs cliquent avec le bouton droit sur du texte dans un document Office ou un objet dans Excel. Lorsque les utilisateurs sélectionnent une commande de complément, ils lancent la tâche spécifiée par la commande de complément, par exemple, l’exécution d’un code JavaScript, l’ouverture d’un volet Office ou le lancement d’une boîte de dialogue.

- Les conteneurs HTML tels que [volets Office](../design/task-pane-add-ins.md), [compléments de contenu](../design/content-add-ins.md)et [boîtes de dialogue](../develop/dialog-api-in-office-add-ins.md) peuvent être utilisés pour afficher une interface utilisateur personnalisée et exposer des fonctionnalités supplémentaires dans une application Office. Le contenu et les fonctionnalités de chaque volet Office, complément de contenu ou boîte de dialogue dérivent d’une page web que vous spécifiez. Ces pages web peuvent utiliser l’API JavaScript Office pour interagir avec le contenu du document Office dans lequel le complément est exécuté, et peuvent également effectuer d’autres actions, telles que appeler des services web externes, faciliter l’authentification des utilisateurs, et bien plus encore.

L’image suivante illustre la commande d’un complément dans le ruban, un volet Office à droite du document et une boîte de dialogue ou un complément de contenu sur le document.

![Diagramme qui illustre les commandes de complément sur le ruban, un volet Office et un complément boîte de dialogue/contenu dans un document Office.](../images/add-in-ui-elements.png)

Pour plus d’informations sur l’extension de l’interface utilisateur d’Office et la conception de l’expérience utilisateur du complément, consultez [Éléments d’interface utilisateur Office pour les compléments Office](../design/interface-elements.md).

## <a name="next-steps"></a>Étapes suivantes

This article has outlined the different ways to create Office Add-ins, introduced the ways that an add-in can extend the Office UI, described the API sets, and introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality. Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.

### <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](../index.yml). Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](../index.yml).

### <a name="learn-more"></a>En savoir plus

Pour en savoir plus sur le développement, le test et la publication de compléments Office, explorez cette documentation.

> [!TIP]
> Pour les compléments que vous créez, vous utiliserez les informations de la section [Cycle de vie de développement](../overview/core-concepts-office-add-ins.md) de cette documentation, ainsi que les informations de la section propre à l’application qui correspond au type de complément que vous créez (par exemple, [Excel](../excel/index.yml)).

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publier des compléments Office](../publish/publish.md)
