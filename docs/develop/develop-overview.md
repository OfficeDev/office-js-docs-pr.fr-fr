---
title: Développement de compléments Office
description: Présentation du développement de compléments Office.
ms.date: 07/08/2021
localization_priority: Priority
ms.openlocfilehash: 4677f50d718234cb0751b192547fe99ec720d680725aeeed2be9caea904001be
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080828"
---
# <a name="develop-office-add-ins"></a>Développement de compléments Office

> [!TIP]
> Avant de lire cet article, nous vous invitons à consulter [Vue d’ensemble de la plateforme de compléments pour Office](../overview/office-add-ins.md).

Tous les compléments Office sont basés sur la plateforme de compléments Office. Pour les compléments que vous créez, vous devrez comprendre les concepts importants tels que la disponibilité de l’application et de la plateforme, les modèles de programmation de l’API JavaScript Office, la spécification des paramètres et fonctionnalités d’un complément dans le fichier manifeste, la conception de l’interface utilisateur et de l’expérience utilisateur et bien plus encore. Les concepts principaux de développement tels que ceux-ci sont abordés dans la section **Cycle de vie de développement** > **Développer** de la documentation. Consultez les informations ci-dessous avant d’explorer la documentation propre à l’application qui correspond au complément que vous créez (par exemple, [Excel](../excel/index.yml)).

## <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer un complément Office à l’aide du générateur Yeoman pour les compléments Office ou de Visual Studio.

### <a name="yeoman-generator-for-office-add-ins"></a>Générateur Yeoman pour compléments Office

Le[Générateur Yeoman pour les compléments Office](https://github.com/officedev/generator-office) peut être utilisé pour créer un projet de complément Office Node.js qui peut être géré à l’aide de Visual Studio Code ou de tout autre éditeur. Le générateur peut créer des compléments Office pour les éléments suivants :

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Fonctions personnalisées dans Excel

Vous pouvez choisir de créer le projet à l’aide de HTML, CSS et JavaScript, ou d’utiliser Angular ou React. Pour l’infrastructure de votre choix, vous pouvez également choisir entre JavaScript et Typescript . Pour en savoir plus sur la création de compléments avec le générateur Yeoman, consultez [Développez des compléments Office avec Visual Studio Code](../develop/develop-add-ins-vscode.md).

### <a name="visual-studio"></a>Visual Studio

Visual Studio peut être utilisé pour créer des compléments Office pour Excel, Outlook, Word, et PowerPoint. Un projet de complément Office est créé dans le cadre d’une solution Visual Studio et utilise HTML, CSS et JavaScript. Pour en savoir plus sur la création de compléments avec Visual Studio, consultez [Développez des compléments Office avec Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understand-the-two-parts-of-an-office-add-in"></a>Comprendre les deux parties d’un complément Office

Un complément Office se compose de deux parties.

- Le manifeste de complément est un fichier XML qui définit les paramètres et les fonctionnalités du complément.

- L'application web qui définit l'interface utilisateur et les fonctionnalités des composants additionnels tels que les volets Office, les compléments de contenu et les boîtes de dialogue.

L’application web utilise l’API JavaScript Office pour interagir avec le contenu du document Office dans lequel le complément est en cours d’exécution. Votre complément peut également effectuer d’autres opérations que les applications web effectuent généralement, comme appeler des services web externes, faciliter l’authentification des utilisateurs, etc.

### <a name="define-an-add-ins-settings-and-capabilities"></a>Définir les paramètres et les fonctionnalités d’un complément

Un manifeste de complément Office (fichier XML) définit les paramètres et les fonctionnalités du complément. Vous allez configurer le manifeste pour spécifier des éléments tels que :

- Métadonnées décrivant le complément (par exemple, ID, version, description, nom complet, paramètres régionaux par défaut).
- Les applications Office dans lesquelles le complément s’exécute.
- Autorisations nécessaires au complément.
- La manière dont le complément est intégré à Office, y compris toute interface utilisateur personnalisée créée par le complément (par exemple, onglets personnalisés, boutons du ruban).
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

Script Lab est un complément qui vous permet d’explorer l’API JavaScript Office et d’exécuter des extraits de code lorsque vous travaillez dans un programme Office tel qu’Excel ou Word. Il est disponible gratuitement via [AppSource](https://appsource.microsoft.com/product/office/WA104380862), il s’agit d’un outil utile pour inclure votre kit de ressources de développement pendant que vous projetez et vérifiez les fonctionnalités de votre complément. Dans Script Lab, vous pouvez accéder à une bibliothèque d'exemples intégrés pour essayer rapidement des API ou même utiliser un exemple comme point de départ pour votre propre code.

La vidéo d’une minute suivante illustre Script Lab en action.

[![Vidéo d’aperçu montrant l’exécution de Script Lab dans Excel, Word et PowerPoint.](../images/screenshot-wide-youtube.png 'Vidéo de la version préliminaire de Script Lab.')](https://aka.ms/scriptlabvideo)

Si vous souhaitez en savoir plus sur Script Lab, veuillez consulter[Axplorer les API Office JavaScript à l’aide d’un Script Lab](../overview/explore-with-script-lab.md).

## <a name="extend-the-office-ui"></a>Étendre l’interface utilisateur d’Office

Un complément Office peut étendre l'interface utilisateur d'Office à l’aide de commandes de complément et de conteneurs HTML tels que les volets de tâches, les compléments de contenu ou les boîtes de dialogue.

- [Commandes de complément](../design/add-in-commands.md) peuvent être utilisé pour ajouter des onglets, boutons et menus personnalisés au ruban par défaut dans Office, ou développer le menu contextuel par défaut qui apparaît lorsque les utilisateurs cliquent avec le bouton droit sur du texte dans un document Office ou un objet dans Excel. Lorsque les utilisateurs sélectionnent une commande de complément, ils lancent la tâche spécifiée par la commande de complément, par exemple, l’exécution d’un code JavaScript, l’ouverture d’un volet Office ou le lancement d’une boîte de dialogue.

- Les conteneurs HTML tels que [volets Office](../design/task-pane-add-ins.md), [compléments de contenu](../design/content-add-ins.md)et [boîtes de dialogue](../design/dialog-boxes.md) peuvent être utilisés pour afficher une interface utilisateur personnalisée et exposer des fonctionnalités supplémentaires dans une application Office. Le contenu et les fonctionnalités de chaque volet Office, complément de contenu ou boîte de dialogue dérivent d’une page web que vous spécifiez. Ces pages web peuvent utiliser l’API JavaScript Office pour interagir avec le contenu du document Office dans lequel le complément est exécuté, et peuvent également effectuer d’autres actions, telles que appeler des services web externes, faciliter l’authentification des utilisateurs, et bien plus encore.

L’image suivante illustre la commande d’un complément dans le ruban, un volet Office à droite du document et une boîte de dialogue ou un complément de contenu sur le document.

![Diagramme qui illustre les commandes de complément sur le ruban, un volet Office et un complément boîte de dialogue/contenu dans un document Office.](../images/add-in-ui-elements.png)

Pour plus d’informations sur l’extension de l’interface utilisateur d’Office et la conception de l’expérience utilisateur du complément, consultez [Éléments d’interface utilisateur Office pour les compléments Office](../design/interface-elements.md).

## <a name="next-steps"></a>Étapes suivantes

Cet article a décrit les différentes façons de créer des compléments Office, a présenté les façons dont un complément peut étendre l’interface utilisateur d’Office, décrit les ensembles d’API et a introduit Script Lab comme outil précieux pour explorer les API JavaScript Office et le prototypage des fonctionnalités de complément. Maintenant que vous avez exploré ces informations d’introduction, envisagez de poursuivre le parcours de vos compléments Office en suivant les chemins d’accès suivants.

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
