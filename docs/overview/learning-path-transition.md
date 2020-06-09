---
title: Transition ! Guide pour les créateurs de compléments VSTO qui créent de compléments web Office
description: Chemin d’accès recommandé pour les développeurs de compléments VSTO expérimentés pour la formation de ressources pour les compléments web Office.
ms.date: 05/10/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 499a8fdf12c2f46c5cf5fc5c37f8bb68af540e57
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604572"
---
# <a name="transition-here-a-guide-for-vsto-add-in-creators-making-office-web-add-ins"></a>Transition ! Guide pour les créateurs de compléments VSTO qui créent de compléments web Office

Par conséquent, vous avez créé des compléments VSTO pour les applications Office qui s’exécutent sur Windows et vous explorez maintenant la nouvelle façon d’étendre Office qui s’exécute sur Windows, Mac et la version en ligne de la suite Office : compléments web Office.

Votre compréhension des modèles objets pour Excel, Word et les autres applications Office sera très utile, car les modèles d’objets dans les compléments web Office suivent les mêmes modèles. Il y a cependant quelques défis :

- Vous utiliserez une langue différente (JavaScript ou TypeScript) au lieu de, C# ou Visual Basic .NET. (Il existe également un moyen, décrit ci-dessous, de réutiliser une partie de votre code existant dans un complément web.)
- Les compléments web Office ne sont pas déployés différemment des compléments VSTO.
- Les compléments web Office sont des applications web qui s’exécutent dans une fenêtre de navigateur simplifiée incorporée dans l’application Office, vous devez donc obtenir une compréhension de base des applications web et de leur hébergement sur des serveurs web ou des comptes cloud. 

Pour ces raisons, l’article aborde nos cours d’apprentissage pour les débutants complets pour les extensions Office : [Commencez ici ! Un guide pour les débutants qui créent des compléments Office](learning-path-beginner.md). Nous avons ajouté des ressources d’apprentissage supplémentaires pour aider les développeurs de compléments VSTO à tirer parti de leur expérience et les aider à réutiliser leur code existant.

## <a name="step-0-prerequisites"></a>Étape 0 : Conditions requises

- Les compléments web Office (également appelés compléments Office) sont essentiellement des applications web incorporées dans Office. Vous devez donc d’abord posséder une connaissance de base des applications web et de la façon dont elles sont hébergées sur le web. Il existe une quantité considérable d’informations à ce sujet sur Internet, dans les livres et dans les cours en ligne. Une bonne façon de commencer, si vous ne possédez aucune connaissance préalable des applications web, consiste à rechercher « Qu’est-ce qu’une application web ? » sur Bing.
- Le langage de programmation principal que vous utiliserez pour créer des compléments Office est JavaScript ou TypeScript. Vous pouvez considérer le langage TypeScript comme une version fortement typée de JavaScript. Si vous n’êtes pas familiarisé avec l’un ou l’autre de ces langages, mais que vous avez de l’expérience avec les langages VBA, VB.Net et C#, vous trouverez probablement TypeScript plus facile à apprendre. Là encore, il existe une multitude d’informations relatives à ces langages sur Internet, dans les livres et dans les cours en ligne.

## <a name="step-1-begin-with-fundamentals"></a>Étape 1 : Commencer par les notions de base

Nous savons que vous êtes impatient de commencer à coder, mais il convient de lire certains points concernant les compléments Office avant d’ouvrir votre IDE ou votre éditeur de code.

- [Vue d’ensemble de la plateforme des compléments Office](office-add-ins.md) : découvrez les compléments web Office et leurs différences par rapport aux anciennes méthodes d’extension d’Office, telles que les compléments VSTO.
- [Création de compléments Office](office-add-ins-fundamentals.md) : obtenez une vue d’ensemble du développement et du cycle de vie des compléments Office, y compris les outils, la création d’une interface utilisateur de complément et l’utilisation des API JavaScript pour interagir avec le document Office.

Ces articles comportent un grand nombre de liens. Toutefois, si vous effectuez une transition vers les compléments web Office, nous vous recommandons de revenir ici lorsque vous les aurez lus et de passer à la section suivante.

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>Étape 2 : Installer les outils et créer votre premier complément

Vous avez maintenant une vue d’ensemble, alors lancez-vous avec l’un de nos guides de démarrage rapide. Pour découvrir la plateforme, nous vous recommandons d’utiliser le guide de démarrage rapide d’Excel. Il existe une version basée sur Visual Studio et une autre basée sur Node.js et Visual Studio Code. Si vous effectuez une transition à partir de compléments VSTO, vous trouverez probablement la version de Visual Studio la plus facile à utiliser.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js et Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>Étape 3 : Code

Vous ne pouvez pas apprendre à conduire en lisant le manuel du propriétaire, alors commencez à coder à l’aide de ce [didacticiel Excel](../tutorials/excel-tutorial.md). Vous utiliserez la bibliothèque JavaScript pour Office et du code XML dans le manifeste du complément. Il n’est pas nécessaire de mémoriser quoi que ce soit, car vous obtiendrez plus d’informations sur ces deux éléments plus tard.

## <a name="step-4-understand-the-javascript-library"></a>Étape 4 : Comprendre la bibliothèque JavaScript

Obtenez une vue d’ensemble de la bibliothèque JavaScript pour Office avec ce didacticiel de Microsoft Learn : [Comprendre les API JavaScript pour Office](/learn/modules/intro-office-add-ins/3-apis).

Explorez ensuite les API JavaScript pour Office à l’aide de l’[outil Script Lab](explore-with-script-lab.md), un bac à sable pour l’exécution et l’exploration des API.

### <a name="special-resource-for-vsto-add-in-developers"></a>Ressources spéciales pour les développeurs de compléments VSTO

Il s’agit d’un bon point de départ pour jeter un coup d’œil à l’exemple de complément, [Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker). Il a été créé pour mettre en évidence les similitudes et les différences entre les compléments VSTO et les compléments web Office, et le fichier Lisez-moi de l’exemple indique les points importants de comparaison.

## <a name="step-5-understand-the-manifest"></a>Étape 5 : Comprendre le manifeste

Découvrez les objectifs du manifeste du complément web et consultez une présentation de ses balisages XML sur la page [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a>Étape 6 (pour les développeurs VSTO uniquement) : réutiliser votre code VSTO

Vous pouvez réutiliser une partie de votre code de complément VSTO dans un complément web Office en le déplaçant vers le serveur principal de votre application web sur le serveur et en le rendant disponible pour votre code JavaScript ou votre dactylographié comme API web. Pour obtenir des instructions, consultez [Didacticiel : partage de codes entre un complément VSTO et un complément Office à l’aide d’une bibliothèque de codes partagée](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).

## <a name="next-steps"></a>Étapes suivantes

Félicitations pour avoir terminé le parcours d’apprentissage pour les développeurs des compléments VSTO pour les compléments web Office ! Voici quelques suggestions pour approfondir les informations contenues dans notre documentation :

- Didacticiels ou guides de démarrage rapide pour les autres applications Office :

  - [Guide de démarrage rapide de OneNote](../quickstarts/onenote-quickstart.md)
  - [Didacticiel Outlook](/outlook/add-ins/addin-tutorial)
  - [Didacticiel PowerPoint](../tutorials/powerpoint-tutorial.md)
  - [Guide de démarrage rapide de Project](../quickstarts/project-quickstart.md)
  - [Didacticiel Word](../tutorials/word-tutorial.md)

- Autres sujets importants :

  - [Développement de compléments Office](../develop/develop-overview.md)
  - [Meilleures pratiques en matière de développement de compléments Office](../concepts/add-in-development-best-practices.md)
  - [Concevoir des compléments Office](../design/add-in-design.md)
  - [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
  - [Déployer et publier des compléments Office](../publish/publish.md)
  - [Resources](../resources/resources-links-help.md)
