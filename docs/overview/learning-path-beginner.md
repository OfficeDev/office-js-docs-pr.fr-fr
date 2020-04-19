---
title: Commencez ici ! Un guide pour les débutants qui créent des compléments Office
description: Un parcours recommandé pour les débutants à travers les ressources d’apprentissage pour les compléments Office.
ms.date: 04/16/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 026f90ea62960cbbf5ab4420d40a4a9165139cae
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547618"
---
# <a name="start-here-a-guide-for-beginners-making-office-add-ins"></a>Commencez ici ! Un guide pour les débutants qui créent des compléments Office

Vous voulez commencer à créer vos propres extensions Office sur plusieurs plateformes ? La procédure suivante vous montre ce qu’il convient de lire en premier, quels outils installer et quels didacticiels il est recommandé de suivre.

## <a name="step-0-prerequisites"></a>Étape 0 : Conditions requises

- Les compléments Office sont avant tout des applications web incorporées dans Office. Vous devez donc d’abord posséder une connaissance de base des applications web et de la façon dont elles sont hébergées sur le web. Il existe une quantité considérable d’informations à ce sujet sur Internet, dans les livres et dans les cours en ligne. Une bonne façon de commencer, si vous ne possédez aucune connaissance préalable des applications web, consiste à rechercher « Qu’est-ce qu’une application web ? » sur Bing.
- Le langage de programmation principal que vous utiliserez pour créer des compléments Office est JavaScript ou TypeScript. Vous pouvez considérer le langage TypeScript comme une version fortement typée de JavaScript. Si vous n’êtes pas familiarisé avec l’un ou l’autre de ces langages, mais que vous avez de l’expérience avec les langages VBA, VB.Net et C#, vous trouverez probablement TypeScript plus facile à apprendre. Là encore, il existe une multitude d’informations relatives à ces langages sur Internet, dans les livres et dans les cours en ligne.

## <a name="step-1-begin-with-fundamentals"></a>Étape 1 : Commencer par les notions de base

Nous savons que vous êtes impatient de commencer à coder, mais il convient de lire certains points concernant les compléments Office avant d’ouvrir votre IDE ou votre éditeur de code.

- [Vue d’ensemble de la plateforme des compléments Office](office-add-ins.md) : découvrez les compléments web Office et leurs différences par rapport aux anciennes méthodes d’extension d’Office, telles que les compléments VSTO.
- [Création de compléments Office](office-add-ins-fundamentals.md) : obtenez une vue d’ensemble du développement et du cycle de vie des compléments Office, y compris les outils, la création d’une interface utilisateur de complément et l’utilisation des API JavaScript pour interagir avec le document Office.

Ces articles comportent un grand nombre de liens. Toutefois, si vous êtes débutant avec les compléments Office, nous vous recommandons de revenir ici lorsque vous les aurez lus et de passer à la section suivante.

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>Étape 2 : Installer les outils et créer votre premier complément

Vous avez maintenant une vue d’ensemble, alors lancez-vous avec l’un de nos guides de démarrage rapide. Pour découvrir la plateforme, nous vous recommandons d’utiliser le guide de démarrage rapide d’Excel. Il existe une version fondée sur Visual Studio et une autre sur Node.js et Visual Studio Code.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js et Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>Étape 3 : Code

Vous ne pouvez pas apprendre à conduire en lisant le manuel du propriétaire, alors commencez à coder à l’aide de ce [didacticiel Excel](../tutorials/excel-tutorial.md). Vous utiliserez la bibliothèque JavaScript pour Office et du code XML dans le manifeste du complément. Il n’est pas nécessaire de mémoriser quoi que ce soit, car vous obtiendrez plus d’informations sur ces deux éléments plus tard.

## <a name="step-4-understand-the-javascript-library"></a>Étape 4 : Comprendre la bibliothèque JavaScript

Tout d’abord, obtenez une vue d’ensemble de la bibliothèque JavaScript pour Office avec ce didacticiel de Microsoft Learn : [Comprendre les API JavaScript pour Office](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).

Explorez ensuite les API JavaScript pour Office à l’aide de notre [outil Script Lab](explore-with-script-lab.md), un bac à sable pour l’exécution et l’exploration des API.

## <a name="step-5-understand-the-manifest"></a>Étape 5 : Comprendre le manifeste

Découvrez les objectifs du manifeste du complément et consultez une présentation de ses balisages XML sur la page [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

## <a name="next-steps"></a>Étapes suivantes

Félicitations pour avoir terminé le parcours d’apprentissage pour les débutants pour les compléments Office ! Voici quelques suggestions pour approfondir les informations contenues dans notre documentation :

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
