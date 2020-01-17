---
title: Explorer l’API JavaScript Office à l’aide de Script Lab
description: Utilisez script Lab pour explorer l’API Office JS et pour prototyper les fonctionnalités.
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
ms.openlocfilehash: 3212aec08cdf4e0185ae5856ae522b1d81e28ea1
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/17/2020
ms.locfileid: "41216972"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Explorer l’API JavaScript Office à l’aide de Script Lab

Le [complément script Lab](https://appsource.microsoft.com/product/office/WA104380862), qui est disponible gratuitement à partir de AppSource, vous permet d’explorer l’API JavaScript Office pendant que vous travaillez dans un programme Office tel qu’Excel ou Word. Script Lab est un outil pratique à ajouter à votre boîte à outils de développement lorsque vous prototypez et vérifiez les fonctionnalités souhaitées dans votre complément.

## <a name="what-is-script-lab"></a>Qu’est-ce que script Lab ?

Script Lab est un outil destiné aux utilisateurs qui souhaitent apprendre à développer des compléments Office à l’aide de l’API JavaScript Office dans Excel, Word ou PowerPoint. Il fournit IntelliSense afin que vous puissiez voir ce qui est disponible et repose sur l’infrastructure Monaco, la même infrastructure utilisée par Visual Studio code. Grâce à script Lab, vous pouvez accéder à une bibliothèque d’exemples pour essayer rapidement des fonctionnalités ou vous pouvez utiliser un exemple comme point de départ pour votre propre code. Vous pouvez même utiliser l’atelier de script pour essayer les API d’aperçu.

Le bruit est-il bien fait ? Jetez un œil à cette vidéo d’une minute pour voir script Lab en action.

[![Vidéo d’aperçu montrant l’exécution d’un Script Lab dans Excel, Word et PowerPoint.](../images/screenshot-wide-youtube.png 'Vidéo de la version préliminaire de Script Lab')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>Principales fonctionnalités

Script Lab offre un certain nombre de fonctionnalités pour vous aider à explorer l’API JavaScript Office et la fonctionnalité de complément prototype.

### <a name="explore-samples"></a>Explorer les exemples

Prise en main rapide avec une collection d’extraits de code intégrés qui montrent comment effectuer des tâches avec l’API. Vous pouvez exécuter les exemples pour voir instantanément le résultat dans le volet Office ou le document, examiner les exemples pour savoir comment fonctionne l’API, et même utiliser des exemples pour prototyper votre propre complément.

![Exemples](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>Code et style

En plus du code JavaScript ou de la machine à écrire qui appelle l’API Office JS, chaque extrait de code contient également un balisage HTML qui définit le contenu du volet de tâches et CSS qui définit l’apparence du volet Office. Vous pouvez personnaliser les balises HTML et CSS pour tester le positionnement et le style des éléments lorsque vous prototypez la conception de volet des tâches pour votre propre complément.

> [!TIP]
> Pour appeler les API d’aperçu dans un extrait de code, vous devez mettre à jour les bibliothèques de l’extrait de`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`code afin d’utiliser la version `@types/office-js-preview`bêta de CDN () et les définitions des types d’aperçu. En outre, certaines API d’aperçu ne sont accessibles que si vous vous êtes inscrit au [programme Office Insider](https://products.office.com/office-insider) et si vous exécutez une version Insider d’Office.

### <a name="save-and-share-snippets"></a>Enregistrer et partager des extraits de code

Par défaut, les extraits de code que vous ouvrez dans script Lab seront enregistrés dans le cache de votre navigateur. Pour enregistrer un extrait de code de manière permanente, vous pouvez l’exporter vers un [GitHub](https://gist.github.com). Créez un annuaire secret pour enregistrer un extrait de code exclusivement pour votre propre usage, ou créez un annuaire public si vous envisagez de le partager avec d’autres personnes.

![Options de partage](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>Importer des extraits de code

Vous pouvez importer un extrait de code dans script Lab en spécifiant l’URL du [GitHub](https://gist.github.com) public où l’extrait de code YAML est stocké ou en collant dans le YAML complet pour l’extrait de code. Cette fonctionnalité peut être utile dans les scénarios où quelqu’un d’autre a partagé son extrait de code avec vous en le publiant dans un GitHub ou en fournissant les YAML de son extrait de code.

![Option Importer un extrait](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>Clients pris en charge

Le script Lab est pris en charge pour Excel, Word et PowerPoint sur les clients suivants.

- Office 2013 ou version ultérieure sur Windows
- Office 2016 ou version ultérieure sur Mac
- Office sur le web

## <a name="next-steps"></a>Étapes suivantes

Pour utiliser script Lab dans Excel, Word ou PowerPoint, installez le [complément script Lab](https://appsource.microsoft.com/product/office/WA104380862) à partir de AppSource. 

Vous pouvez développer l’exemple de bibliothèque dans script Lab en apposant de nouveaux extraits de code dans le référentiel GitHub [Office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) .

Lorsque vous êtes prêt à créer votre premier complément Office, essayez le démarrage rapide pour [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md)ou [Project](../quickstarts/project-quickstart.md).

## <a name="see-also"></a>Voir aussi

- [Obtenir un laboratoire de script](https://appsource.microsoft.com/product/office/WA104380862)
- [En savoir plus sur script Lab](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Rejoindre le programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program)
- [Création de compléments Office](../overview/office-add-ins-fundamentals.md)
