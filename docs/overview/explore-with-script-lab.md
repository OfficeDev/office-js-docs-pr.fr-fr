---
title: Explorer l’API JavaScript Office à l’aide de Script Lab
description: Utilisez Script Lab pour explorer l’API JS Office et pour prototyper les fonctionnalités.
ms.date: 01/14/2022
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 568628f60b36a3c6bcb7dd3ab3856e86d96956b7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744630"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Explorer l’API JavaScript Office à l’aide de Script Lab

Les compléments [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) et [Script Lab pour Outlook](https://appsource.microsoft.com/product/office/WA200001603), disponibles gratuitement depuis AppSource, vous permettent d’explorer l’API JavaScript Office lorsque vous travaillez dans un programme Office tel qu’Excel ou Outlook. Script Lab est un outil pratique à ajouter à votre kit de ressources de développement lorsque vous réalisez un prototype et vérifiez les fonctionnalités souhaitées dans votre complément.

## <a name="what-is-script-lab"></a>Qu’est-ce que script Lab ?

Script Lab est un outil destiné à toute personne souhaitant en savoir plus sur la manière de développer des compléments Office à l’aide de l’API JavaScript Office dans Excel, Outlook, Word et PowerPoint. Il fournit IntelliSense, si bien que vous pouvez voir ce qui est disponible et qui repose sur l’infrastructure de Monaco, l’infrastructure utilisée par Visual Studio Code. Via Script Lab, vous pouvez accéder à une bibliothèque d'exemples pour essayer rapidement des fonctionnalités ou utiliser un exemple comme point de départ pour votre propre code. Vous pouvez même utiliser Script Lab pour essayer les API d’aperçu.

Cela vous séduit ? Visionnez cette vidéo d’une minute pour découvrir Script Lab en action.

[![Vidéo d’aperçu montrant l’exécution d’un Script Lab dans Excel, Word et PowerPoint.](../images/screenshot-wide-youtube.png 'Vidéo de la version préliminaire de Script Lab')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>Principales fonctionnalités

Script Lab propose de nombreuses fonctionnalités pour vous aider à explorer l’API JavaScript Office et la fonctionnalité de complément prototype.

### <a name="explore-samples"></a>Explorer des exemples

Commencez rapidement avec une collection d’exemples d’extraits de code intégrés qui montrent comment effectuer des tâches avec l’API. Vous pouvez exécuter les exemples pour afficher instantanément le résultat dans le volet des tâches ou le document, examiner les exemples pour découvrir le fonctionnement de l’API, voire utiliser les exemples pour prototyper votre propre complément.

![Exemples.](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>Code et style

En plus du code JavaScript ou TypeScript qui appelle l’API Office JS, chaque extrait de code contient également une balise HTML qui définit le contenu du volet des tâches et CSS qui définit l’apparence de ce dernier. Vous pouvez personnaliser la balise HTML et CSS pour tester le placement des éléments et les styles lorsque vous prototypez la conception du volet des tâches pour votre propre complément.

> [!TIP]
> Pour appeler des API d’aperçu dans un extrait de code, vous devez mettre à jour les bibliothèques de l’extrait de code pour utiliser le réseau de distribution de contenu bêta (CDN) (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) et les définitions de type d’aperçu `@types/office-js-preview`. De plus, certaines API d’aperçu sont accessibles uniquement si vous êtes inscrit au [programme Office Insider](https://insider.office.com) et que vous exécutez une version Insider d’Office.

### <a name="save-and-share-snippets"></a>Enregistrer et partager des extraits de code

Par défaut, les extraits de code que vous ouvrez dans Script Lab sont enregistrés dans le cache de votre navigateur. Pour enregistrer définitivement un extrait de code, vous pouvez l’exporter dans un contenu [Gist GitHub](https://gist.github.com). Créez un contenu Gist secret pour enregistrer un extrait de code exclusivement pour votre usage personnel ou créez un contenu Gist public si vous envisagez de le partager avec d’autres personnes.

![Options de partage.](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>Importer des extraits de code

Vous pouvez importer un extrait de code dans Script Lab en spécifiant l’URL du [contenu Gist GitHub](https://gist.github.com) public où le YAML de l’extrait de code est stocké ou en collant dans le YAML complet de l’extrait de code. Cette fonctionnalité peut être utile dans les cas où quelqu’un d’autre a partagé son extrait de code avec vous, soit en le publiant dans un contenu Gist GitHub, soit en fournissant le YAML de son extrait de code.

![Option Importer l’extrait de code.](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>Clients pris en charge

Script Lab est pris en charge pour Excel, Word et PowerPoint sur les clients suivants.

- Abonnement Microsoft 365 Office
- Office 2016 ou version ultérieure sous Mac
- Office sur le web

Script Lab pour Outlook est disponible sur les clients suivants.

- Abonnement Microsoft 365 Office
- Outlook 2016 ou version ultérieure sur Mac
- Outlook sur le web avec des navigateurs Chrome, Microsoft Edge ou Safari

Pour obtenir plus d’informations sur Script Lab pour Outlook, voir le [billet de blog](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/) associé.

> [!IMPORTANT]
> Script Lab ne fonctionne plus avec des combinaisons de plate-forme et de version Office qui utilisent Internet Explorer pour héberger des compléments. Cela inclut les versions d’Office à achat unique via Office 2019. Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="next-steps"></a>Étapes suivantes

Pour utiliser Script Lab dans Excel, Word ou PowerPoint, installez le [complément Script Lab](https://appsource.microsoft.com/product/office/WA104380862) à partir d’AppSource.

Pour utiliser le Script Lab pour Outlook, vous pouvez installer le [complément Script Lab pour Outlook](https://appsource.microsoft.com/product/office/wa200001603) depuis AppSource.

Nous vous invitons à développer l’exemple de bibliothèque dans Script Lab en apportant de nouveaux extraits de code dans le référentiel GitHub [Office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).

Lorsque vous êtes prêt à créer votre premier complément Office, essayez le guide de démarrage rapide pour [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md) ou [Project](../quickstarts/project-quickstart.md).

## <a name="see-also"></a>Voir aussi

- [Obtenir un Script Lab pour Excel, Word ou PowerPoint](https://appsource.microsoft.com/product/office/WA104380862)
- [Obtenir un Script Lab pour Outlook](https://appsource.microsoft.com/product/office/wa200001603)
- [Si vous souhaitez en savoir plus sur Script Lab](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Rejoignez le programme développeur de Microsoft 365](https://developer.microsoft.com/office/dev-program)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
