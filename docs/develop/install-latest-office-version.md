---
title: Installer la dernière version d’Office
description: Découvrez comment s’inscrire afin d’obtenir les dernières versions d’Office.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 6c18e91c93d8125c30bd52ffa7ee602c555e53accea108fb6b0645eb03610caf
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080578"
---
# <a name="install-the-latest-version-of-office"></a>Installer la dernière version d’Office

De nouvelles fonctionnalités de développeur, y compris celles en version d’évaluation, sont mises à la disposition des abonnés qui souhaitent obtenir les dernières versions d’Office.

## <a name="opt-in-to-getting-the-latest-builds-of-office"></a>Optez pour obtenir les dernières builds de Office

- Si vous êtes abonné à Microsoft 365 Famille, Personnel ou Université, consultez l’article Office [Insider.](https://insider.office.com)
- Si vous êtes un client Applications Microsoft 365 pour les PME, voir Installer la build [First Release pour Applications Microsoft 365 pour les PME clients.](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)
- Si vous exécutez Office sur un Mac :
  - Démarrez une application Office.
  - Sélectionnez **Vérifier les mises à jour** dans le menu Aide.
  - Dans la zone Mise à jour automatique Microsoft (AutoUpdate), cochez la case pour participer au programme Office Insider.

## <a name="get-the-latest-build-of-office"></a>Obtenir la dernière version de Office

1. Télécharger [l’outil Déploiement d’Office](https://www.microsoft.com/download/details.aspx?id=49117).
2. Exécutez l’outil. Cette opération extrait les deux fichiers suivants : Setup.exe et configuration.xml.
3. Remplacez le fichier configuration.xml par le [fichier de configuration First Release](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. En tant qu’administrateur, exécutez la commande suivante : `setup.exe /configure configuration.xml`

> [!NOTE]
> L’exécution de la commande peut durer plusieurs minutes sans vous en indiquer la progression.

Une fois le processus d’installation terminé, les dernières applications d’Office sont installées. Pour vérifier que la dernière version est bien installée, accédez à **Fichier** > **Compte** à partir de n’importe quelle application Office. Sous Mises à jour Office, vous verrez la mention (Office Insiders) au-dessus du numéro de version.

![Capture d’écran shows product information with the Office Insiders label.](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Builds Office minimum pour les ensembles de conditions requises pour l’API JavaScript pour Office

- [Ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API JavaScript pour Word](../reference/requirement-sets/word-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API de boîte de dialogue](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Ensembles de conditions requises des API communes pour Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
