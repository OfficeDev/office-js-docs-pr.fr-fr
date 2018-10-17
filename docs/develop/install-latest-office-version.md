---
title: Installer la dernière version d’Office
description: Informations relatives au choix des dernières versions de Microsoft Office.
ms.date: 12/04/2017
ms.openlocfilehash: 0e6e147144757004575fa086e1066b7cdf133ee8
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505789"
---
# <a name="install-the-latest-version-of-office"></a>Installer la dernière version d’Office

De nouvelles fonctionnalités pour développeur, y compris celles en préversion, sont d'abord mises à la disposition des abonnés qui choisissent de s'inscrire pour obtenir les dernières versions d’Office. 

## <a name="opt-in-to-getting-the-latest-builds"></a>Inscription pour l’obtention des versions les plus récentes

Pour s’inscrire afin d’obtenir les dernières versions d’Office, procédez comme suit : 

- Si vous êtes abonné à Office 365 Famille, Personnel ou Université, consultez la page [Participez au programme Office Insider](https://products.office.com/office-insider).
- Si vous êtes un client d’Office 365 pour les entreprises, consultez l’article [Installer la première version Office 365 pour les entreprises](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Si vous exécutez Office sur un Mac :
    - Démarrez un programme Office pour Mac.
    - Sélectionnez **Vérifier les mises à jour** dans le menu Aide.
    - Dans la zone Mise à jour automatique de Microsoft (AutoUpdate), cochez la case pour rejoindre le programme Office Insider. 

## <a name="get-the-latest-build"></a>Obtenez la dernière version

Pour obtenir la dernière version d’Office, procédez comme suit : 

1. Téléchargez l’outil [Déploiement d’Office](https://www.microsoft.com/download/details.aspx?id=49117). 
2. Exécutez l’outil. Cette opération extrait deux fichiers : Setup.exe et configuration.xml.
3. Remplacez le fichier configuration.xml par le [fichier de configuration de la première version](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. En tant qu’administrateur, exécutez la commande suivante :  `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > L’exécution de la commande peut durer plusieurs minutes sans vous en indiquer la progression.

Une fois le processus d’installation terminé, les dernières applications d’Office 2016 sont installées. Pour vérifier que la dernière version est bien installée, accédez à **Fichier**  >  **Compte** à partir de n’importe quelle application Office. Sous Mises à jour Office, vous verrez la mention (Office Insiders) au-dessus du numéro de version.

![Capture d’écran affichant les informations du produit avec la mention Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Version Office minimum pour les ensembles de conditions requises pour l’API JavaScript pour Office

Pour plus d’informations sur les versions minimum des produits pour chaque plate-forme pour les ensembles de conditions requises pour les API, voir les rubriques suivantes :

- [Ensembles de conditions requises de l’API JavaScript pour Word](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API de Dialog](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises des API communes pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
