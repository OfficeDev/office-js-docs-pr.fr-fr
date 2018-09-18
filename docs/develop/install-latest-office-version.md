---
title: Installer la dernière version d’Office
description: Informations relatives au choix des dernières versions de Microsoft Office.
ms.date: 12/04/2017
ms.openlocfilehash: 14e26d9fa9f7ec3b2724cbf2e9787cde9dbe4094
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943879"
---
# <a name="install-the-latest-version-of-office"></a>Installer la dernière version d’Office

De nouvelles fonctionnalités pour développeur, y compris celles en préversion, sont d'abord mises à la disposition des abonnés qui choisissent de s'inscrire pour obtenir les dernières versions d’Office. 

## <a name="opt-in-to-getting-the-latest-builds"></a>Inscription pour l’obtention des versions les plus récentes

Pour s’inscrire afin d’obtenir les dernières versions d’Office, procédez comme suit : 

- Si vous êtes abonné à Office 365 Famille, Personnel ou Université, consultez la page [Participez au programme Office Insider](https://products.office.com/office-insider).
- Si vous êtes un client d’Office 365 pour les entreprises, consultez l’article [Installer la version First Release pour Office 365 pour les entreprises](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Si vous exécutez Office sur un Mac :
    - Démarrez un programme Office pour Mac.
    - Sélectionnez **Vérifier les mises à jour** dans le menu Aide.
    - Dans la zone Mise à jour automatique Microsoft (AutoUpdate), cochez la case pour participer au programme Office Insider. 

## <a name="get-the-latest-build"></a>Obtention de la dernière version

Pour obtenir la dernière version d’Office, procédez comme suit : 

1. Téléchargez l’outil [Déploiement d’Office](https://www.microsoft.com/download/details.aspx?id=49117). 
2. Exécutez l’outil. Cette opération extrait deux fichiers : Setup.exe et configuration.xml.
3. Remplacez le fichier configuration.xml par le [fichier de configuration First Release](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. En tant qu’administrateur, exécutez la commande suivante :  `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > L’exécution de la commande peut durer plusieurs minutes sans indication de progression.

Une fois le processus d’installation terminé, les applications Office les plus récentes seront installées. Pour vérifier que vous disposez de la dernière version, accédez à **Fichier** > **Compte** à partir de n’importe quelle application Office. Sous Mises à jour Office, vous verrez l’étiquette (Office Insiders) au-dessus du numéro de version.

![Capture d’écran affichant les informations du produit avec la mention Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Builds Office minimum pour les ensembles de conditions requises pour l’API JavaScript pour Office

Pour plus d’informations sur les versions minimum des produits pour chaque plate-forme pour les ensembles de conditions requises pour les API, voir les rubriques suivantes :

- [Ensembles de conditions requises de l’API JavaScript pour Word](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API de boîte de dialogue](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises des API communes pour Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)
