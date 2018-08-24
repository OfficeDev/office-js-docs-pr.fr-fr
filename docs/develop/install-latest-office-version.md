---
title: Installer la dernière version d’Office 2016
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925233"
---
# <a name="install-the-latest-version-of-office-2016"></a><span data-ttu-id="492d7-102">Installer la dernière version d’Office 2016</span><span class="sxs-lookup"><span data-stu-id="492d7-102">Install the latest version of Office 2016</span></span>

<span data-ttu-id="492d7-103">De nouvelles fonctionnalités de développeur, y compris celles en version d’évaluation, sont mises à la disposition des abonnés qui souhaitent obtenir les dernières versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="492d7-103">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="492d7-104">Inscription pour l’obtention des versions les plus récentes</span><span class="sxs-lookup"><span data-stu-id="492d7-104">Opt in to getting the latest builds</span></span>

<span data-ttu-id="492d7-105">Pour s’inscrire afin d’obtenir les dernières versions d’Office 2016, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="492d7-105">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="492d7-106">Si vous êtes abonné à Office 365 Famille, Personnel ou Université, consultez la page [Participez au programme Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="492d7-106">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="492d7-107">Si vous êtes un client d’Office 365 pour les entreprises, consultez l’article [Installer la version First Release pour Office 365 pour les entreprises](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="492d7-107">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="492d7-108">Si vous exécutez Office 2016 sur un Mac :</span><span class="sxs-lookup"><span data-stu-id="492d7-108">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="492d7-109">Démarrez un programme Office 2016 pour Mac.</span><span class="sxs-lookup"><span data-stu-id="492d7-109">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="492d7-110">Sélectionnez **Vérifier les mises à jour** dans le menu Aide.</span><span class="sxs-lookup"><span data-stu-id="492d7-110">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="492d7-111">Dans la zone Mise à jour automatique Microsoft (AutoUpdate), cochez la case pour participer au programme Office Insider.</span><span class="sxs-lookup"><span data-stu-id="492d7-111">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="492d7-112">Obtention de la dernière version</span><span class="sxs-lookup"><span data-stu-id="492d7-112">Get the latest build</span></span>

<span data-ttu-id="492d7-113">Pour obtenir la dernière version d’Office 2016, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="492d7-113">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="492d7-114">Téléchargez l’[outil Déploiement d’Office 2016](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="492d7-114">Download the [Office 2016 Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span> 
2. <span data-ttu-id="492d7-p101">Exécutez l’outil. Cette opération extrait les deux fichiers suivants : Setup.exe et configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="492d7-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="492d7-117">Remplacez le fichier configuration.xml par le [fichier de configuration First Release](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="492d7-117">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="492d7-118">En tant qu’administrateur, exécutez la commande suivante : `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="492d7-118">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="492d7-119">L’exécution de la commande peut durer plusieurs minutes sans vous en indiquer la progression.</span><span class="sxs-lookup"><span data-stu-id="492d7-119">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="492d7-p102">Une fois le processus d’installation terminé, les dernières applications d’Office 2016 sont installées. Pour vérifier que la dernière version est bien installée, accédez à **Fichier**  >  **Compte** à partir de n’importe quelle application Office. Sous Mises à jour Office, vous verrez la mention (Office Insiders) au-dessus du numéro de version.</span><span class="sxs-lookup"><span data-stu-id="492d7-p102">When the installation process finishes, you will have the latest Office 2016 applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Capture d’écran affichant les informations du produit avec la mention Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="492d7-124">Builds Office minimum pour les ensembles de conditions requises pour l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="492d7-124">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="492d7-125">Pour plus d’informations sur les versions minimum des produits pour chaque plate-forme pour les ensembles de conditions requises pour les API, voir les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="492d7-125">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="492d7-126">Ensembles de conditions requises de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="492d7-126">Word JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="492d7-127">Ensembles de conditions requises de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="492d7-127">Excel JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="492d7-128">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="492d7-128">OneNote JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="492d7-129">Ensembles de conditions requises de l’API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="492d7-129">Dialog API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="492d7-130">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="492d7-130">Office common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
