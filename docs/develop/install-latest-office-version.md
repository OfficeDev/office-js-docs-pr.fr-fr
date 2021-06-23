---
title: Installer la dernière version d’Office
description: Découvrez comment s’inscrire afin d’obtenir les dernières versions d’Office.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: d72012e3c2e642c74d8573c4d9bb3b29a8fc0274
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076019"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="09c38-103">Installer la dernière version d’Office</span><span class="sxs-lookup"><span data-stu-id="09c38-103">Install the latest version of Office</span></span>

<span data-ttu-id="09c38-104">De nouvelles fonctionnalités de développeur, y compris celles en version d’évaluation, sont mises à la disposition des abonnés qui souhaitent obtenir les dernières versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="09c38-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="09c38-105">Inscription pour l’obtention des versions les plus récentes</span><span class="sxs-lookup"><span data-stu-id="09c38-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="09c38-106">Pour s’inscrire afin d’obtenir les dernières versions d’Office, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="09c38-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="09c38-107">Si vous êtes abonné à Microsoft 365 Famille, Personnel ou Université, consultez l’article Office [Insider.](https://insider.office.com)</span><span class="sxs-lookup"><span data-stu-id="09c38-107">If you're a Microsoft 365 Family, Personal, or University subscriber, see [Be an Office Insider](https://insider.office.com).</span></span>
- <span data-ttu-id="09c38-108">Si vous êtes un client Applications Microsoft 365 pour les PME, voir Installer la build [First Release pour Applications Microsoft 365 pour les PME clients.](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)</span><span class="sxs-lookup"><span data-stu-id="09c38-108">If you're a Microsoft 365 Apps for business customer, see [Install the First Release build for Microsoft 365 Apps for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="09c38-109">Si vous exécutez Office sur un Mac :</span><span class="sxs-lookup"><span data-stu-id="09c38-109">If you're running Office on a Mac:</span></span>
  - <span data-ttu-id="09c38-110">Démarrez une application Office.</span><span class="sxs-lookup"><span data-stu-id="09c38-110">Start an Office application.</span></span>
  - <span data-ttu-id="09c38-111">Sélectionnez **Vérifier les mises à jour** dans le menu Aide.</span><span class="sxs-lookup"><span data-stu-id="09c38-111">Select **Check for Updates** on the Help menu.</span></span>
  - <span data-ttu-id="09c38-112">Dans la zone Mise à jour automatique Microsoft (AutoUpdate), cochez la case pour participer au programme Office Insider.</span><span class="sxs-lookup"><span data-stu-id="09c38-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="09c38-113">Obtention de la dernière version</span><span class="sxs-lookup"><span data-stu-id="09c38-113">Get the latest build</span></span>

<span data-ttu-id="09c38-114">Pour obtenir la dernière version d’Office, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="09c38-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="09c38-115">Télécharger [l’outil Déploiement d’Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="09c38-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="09c38-p101">Exécutez l’outil. Cette opération extrait les deux fichiers suivants : Setup.exe et configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="09c38-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="09c38-118">Remplacez le fichier configuration.xml par le [fichier de configuration First Release](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="09c38-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="09c38-119">En tant qu’administrateur, exécutez la commande suivante : `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="09c38-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

> [!NOTE]
> <span data-ttu-id="09c38-120">L’exécution de la commande peut durer plusieurs minutes sans vous en indiquer la progression.</span><span class="sxs-lookup"><span data-stu-id="09c38-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="09c38-121">Une fois le processus d’installation terminé, les dernières applications d’Office sont installées.</span><span class="sxs-lookup"><span data-stu-id="09c38-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="09c38-122">Pour vérifier que la dernière version est bien installée, accédez à **Fichier** > **Compte** à partir de n’importe quelle application Office.</span><span class="sxs-lookup"><span data-stu-id="09c38-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="09c38-123">Sous Mises à jour Office, vous verrez la mention (Office Insiders) au-dessus du numéro de version.</span><span class="sxs-lookup"><span data-stu-id="09c38-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Capture d’écran shows product information with the Office Insiders label.](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="09c38-125">Builds Office minimum pour les ensembles de conditions requises pour l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="09c38-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="09c38-126">Pour plus d’informations sur les versions minimum des produits pour chaque plate-forme pour les ensembles de conditions requises pour les API, voir les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="09c38-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="09c38-127">Ensembles de conditions requises de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="09c38-127">Excel JavaScript API requirement sets</span></span>](../reference/requirement-sets/excel-api-requirement-sets.md)
- [<span data-ttu-id="09c38-128">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="09c38-128">OneNote JavaScript API requirement sets</span></span>](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [<span data-ttu-id="09c38-129">Ensembles de conditions requises de l’API JavaScript pour Outlook</span><span class="sxs-lookup"><span data-stu-id="09c38-129">Outlook JavaScript API requirement sets</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="09c38-130">Ensembles de conditions requises de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="09c38-130">PowerPoint JavaScript API requirement sets</span></span>](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [<span data-ttu-id="09c38-131">Ensembles de conditions requises de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="09c38-131">Word JavaScript API requirement sets</span></span>](../reference/requirement-sets/word-api-requirement-sets.md)
- [<span data-ttu-id="09c38-132">Ensembles de conditions requises de l’API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="09c38-132">Dialog API requirement sets</span></span>](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [<span data-ttu-id="09c38-133">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="09c38-133">Office Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
