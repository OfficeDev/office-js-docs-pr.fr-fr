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
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="d95ae-103">Installer la dernière version d’Office</span><span class="sxs-lookup"><span data-stu-id="d95ae-103">Install the latest version of Office</span></span>

<span data-ttu-id="d95ae-104">De nouvelles fonctionnalités pour développeur, y compris celles en préversion, sont d'abord mises à la disposition des abonnés qui choisissent de s'inscrire pour obtenir les dernières versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="d95ae-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="d95ae-105">Inscription pour l’obtention des versions les plus récentes</span><span class="sxs-lookup"><span data-stu-id="d95ae-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="d95ae-106">Pour s’inscrire afin d’obtenir les dernières versions d’Office, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="d95ae-106">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="d95ae-107">Si vous êtes abonné à Office 365 Famille, Personnel ou Université, consultez la page [Participez au programme Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="d95ae-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="d95ae-108">Si vous êtes un client d’Office 365 pour les entreprises, consultez l’article [Installer la version First Release pour Office 365 pour les entreprises](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="d95ae-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="d95ae-109">Si vous exécutez Office sur un Mac :</span><span class="sxs-lookup"><span data-stu-id="d95ae-109">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="d95ae-110">Démarrez un programme Office pour Mac.</span><span class="sxs-lookup"><span data-stu-id="d95ae-110">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="d95ae-111">Sélectionnez **Vérifier les mises à jour** dans le menu Aide.</span><span class="sxs-lookup"><span data-stu-id="d95ae-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="d95ae-112">Dans la zone Mise à jour automatique Microsoft (AutoUpdate), cochez la case pour participer au programme Office Insider.</span><span class="sxs-lookup"><span data-stu-id="d95ae-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="d95ae-113">Obtention de la dernière version</span><span class="sxs-lookup"><span data-stu-id="d95ae-113">Get the latest build</span></span>

<span data-ttu-id="d95ae-114">Pour obtenir la dernière version d’Office, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="d95ae-114">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="d95ae-115">Téléchargez l’outil [Déploiement d’Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="d95ae-115">Download the Office Deployment Tool</span></span> 
2. <span data-ttu-id="d95ae-p101">Exécutez l’outil. Cette opération extrait deux fichiers : Setup.exe et configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="d95ae-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="d95ae-118">Remplacez le fichier configuration.xml par le [fichier de configuration First Release](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="d95ae-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="d95ae-119">En tant qu’administrateur, exécutez la commande suivante :  `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="d95ae-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="d95ae-120">L’exécution de la commande peut durer plusieurs minutes sans indication de progression.</span><span class="sxs-lookup"><span data-stu-id="d95ae-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="d95ae-121">Une fois le processus d’installation terminé, les applications Office les plus récentes seront installées.</span><span class="sxs-lookup"><span data-stu-id="d95ae-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="d95ae-122">Pour vérifier que vous disposez de la dernière version, accédez à **Fichier** > **Compte** à partir de n’importe quelle application Office.</span><span class="sxs-lookup"><span data-stu-id="d95ae-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="d95ae-123">Sous Mises à jour Office, vous verrez l’étiquette (Office Insiders) au-dessus du numéro de version.</span><span class="sxs-lookup"><span data-stu-id="d95ae-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Capture d’écran affichant les informations du produit avec la mention Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="d95ae-125">Builds Office minimum pour les ensembles de conditions requises pour l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="d95ae-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="d95ae-126">Pour plus d’informations sur les versions minimum des produits pour chaque plate-forme pour les ensembles de conditions requises pour les API, voir les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="d95ae-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="d95ae-127">Ensembles de conditions requises de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="d95ae-127">Word JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [<span data-ttu-id="d95ae-128">Ensembles de conditions requises de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="d95ae-128">Excel JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [<span data-ttu-id="d95ae-129">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="d95ae-129">OneNote JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [<span data-ttu-id="d95ae-130">Ensembles de conditions requises de l’API de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="d95ae-130">Dialog API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="d95ae-131">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="d95ae-131">Office common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)
