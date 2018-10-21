---
title: Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication | Microsoft Docs
description: Comment déployer votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2017.
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681762"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="eb35c-103">Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication</span><span class="sxs-lookup"><span data-stu-id="eb35c-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="eb35c-p101">Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous utiliserez pour publier le complément. Vous devez publier séparément les fichiers d’application web de votre projet. Cet article décrit le déploiement de votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="eb35c-p101">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="eb35c-107">Déploiement de votre projet web à l’aide de Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="eb35c-107">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="eb35c-108">Procédez comme suit pour déployer votre projet web à l’aide de Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="eb35c-108">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="eb35c-109">Dans l’**explorateur de solutions**, ouvrez le menu contextuel du projet de complément, puis sélectionnez **Publier**.</span><span class="sxs-lookup"><span data-stu-id="eb35c-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="eb35c-110">La page **Publier votre complément** s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="eb35c-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="eb35c-111">Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau…** pour créer un profil.</span><span class="sxs-lookup"><span data-stu-id="eb35c-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="eb35c-112">Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.</span><span class="sxs-lookup"><span data-stu-id="eb35c-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="eb35c-p102">Si vous choisissez  **Nouveau...**, l’Assistant **Créer un profil de publication** s’ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.</span><span class="sxs-lookup"><span data-stu-id="eb35c-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="eb35c-115">Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="eb35c-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="eb35c-116">Sur la page  **Publier votre complément**, cliquez sur le lien  **Déployer votre projet Web**.</span><span class="sxs-lookup"><span data-stu-id="eb35c-116">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="eb35c-p103">La boîte de dialogue **Publier** apparaît. Pour plus d’information sur l’utilisation de cet assistant, reportez-vous à l’article [Procédure : Déployer un projet d’application Web à l’aide de la publication en un clic dans Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="eb35c-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="eb35c-119">Création d’un package de votre complément avec Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="eb35c-119">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="eb35c-120">Procédez comme suit pour créer un package de votre projet de complément à l’aide de Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="eb35c-120">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="eb35c-121">Sur la page **Publier votre complément**, cliquez sur le bouton **Empaqueter le complément**.</span><span class="sxs-lookup"><span data-stu-id="eb35c-121">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="eb35c-122">Un Assistant s’affiche avec la page **Empaqueter le complément**.</span><span class="sxs-lookup"><span data-stu-id="eb35c-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="eb35c-123">Dans la liste déroulante  **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur  **Terminer**.</span><span class="sxs-lookup"><span data-stu-id="eb35c-123">In the  **Where is your website hosted?** dropdown list, select or enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="eb35c-124">Les sites web Azure fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="eb35c-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="eb35c-125">Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication.</span><span class="sxs-lookup"><span data-stu-id="eb35c-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="eb35c-126">Si vous prévoyez d’envoyer votre complément à AppSource, vous pouvez sélectionner le bouton **Effectuer la vérification de la validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté.</span><span class="sxs-lookup"><span data-stu-id="eb35c-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="eb35c-127">Vous devez régler tous ces problèmes avant d’envoyer votre complément au magasin.</span><span class="sxs-lookup"><span data-stu-id="eb35c-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="eb35c-p105">Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="eb35c-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="eb35c-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eb35c-131">See also</span></span>

- [<span data-ttu-id="eb35c-132">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="eb35c-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="eb35c-133">Mise à disposition de vos solutions sur AppSource et dans Office</span><span class="sxs-lookup"><span data-stu-id="eb35c-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
