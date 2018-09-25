---
title: Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication | Microsoft Docs
description: Comment déployer votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2015.
ms.date: 01/25/2018
ms.openlocfilehash: d74ead03b8ac5b7652c7c98851e7e082f4b31ba8
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004916"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="a8e21-103">Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication</span><span class="sxs-lookup"><span data-stu-id="a8e21-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="a8e21-104">Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le complément.</span><span class="sxs-lookup"><span data-stu-id="a8e21-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="a8e21-105">Vous devez publier les fichiers d’application web de votre projet séparément.</span><span class="sxs-lookup"><span data-stu-id="a8e21-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="a8e21-106">Cet article décrit le déploiement de votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="a8e21-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a><span data-ttu-id="a8e21-107">Déploiement de votre projet web à l’aide de Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="a8e21-107">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="a8e21-108">Procédez comme suit pour déployer votre projet web à l’aide de Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="a8e21-108">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="a8e21-109">Dans l’**explorateur de solutions**, ouvrez le menu contextuel du projet de complément, puis sélectionnez **Publier**.</span><span class="sxs-lookup"><span data-stu-id="a8e21-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="a8e21-110">La page **Publier votre complément** s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="a8e21-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="a8e21-111">Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau…** pour créer un profil.</span><span class="sxs-lookup"><span data-stu-id="a8e21-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="a8e21-112">Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.</span><span class="sxs-lookup"><span data-stu-id="a8e21-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="a8e21-113">Si vous choisissez **Nouveau...**, l’Assistant Créer un profil de publication s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="a8e21-113">If you choose  New ..., the  Create publishing profile wizard appears.</span></span> <span data-ttu-id="a8e21-114">Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.</span><span class="sxs-lookup"><span data-stu-id="a8e21-114">You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="a8e21-115">Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="a8e21-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="a8e21-116">Sur la page  **Publier votre complément**, cliquez sur le lien  **Déployer votre projet Web**.</span><span class="sxs-lookup"><span data-stu-id="a8e21-116">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="a8e21-p103">La boîte de dialogue **Publier Web** apparaît. Pour plus d’information sur l’utilisation de cet assistant, reportez-vous à l’article [Procédure : Déployer un projet d’application Web à l’aide de la publication en un clic dans Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="a8e21-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a><span data-ttu-id="a8e21-119">Création d’un package de votre complément avec Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="a8e21-119">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="a8e21-120">Procédez comme suit pour créer un package de votre projet de complément à l’aide de Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="a8e21-120">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="a8e21-121">Sur la page **Publier votre complément**, cliquez sur le lien **Empaqueter le complément**.</span><span class="sxs-lookup"><span data-stu-id="a8e21-121">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="a8e21-122">L’assistant de Publication des compléments SharePoint et Office apparaît.</span><span class="sxs-lookup"><span data-stu-id="a8e21-122">The Publish Office and SharePoint Add-ins wizard appears.</span></span>
    
2. <span data-ttu-id="a8e21-123">Dans la liste déroulante **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL HTTPS du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur **Terminer**.</span><span class="sxs-lookup"><span data-stu-id="a8e21-123">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> 
    
    <span data-ttu-id="a8e21-p104">Vous devez spécifier une URL qui commence par le préfixe HTTPS pour terminer cet assistant. Si vous souhaitez utiliser un point de terminaison HTTP pour votre site web, vous pouvez ouvrir le fichier manifeste XML dans un éditeur de texte une fois que le package a été créé et remplacer le préfixe HTTPS de votre site web par un préfixe HTTP.</span><span class="sxs-lookup"><span data-stu-id="a8e21-p104">You must specify a URL that begins with the HTTPS prefix to complete this wizard. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="a8e21-126"> Les sites web Azure fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="a8e21-126">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="a8e21-127">Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication.</span><span class="sxs-lookup"><span data-stu-id="a8e21-127">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span> 
    
<span data-ttu-id="a8e21-p105">Si vous prévoyez de soumettre votre complément à AppSource, vous pouvez sélectionner le lien **Effectuer la vérification de la validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez régler tous ces problèmes avant de soumettre votre complément au magasin.</span><span class="sxs-lookup"><span data-stu-id="a8e21-p105">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="a8e21-p106">Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="a8e21-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="a8e21-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a8e21-133">See also</span></span>

- [<span data-ttu-id="a8e21-134">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="a8e21-134">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="a8e21-135">Mise à disposition de vos solutions sur AppSource et dans Office</span><span class="sxs-lookup"><span data-stu-id="a8e21-135">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
