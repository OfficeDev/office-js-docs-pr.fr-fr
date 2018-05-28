---
title: Cr?er le package de votre compl?ment ? l?aide de Visual Studio pour pr?parer la publication
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e03959294536eeb416a1531d2d281ba83f2d3732
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="aac26-102">Cr?er le package de votre compl?ment ? l?aide de Visual Studio pour pr?parer la publication</span><span class="sxs-lookup"><span data-stu-id="aac26-102">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="aac26-103">Votre package de compl?ment Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le compl?ment.</span><span class="sxs-lookup"><span data-stu-id="aac26-103">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="aac26-104">Vous devez publier les fichiers d?application web de votre projet s?par?ment.</span><span class="sxs-lookup"><span data-stu-id="aac26-104">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="aac26-105">Cet article d?crit le d?ploiement de votre projet web et l?empaquetage de votre compl?ment ? l?aide de Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="aac26-105">This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a><span data-ttu-id="aac26-106">D?ploiement de votre projet web ? l?aide de Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="aac26-106">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="aac26-107">Proc?dez comme suit pour d?ployer votre projet web ? l?aide de Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="aac26-107">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="aac26-108">Dans l?**explorateur de solutions**, ouvrez le menu contextuel du projet de compl?ment, puis s?lectionnez **Publier**.</span><span class="sxs-lookup"><span data-stu-id="aac26-108">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="aac26-109">La page **Publier votre compl?ment** s?ouvre.</span><span class="sxs-lookup"><span data-stu-id="aac26-109">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="aac26-110">Dans la liste d?roulante **Profil actuel**, s?lectionnez un profil ou choisissez **Nouveau?** pour cr?er un profil.</span><span class="sxs-lookup"><span data-stu-id="aac26-110">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="aac26-111">Un profil de publication indique le serveur sur lequel vous effectuez le d?ploiement, les informations d?identification n?cessaires pour se connecter au serveur, les bases de donn?es ? d?ployer, ainsi que d?autres options de d?ploiement.</span><span class="sxs-lookup"><span data-stu-id="aac26-111">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="aac26-p102">Si vous choisissez  **Nouveau...**, l?Assistant **Cr?er un profil de publication** s?ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication ? partir d?un site web d?h?bergement comme Microsoft Azure ou cr?er un profil et ajouter votre serveur, vos informations d?identification et d?autres param?tres, comme d?crit dans la proc?dure suivante.</span><span class="sxs-lookup"><span data-stu-id="aac26-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="aac26-114">Pour plus d?informations sur l?importation et la cr?ation de profils de publication, voir [Cr?ation d?un profil de publication](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="aac26-114">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="aac26-115">Sur la page  **Publier votre compl?ment**, cliquez sur le lien  **D?ployer votre projet Web**.</span><span class="sxs-lookup"><span data-stu-id="aac26-115">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="aac26-p103">La bo?te de dialogue **Publier Web** appara?t. Pour plus d?information sur l?utilisation de cet assistant, reportez-vous ? l?article [Proc?dure?: D?ployer un projet d?application Web ? l?aide de la publication en un clic dans Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="aac26-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a><span data-ttu-id="aac26-118">Cr?ation d?un package de votre compl?ment avec Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="aac26-118">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="aac26-119">Proc?dez comme suit pour cr?er un package de votre projet de compl?ment ? l?aide de Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="aac26-119">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="aac26-120">Sur la page **Publier votre compl?ment**, cliquez sur le lien **Empaqueter le compl?ment**.</span><span class="sxs-lookup"><span data-stu-id="aac26-120">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="aac26-121">L?Assistant **Publication des compl?ments SharePoint et Office** appara?t.</span><span class="sxs-lookup"><span data-stu-id="aac26-121">The **Publish Office and SharePoint Add-ins** wizard appears.</span></span>
    
2. <span data-ttu-id="aac26-122">Dans la liste d?roulante **O? votre site web est-il h?berg? ?**, s?lectionnez ou saisissez l?URL HTTPS du site web qui h?bergera les fichiers de contenu de votre compl?ment, puis cliquez sur **Terminer**.</span><span class="sxs-lookup"><span data-stu-id="aac26-122">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> 
    
    <span data-ttu-id="aac26-p104">Vous devez sp?cifier une URL qui commence par le pr?fixe HTTPS pour terminer cet assistant. Si vous souhaitez utiliser un point de terminaison HTTP pour votre site web, vous pouvez ouvrir le fichier manifeste XML dans un ?diteur de texte une fois que le package a ?t? cr?? et remplacer le pr?fixe HTTPS de votre site web par un pr?fixe HTTP.</span><span class="sxs-lookup"><span data-stu-id="aac26-p104">You must specify a URL that begins with the HTTPS prefix to complete this wizard. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="aac26-125"> Les sites Web Azure fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="aac26-125">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="aac26-126">Visual Studio g?n?re les fichiers n?cessaires ? la publication de votre compl?ment, puis ouvre le dossier de sortie de publication.</span><span class="sxs-lookup"><span data-stu-id="aac26-126">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span> 
    
<span data-ttu-id="aac26-p105">Si vous pr?voyez de soumettre votre compl?ment ? AppSource, vous pouvez s?lectionner le lien **Effectuer la v?rification de la validation** pour identifier les probl?mes susceptibles d?emp?cher votre compl?ment d??tre accept?. Vous devez r?gler tous ces probl?mes avant de soumettre votre compl?ment au magasin.</span><span class="sxs-lookup"><span data-stu-id="aac26-p105">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="aac26-p106">Vous pouvez d?sormais t?l?charger votre manifeste XML ? l?emplacement appropri? pour [publier votre compl?ment](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="aac26-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="aac26-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="aac26-132">See also</span></span>

- [<span data-ttu-id="aac26-133">Publier votre compl?ment Office</span><span class="sxs-lookup"><span data-stu-id="aac26-133">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="aac26-134">Mise ? disposition de vos solutions sur AppSource et dans Office</span><span class="sxs-lookup"><span data-stu-id="aac26-134">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)
    
