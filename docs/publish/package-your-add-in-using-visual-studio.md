---
title: Publier votre complément à l’aide de Visual Studio
description: Déploiement de votre projet web et création d’un package de votre complément à l’aide de Visual Studio 2019.
ms.date: 12/02/2019
localization_priority: Priority
ms.openlocfilehash: 5da7fc643eb517f777325658d01889f3e51906bd
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670194"
---
# <a name="publish-your-add-in-using-visual-studio"></a><span data-ttu-id="35b4f-103">Publier votre complément à l’aide de Visual Studio</span><span class="sxs-lookup"><span data-stu-id="35b4f-103">Package your add-in using Visual Studio</span></span>

<span data-ttu-id="35b4f-104">Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le complément.</span><span class="sxs-lookup"><span data-stu-id="35b4f-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="35b4f-105">Vous devez publier les fichiers d’application web de votre projet séparément.</span><span class="sxs-lookup"><span data-stu-id="35b4f-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="35b4f-106">Cet article décrit le déploiement de votre projet web et création d’un package de votre complément à l’aide de Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="35b4f-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2019.</span></span>

> [!NOTE]
> <span data-ttu-id="35b4f-107">Pour plus d’informations sur la publication d’un complément Office que vous avez créé à l’aide du générateur Yeoman et développé avec Visual Studio Code ou un autre éditeur, voir [Publier un complément développé avec Visual Studio Code](publish-add-in-vs-code.md).</span><span class="sxs-lookup"><span data-stu-id="35b4f-107">For information about publishing an Office Add-in that you created using the Yeoman generator and developed with Visual Studio Code or any other editor, see [Publish an add-in developed with Visual Studio Code](publish-add-in-vs-code.md).</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="35b4f-108">Pour déployer votre projet web à l’aide de Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="35b4f-108">To deploy your web project using Visual Studio 2019</span></span>

<span data-ttu-id="35b4f-109">Réalisez les étapes suivantes pour déployer votre projet Web à l'aide de Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="35b4f-109">Complete the following steps to deploy your web project using Visual Studio 2019.</span></span>

1. <span data-ttu-id="35b4f-110">Depuis l’onglet **Build**, sélectionnez **Publier [nom de votre complément]**.</span><span class="sxs-lookup"><span data-stu-id="35b4f-110">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="35b4f-111">Dans la fenêtre \*\*Choisir une cible de publication \*\*, sélectionnez une des options pour publier sur votre cible préférée.</span><span class="sxs-lookup"><span data-stu-id="35b4f-111">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="35b4f-112">Chaque cible de publication nécessite que vous incluiez plus d'informations pour commencer, comme l'emplacement d'une machine virtuelle Azure ou d'un emplacement de dossier.</span><span class="sxs-lookup"><span data-stu-id="35b4f-112">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="35b4f-113">Une fois que vous avez spécifié un emplacement de publication et renseigné toutes les informations requises, sélectionnez **Publier**</span><span class="sxs-lookup"><span data-stu-id="35b4f-113">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="35b4f-114">Le choix d’une cible de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.</span><span class="sxs-lookup"><span data-stu-id="35b4f-114">Picking a publish target specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="35b4f-115">Pour plus d’informations sur les étapes de déploiement de chaque option cible de publication, voir [Premier aperçu du déploiement dans Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span><span class="sxs-lookup"><span data-stu-id="35b4f-115">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="35b4f-116">Pour créer un package et publier votre complément à l’aide d’IIS, de FTP ou du déploiement Web à l’aide de Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="35b4f-116">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="35b4f-117">Procédez comme suit pour créer un package de votre complément à l’aide de Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="35b4f-117">Complete the following steps to package your add-in using Visual Studio 2019.</span></span>

1. <span data-ttu-id="35b4f-118">Depuis l’onglet **Build**, sélectionnez **Publier [nom de votre complément]**.</span><span class="sxs-lookup"><span data-stu-id="35b4f-118">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="35b4f-119">Dans la fenêtre **Choisir une cible de publication**, choisissez **IIS, FTP, etc.** et sélectionnez **Configurer**.</span><span class="sxs-lookup"><span data-stu-id="35b4f-119">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="35b4f-120">Sélectionnez ensuite **Publier**.</span><span class="sxs-lookup"><span data-stu-id="35b4f-120">Next, select **Publish**.</span></span>
3. <span data-ttu-id="35b4f-121">Un assistant s’affiche pour vous guider tout au long du processus.</span><span class="sxs-lookup"><span data-stu-id="35b4f-121">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="35b4f-122">Assurez-vous que la méthode de publication est votre méthode préférée, telle que Web Deploy.</span><span class="sxs-lookup"><span data-stu-id="35b4f-122">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="35b4f-123">Dans la zone **URL de destination**, entrez l'URL du site Web qui hébergera les fichiers de contenu de votre complément, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="35b4f-123">In the **Destination URL** box, enter the URL of the website that will host the content files of your add-in, and then select **Next**.</span></span> <span data-ttu-id="35b4f-124">Si vous prévoyez de soumettre votre complément à AppSource, vous pouvez choisir le bouton **Valider la connexion** pour identifier tout problème susceptible d'empêcher votre complément d'être accepté.</span><span class="sxs-lookup"><span data-stu-id="35b4f-124">If you plan to submit your add-in to AppSource, you can choose the **Validate Connection** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="35b4f-125">Vous devez corriger tous les problèmes avant d’envoyer votre complément au Store.</span><span class="sxs-lookup"><span data-stu-id="35b4f-125">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="35b4f-126">Confirmez tous les paramètres souhaités, y compris les **Options de publication de fichiers**, puis sélectionnez **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="35b4f-126">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="35b4f-127">Les sites web Azure [!include[HTTPS guidance](../includes/https-guidance.md)] fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="35b4f-127">[!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="35b4f-p106">Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="35b4f-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="35b4f-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="35b4f-131">See also</span></span>

- [<span data-ttu-id="35b4f-132">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="35b4f-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="35b4f-133">Mise à disposition de vos solutions sur AppSource et dans Office</span><span class="sxs-lookup"><span data-stu-id="35b4f-133">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
