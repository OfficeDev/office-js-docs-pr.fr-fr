---
title: Débogage de compléments dans Office Online
description: Découvrez comment utiliser Office Online pour tester et déboguer vos compléments.
ms.date: 03/14/2018
ms.openlocfilehash: ee458352c78a3bb7828e66df9fcde12958f3df93
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945763"
---
# <a name="debug-add-ins-in-office-online"></a><span data-ttu-id="535d1-103">Débogage de compléments dans Office Online</span><span class="sxs-lookup"><span data-stu-id="535d1-103">Debug add-ins in Office Online</span></span>


<span data-ttu-id="535d1-104">Vous pouvez créer et déboguer des compléments sur un ordinateur qui n’exécute pas Windows ou le client de bureau Office&mdash;par exemple, si vous développez sur un Mac.</span><span class="sxs-lookup"><span data-stu-id="535d1-104">You can build and debug add-ins on a computer that isn't running Windows or the Office 2013 or Office 2016 desktop client - for example, if you're developing on a Mac. This article describes how to use Office Online to test and debug your add-ins.</span></span> <span data-ttu-id="535d1-105">Cet article explique comment utiliser Office Online pour tester et déboguer vos compléments.</span><span class="sxs-lookup"><span data-stu-id="535d1-105">How to use Office Online to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="535d1-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="535d1-106">Prerequisites</span></span>

<span data-ttu-id="535d1-107">Mise en route :</span><span class="sxs-lookup"><span data-stu-id="535d1-107">To get started:</span></span>

- <span data-ttu-id="535d1-108">Si vous n’en avez pas encore, créez un compte de développeur Office 365, ou accédez à un site SharePoint.</span><span class="sxs-lookup"><span data-stu-id="535d1-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>
    
  > [!NOTE]
  > <span data-ttu-id="535d1-109">Pour vous inscrire et obtenir gratuitement un abonnement Office 365 Développeur, participez à notre [programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="535d1-109">To sign up for a free Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span> <span data-ttu-id="535d1-110">Consultez la [documentation relative au programme pour les développeurs Office 365](https://docs.microsoft.com/office/developer-program/office-365-developer-program) pour obtenir des instructions détaillées sur la façon de rejoindre le programme, de vous inscrire et de configurer votre abonnement.</span><span class="sxs-lookup"><span data-stu-id="535d1-110">See the [Office 365 Developer Program documentation](https://docs.microsoft.com/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription.</span></span>
     
- <span data-ttu-id="535d1-p103">Configurez un catalogue de compléments sur Office 365 (SharePoint Online). Un catalogue de compléments est une collection de sites dédiée dans SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office. Si vous disposez de votre propre site SharePoint, vous pouvez configurer une bibliothèque de document de catalogue de compléments. Pour plus d’informations, voir [Publier des compléments de contenu et du volet Office dans un catalogue de compléments sur SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span><span class="sxs-lookup"><span data-stu-id="535d1-p103">Set up an add-in catalog on Office 365 (SharePoint Online). An add-in catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an add-in catalog document library. For more information, see [Publish task pane and content add-ins to an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a><span data-ttu-id="535d1-114">Débogage de compléments à partir d’Excel Online ou de Word Online</span><span class="sxs-lookup"><span data-stu-id="535d1-114">Debug your add-in from Excel Online or Word Online</span></span>

<span data-ttu-id="535d1-115">Pour déboguer votre complément à l’aide d’Office Online, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="535d1-115">To debug your add-in by using Office Online:</span></span>

1. <span data-ttu-id="535d1-116">Déployez votre complément vers un serveur prenant en charge le protocole SSL.</span><span class="sxs-lookup"><span data-stu-id="535d1-116">Deploy your add-in to a server that supports SSL.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="535d1-117">Nous vous recommandons d’utiliser le [générateur Yeoman](https://github.com/OfficeDev/generator-office) pour créer et héberger votre complément.</span><span class="sxs-lookup"><span data-stu-id="535d1-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>
     
2. <span data-ttu-id="535d1-p104">Dans le [fichier manifeste de votre complément](../develop/add-in-manifests.md), mettez à jour la valeur de l’élément **SourceLocation** afin d’inclure un URI absolu, plutôt que relatif. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="535d1-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. <span data-ttu-id="535d1-120">Téléchargez le manifeste dans la bibliothèque de compléments Office du catalogue de compléments sur SharePoint.</span><span class="sxs-lookup"><span data-stu-id="535d1-120">Upload the manifest to the Office Add-ins library in the add-in catalog on SharePoint.</span></span>
    
4. <span data-ttu-id="535d1-121">Lancez Excel Online ou Word Online à partir du lanceur d’applications dans Office 365, puis ouvrez un nouveau document.</span><span class="sxs-lookup"><span data-stu-id="535d1-121">Launch Excel Online or Word Online from the app launcher in Office 365, and open a new document.</span></span>
    
5. <span data-ttu-id="535d1-122">Sur l’onglet Insérer, sélectionnez  **Mes compléments** ou **Compléments Office** pour insérer votre complément et le tester dans l’application.</span><span class="sxs-lookup"><span data-stu-id="535d1-122">On the Insert tab, choose  **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>
    
6. <span data-ttu-id="535d1-123">Utilisez l’outil de débogage de votre navigateur préféré pour déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="535d1-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="535d1-124">Problèmes potentiels</span><span class="sxs-lookup"><span data-stu-id="535d1-124">Potential issues</span></span>    

<span data-ttu-id="535d1-125">Voici certains problèmes que vous pouvez rencontrer lorsque vous effectuez des opérations de débogage :</span><span class="sxs-lookup"><span data-stu-id="535d1-125">The following are some issues that you might encounter as you debug:</span></span>
    
- <span data-ttu-id="535d1-126">Certaines erreurs JavaScript peuvent provenir d’Office Online.</span><span class="sxs-lookup"><span data-stu-id="535d1-126">Some JavaScript errors that you see might originate from Office Online.</span></span>
      
- <span data-ttu-id="535d1-127">Le navigateur peut afficher une erreur liée à un certificat non valide que vous devrez contourner.</span><span class="sxs-lookup"><span data-stu-id="535d1-127">The browser might show an invalid certificate error that you will need to bypass.</span></span>
      
- <span data-ttu-id="535d1-128">Si vous définissez des points d’arrêt dans votre code, Office Online peut générer une erreur indiquant qu’il ne peut pas effectuer d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="535d1-128">If you set breakpoints in your code, Office Online might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="535d1-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="535d1-129">See also</span></span>

- [<span data-ttu-id="535d1-130">Bonnes pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="535d1-130">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="535d1-131">Stratégies de validation AppSource</span><span class="sxs-lookup"><span data-stu-id="535d1-131">AppSource validation policies</span></span>](https://docs.microsoft.com/office/dev/store/validation-policies)  
- [<span data-ttu-id="535d1-132">Création d’applications et de compléments AppSource efficaces</span><span class="sxs-lookup"><span data-stu-id="535d1-132">Create effective AppSource apps and add-ins</span></span>](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="535d1-133">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="535d1-133">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
    
