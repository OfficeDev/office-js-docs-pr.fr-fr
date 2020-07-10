---
title: Débogage de compléments dans Office sur le web
description: Découvrez comment utiliser Office sur le web pour tester et déboguer vos compléments.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f7ef3fa3d6389629e28b428b9bdbe3b128896b1f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094490"
---
# <a name="debug-add-ins-in-office-on-the-web"></a><span data-ttu-id="f20ef-103">Débogage de compléments dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f20ef-103">Debug add-ins in Office on the web</span></span>

<span data-ttu-id="f20ef-104">Vous pouvez créer et déboguer des compléments sur un ordinateur n’exécutant pas Windows, ou le client de bureau Office 2013 ou Office 2016 (par exemple, si vous développez sur un Mac). Cet article décrit la procédure d’utilisation d’Office Online dans le but de tester et de déboguer vos compléments.</span><span class="sxs-lookup"><span data-stu-id="f20ef-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="f20ef-105">Cet article décrit comment utiliser Office sur le web pour tester et déboguer vos compléments.</span><span class="sxs-lookup"><span data-stu-id="f20ef-105">This article describes how to use Office on the web to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="f20ef-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="f20ef-106">Prerequisites</span></span>

<span data-ttu-id="f20ef-107">Mise en route :</span><span class="sxs-lookup"><span data-stu-id="f20ef-107">To get started:</span></span>

- <span data-ttu-id="f20ef-108">Obtenir un compte de développeur Microsoft 365 si vous n’en avez pas ou n’avez pas accès à un site SharePoint.</span><span class="sxs-lookup"><span data-stu-id="f20ef-108">Get a Microsoft 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>

  > [!NOTE]
  > <span data-ttu-id="f20ef-109">To get a free, 90-day renewable Microsoft 365 developer subscription, join our [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="f20ef-109">To get a free, 90-day renewable Microsoft 365 developer subscription, join our [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span> <span data-ttu-id="f20ef-110">See the [Microsoft 365 developer program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Microsoft 365 developer program and configure your subscription.</span><span class="sxs-lookup"><span data-stu-id="f20ef-110">See the [Microsoft 365 developer program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Microsoft 365 developer program and configure your subscription.</span></span>

- <span data-ttu-id="f20ef-111">Set up an app catalog on SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="f20ef-111">Set up an app catalog on SharePoint Online.</span></span> <span data-ttu-id="f20ef-112">An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library.</span><span class="sxs-lookup"><span data-stu-id="f20ef-112">An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library.</span></span> <span data-ttu-id="f20ef-113">For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span><span class="sxs-lookup"><span data-stu-id="f20ef-113">For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a><span data-ttu-id="f20ef-114">Débogage de compléments à partir d’Excel ou de Word sur le web</span><span class="sxs-lookup"><span data-stu-id="f20ef-114">Debug your add-in from Excel or Word on the web</span></span>

<span data-ttu-id="f20ef-115">Pour déboguer votre complément à l’aide d’Office sur le web, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="f20ef-115">To debug your add-in by using Office on the web:</span></span>

1. <span data-ttu-id="f20ef-116">Déployez votre complément vers un serveur prenant en charge le protocole SSL.</span><span class="sxs-lookup"><span data-stu-id="f20ef-116">Deploy your add-in to a server that supports SSL.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f20ef-117">Nous vous recommandons d’utiliser le [générateur Yeoman](https://github.com/OfficeDev/generator-office) pour créer et héberger votre complément.</span><span class="sxs-lookup"><span data-stu-id="f20ef-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>

2. <span data-ttu-id="f20ef-118">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI.</span><span class="sxs-lookup"><span data-stu-id="f20ef-118">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI.</span></span> <span data-ttu-id="f20ef-119">For example:</span><span class="sxs-lookup"><span data-stu-id="f20ef-119">For example:</span></span>

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. <span data-ttu-id="f20ef-120">Téléchargez le manifeste dans la bibliothèque de compléments Office du catalogue d’applications sur SharePoint.</span><span class="sxs-lookup"><span data-stu-id="f20ef-120">Upload the manifest to the Office Add-ins library in the app catalog on SharePoint.</span></span>

4. <span data-ttu-id="f20ef-121">Lancez Excel ou Word sur le Web à partir du lanceur d’applications dans Microsoft 365, puis ouvrez un nouveau document.</span><span class="sxs-lookup"><span data-stu-id="f20ef-121">Launch Excel or Word on the web from the app launcher in Microsoft 365, and open a new document.</span></span>

5. <span data-ttu-id="f20ef-122">Sur l’onglet Insérer, sélectionnez **Mes compléments** ou **Compléments Office** pour insérer votre complément et le tester dans l’application.</span><span class="sxs-lookup"><span data-stu-id="f20ef-122">On the Insert tab, choose **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>

6. <span data-ttu-id="f20ef-123">Utilisez l’outil de débogage de votre navigateur préféré pour déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="f20ef-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="f20ef-124">Problèmes potentiels</span><span class="sxs-lookup"><span data-stu-id="f20ef-124">Potential issues</span></span>

<span data-ttu-id="f20ef-125">Voici certains problèmes que vous pouvez rencontrer lorsque vous effectuez des opérations de débogage :</span><span class="sxs-lookup"><span data-stu-id="f20ef-125">The following are some issues that you might encounter as you debug:</span></span>

- <span data-ttu-id="f20ef-126">Certaines erreurs JavaScript peuvent provenir d’Office sur le web.</span><span class="sxs-lookup"><span data-stu-id="f20ef-126">Some JavaScript errors that you see might originate from Office on the web.</span></span>

- <span data-ttu-id="f20ef-127">Le navigateur peut afficher une erreur relative à un certificat non valide que vous devrez contourner.</span><span class="sxs-lookup"><span data-stu-id="f20ef-127">The browser might show an invalid certificate error that you will need to bypass.</span></span> <span data-ttu-id="f20ef-128">Le processus d’exécution de cette opération varie en fonction du navigateur et des interfaces utilisateur des différents navigateurs permettant d’effectuer cette modification régulièrement.</span><span class="sxs-lookup"><span data-stu-id="f20ef-128">The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically.</span></span> <span data-ttu-id="f20ef-129">Vous devez effectuer une recherche dans l’aide du navigateur ou rechercher des instructions en ligne.</span><span class="sxs-lookup"><span data-stu-id="f20ef-129">You should search the browser's help or search online for instructions.</span></span> <span data-ttu-id="f20ef-130">(Par exemple, recherchez « Avertissement de certificat Microsoft Edge non valide ».) La plupart des navigateurs, sur la page d’avertissement, comportent un lien qui vous permet d’accéder à la page du complément.</span><span class="sxs-lookup"><span data-stu-id="f20ef-130">(For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page.</span></span> <span data-ttu-id="f20ef-131">Par exemple, Microsoft Edge comporte un lien « Accéder à la page web (non recommandé) ».</span><span class="sxs-lookup"><span data-stu-id="f20ef-131">For example, Microsoft Edge has a link "Go on to the webpage (Not recommended)".</span></span> <span data-ttu-id="f20ef-132">En général, vous devez passer par ce lien chaque fois que le complément est rechargé.</span><span class="sxs-lookup"><span data-stu-id="f20ef-132">But you will usually have to go through this link every time the add-in reloads.</span></span> <span data-ttu-id="f20ef-133">Pour un contournement plus long, consultez l’aide comme suggéré.</span><span class="sxs-lookup"><span data-stu-id="f20ef-133">For a longer lasting bypass, see the help as suggested.</span></span>

- <span data-ttu-id="f20ef-134">Si vous définissez des points d’arrêt dans votre code, Office sur le web peut générer une erreur indiquant qu’il ne peut pas effectuer d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="f20ef-134">If you set breakpoints in your code, Office on the web might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="f20ef-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f20ef-135">See also</span></span>

- [<span data-ttu-id="f20ef-136">Bonnes pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="f20ef-136">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="f20ef-137">Stratégies de validation AppSource</span><span class="sxs-lookup"><span data-stu-id="f20ef-137">AppSource validation policies</span></span>](/legal/marketplace/certification-policies)  
- [<span data-ttu-id="f20ef-138">Création d’applications et de compléments AppSource efficaces</span><span class="sxs-lookup"><span data-stu-id="f20ef-138">Create effective AppSource apps and add-ins</span></span>](/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="f20ef-139">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f20ef-139">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
