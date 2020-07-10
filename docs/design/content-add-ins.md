---
title: Compléments Office de contenu
description: Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Excel ou PowerPoint. Ils permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f228ae8e7cca0426b0b43e31e38454029e4c7614
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093846"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="d57c2-103">Compléments Office de contenu</span><span class="sxs-lookup"><span data-stu-id="d57c2-103">Content Office Add-ins</span></span>

<span data-ttu-id="d57c2-104">Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d57c2-104">Content add-ins are surfaces that can be embedded directly into Excel or PowerPoint documents.</span></span> <span data-ttu-id="d57c2-105">Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données.</span><span class="sxs-lookup"><span data-stu-id="d57c2-105">Content add-ins give users access to interface controls that run code to modify documents or display data from a data source.</span></span> <span data-ttu-id="d57c2-106">Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document.</span><span class="sxs-lookup"><span data-stu-id="d57c2-106">Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="d57c2-107">*Figure 1. Mise en page type pour les compléments de contenu*</span><span class="sxs-lookup"><span data-stu-id="d57c2-107">*Figure 1. Typical layout for content add-ins*</span></span>

![Exemple d’image affichant une mise en page typique pour des compléments de contenu.](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="d57c2-109">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="d57c2-109">Best practices</span></span>

- <span data-ttu-id="d57c2-110">Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.</span><span class="sxs-lookup"><span data-stu-id="d57c2-110">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="d57c2-111">Inclure un élément de la marque tel que BrandBar en bas de votre complément (s’applique aux compléments Excel et PowerPoint uniquement).</span><span class="sxs-lookup"><span data-stu-id="d57c2-111">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Excel and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="d57c2-112">Variantes</span><span class="sxs-lookup"><span data-stu-id="d57c2-112">Variants</span></span>

<span data-ttu-id="d57c2-113">Les tailles des compléments de contenu pour Excel et PowerPoint dans le bureau Office et Microsoft 365 sont spécifiées par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d57c2-113">Content add-in sizes for Excel and PowerPoint in Office desktop and Microsoft 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="d57c2-114">Menu Caractéristique</span><span class="sxs-lookup"><span data-stu-id="d57c2-114">Personality menu</span></span>

<span data-ttu-id="d57c2-115">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in.</span><span class="sxs-lookup"><span data-stu-id="d57c2-115">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in.</span></span> <span data-ttu-id="d57c2-116">The following are the current dimensions of the personality menu on Windows and Mac.</span><span class="sxs-lookup"><span data-stu-id="d57c2-116">The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="d57c2-117">Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.</span><span class="sxs-lookup"><span data-stu-id="d57c2-117">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="d57c2-118">*Figure 2. Menu Caractéristique sur Windows*</span><span class="sxs-lookup"><span data-stu-id="d57c2-118">*Figure 2. Personality menu on Windows*</span></span> 

![Image illustrant le menu Caractéristique sur le bureau Windows](../images/personality-menu-win.png)


<span data-ttu-id="d57c2-120">Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace occupé à 34 x 32 pixels, comme illustré.</span><span class="sxs-lookup"><span data-stu-id="d57c2-120">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="d57c2-121">*Figure 3. Menu Caractéristique sur Mac*</span><span class="sxs-lookup"><span data-stu-id="d57c2-121">*Figure 3. Personality menu on Mac*</span></span>

![Image illustrant le menu Caractéristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="d57c2-123">Implémentation</span><span class="sxs-lookup"><span data-stu-id="d57c2-123">Implementation</span></span>

<span data-ttu-id="d57c2-124">Pour consulter un exemple qui implémente un complément de contenu, reportez-vous à [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="d57c2-124">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="d57c2-125">Considérations relatives à la prise en charge</span><span class="sxs-lookup"><span data-stu-id="d57c2-125">Support considerations</span></span>

- <span data-ttu-id="d57c2-126">Vérifiez si votre complément Office fonctionne sur une [plateforme hôte Office spécifique](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="d57c2-126">Check to see if your Office Add-in will work on a [specific Office host platform](../overview/office-add-in-availability.md).</span></span>
- <span data-ttu-id="d57c2-127">Certains compléments de contenu peuvent exiger que l’utilisateur accepte que le complément lise et écrive dans Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d57c2-127">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="d57c2-128">Vous pouvez déclarer le [niveau des autorisations](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) que vous souhaitez attribuer à votre utilisateur dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="d57c2-128">You can declare what [level of permissions](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) you want your user to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="d57c2-129">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span><span class="sxs-lookup"><span data-stu-id="d57c2-129">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="d57c2-130">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span><span class="sxs-lookup"><span data-stu-id="d57c2-130">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="d57c2-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d57c2-131">See also</span></span>

- [<span data-ttu-id="d57c2-132">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="d57c2-132">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="d57c2-133">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="d57c2-133">Office UI Fabric in Office Add-ins</span></span>](../design/office-ui-fabric.md)
- [<span data-ttu-id="d57c2-134">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="d57c2-134">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="d57c2-135">Demande d’autorisations d’utilisation de l’API dans des compléments</span><span class="sxs-lookup"><span data-stu-id="d57c2-135">Requesting permissions for API use in add-ins</span></span>](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
