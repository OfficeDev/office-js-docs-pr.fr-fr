---
title: Compléments Office de contenu
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f2632e94e0a797836f73caf0d53fdc0f24bd6790
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703916"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="f086c-102">Compléments Office de contenu</span><span class="sxs-lookup"><span data-stu-id="f086c-102">Content Office Add-ins</span></span>

<span data-ttu-id="f086c-p101">Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Word, Excel ou PowerPoint. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document.</span><span class="sxs-lookup"><span data-stu-id="f086c-p101">Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="f086c-106">*Figure 1. Mise en page type pour les compléments de contenu*</span><span class="sxs-lookup"><span data-stu-id="f086c-106">*Figure 1. Typical layout for content add-ins*</span></span>

![Exemple d’image affichant une mise en page typique pour des compléments de contenu.](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="f086c-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="f086c-108">Best practices</span></span>

- <span data-ttu-id="f086c-109">Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.</span><span class="sxs-lookup"><span data-stu-id="f086c-109">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="f086c-110">Inclure un élément de la marque tel que le BrandBar en bas de votre complément (s’applique aux compléments Word, Excel et PowerPoint uniquement).</span><span class="sxs-lookup"><span data-stu-id="f086c-110">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="f086c-111">Variantes</span><span class="sxs-lookup"><span data-stu-id="f086c-111">Variants</span></span>

<span data-ttu-id="f086c-112">Les tailles des compléments de contenu pour Word, Excel et PowerPoint dans le bureau Office 2016 et Office 365 sont spécifiées par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f086c-112">Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="f086c-113">Menu Caractéristique</span><span class="sxs-lookup"><span data-stu-id="f086c-113">Personality menu</span></span>

<span data-ttu-id="f086c-p102">Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="f086c-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="f086c-116">Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.</span><span class="sxs-lookup"><span data-stu-id="f086c-116">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="f086c-117">*Figure 2. Menu Caractéristique sur Windows*</span><span class="sxs-lookup"><span data-stu-id="f086c-117">*Figure 2. Personality menu on Windows*</span></span> 

![Image illustrant le menu Caractéristique sur le bureau Windows](../images/personality-menu-win.png)


<span data-ttu-id="f086c-119">Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace occupé à 34 x 32 pixels, comme illustré.</span><span class="sxs-lookup"><span data-stu-id="f086c-119">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="f086c-120">*Figure 3. Menu Caractéristique sur Mac*</span><span class="sxs-lookup"><span data-stu-id="f086c-120">*Figure 3. Personality menu on Mac*</span></span>

![Image illustrant le menu Caractéristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="f086c-122">Implémentation</span><span class="sxs-lookup"><span data-stu-id="f086c-122">Implementation</span></span>

<span data-ttu-id="f086c-123">Pour consulter un exemple qui implémente un complément de contenu, reportez-vous à [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="f086c-123">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="f086c-124">Considérations relatives à la prise en charge</span><span class="sxs-lookup"><span data-stu-id="f086c-124">Support considerations</span></span>
- <span data-ttu-id="f086c-125">Vérifiez si votre complément Office fonctionne sur une [plateforme hôte Office spécifique](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span><span class="sxs-lookup"><span data-stu-id="f086c-125">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="f086c-126">Certains compléments de contenu peuvent exiger que l’utilisateur « approuve » que le complément lise et écrive sur Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f086c-126">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="f086c-127">Vous pouvez décider du [niveau d'autorisations](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) que vous souhaitez attribuer à votre utilisateur dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="f086c-127">You can declare what [level of permissions](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your use to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="f086c-128">Les compléments de contenu sont pris en charge dans Excel et PowerPoint dans Office 2013 et versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="f086c-128">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="f086c-129">Si vous ouvrez un complément dans une version d’Office qui ne prend pas en charge les compléments web Office, le complément s’affichera comme une image.</span><span class="sxs-lookup"><span data-stu-id="f086c-129">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="f086c-130">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f086c-130">See also</span></span>
- [<span data-ttu-id="f086c-131">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="f086c-131">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="f086c-132">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f086c-132">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="f086c-133">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="f086c-133">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [<span data-ttu-id="f086c-134">Demande d’autorisations d’utilisation de l’API dans des compléments de contenu et de volet des tâches</span><span class="sxs-lookup"><span data-stu-id="f086c-134">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
