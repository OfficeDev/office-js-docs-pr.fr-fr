---
title: Compl?ments Office de contenu
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd0dcea7a3f37175a48946fc9dcd61d2b89f9c08
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="content-office-add-ins"></a><span data-ttu-id="83bee-102">Compl?ments Office de contenu</span><span class="sxs-lookup"><span data-stu-id="83bee-102">Content Office Add-ins</span></span>

<span data-ttu-id="83bee-p101">Les compl?ments de contenu sont des surfaces qui peuvent ?tre incorpor?es directement dans des documents Word, Excel ou PowerPoint. Les compl?ments de contenu permettent aux utilisateurs d?acc?der aux contr?les d?interface qui ex?cutent le code pour modifier des documents ou afficher des donn?es d?une source de donn?es. Utilisez les compl?ments de contenu lorsque vous souhaitez incorporer des fonctionnalit?s directement dans le document.</span><span class="sxs-lookup"><span data-stu-id="83bee-p101">Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="83bee-106">*Figure 1. Mise en page type pour les compl?ments de contenu*</span><span class="sxs-lookup"><span data-stu-id="83bee-106">*Figure 1. Typical layout for content add-ins*</span></span>

![Exemple d?image affichant une mise en page typique pour des compl?ments de contenu.](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="83bee-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="83bee-108">Best practices</span></span>

- <span data-ttu-id="83bee-109">Inclure un ?l?ment de navigation ou de commande comme le CommandBar ou le tableau crois? dynamique en haut de votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="83bee-109">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="83bee-110">Inclure un ?l?ment de la marque tel que le BrandBar en bas de votre compl?ment (s?applique aux compl?ments Word, Excel et PowerPoint uniquement).</span><span class="sxs-lookup"><span data-stu-id="83bee-110">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="83bee-111">Variantes</span><span class="sxs-lookup"><span data-stu-id="83bee-111">Variants</span></span>

<span data-ttu-id="83bee-112">Les tailles des compl?ments de contenu pour Word, Excel et PowerPoint dans le bureau Office 2016 et Office 365 sont sp?cifi?es par l?utilisateur.</span><span class="sxs-lookup"><span data-stu-id="83bee-112">Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="83bee-113">Menu Caract?ristique</span><span class="sxs-lookup"><span data-stu-id="83bee-113">Personality menu</span></span>

<span data-ttu-id="83bee-p102">Les menus Caract?ristique peuvent entraver les ?l?ments de navigation et de commande se trouvant en haut ? droite du compl?ment. Voici les dimensions actuelles du menu Caract?ristique sur Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="83bee-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="83bee-116">Pour Windows, le menu Caract?ristique mesure 12 x 32 pixels, comme illustr?.</span><span class="sxs-lookup"><span data-stu-id="83bee-116">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="83bee-117">*Figure 2. Menu Caract?ristique sur Windows*</span><span class="sxs-lookup"><span data-stu-id="83bee-117">*Figure 2. Personality menu on Windows*</span></span> 

![Image illustrant le menu Caract?ristique sur le bureau Windows](../images/personality-menu-win.png)


<span data-ttu-id="83bee-119">Pour Mac, le menu Caract?ristique mesure 26 x 26 pixels, mais flotte ? 8 pixels de la droite et ? 6 pixels du haut, ce qui permet d?augmenter l?espace occup? ? 34 x 32 pixels, comme illustr?.</span><span class="sxs-lookup"><span data-stu-id="83bee-119">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="83bee-120">*Figure 3. Menu Caract?ristique sur Mac*</span><span class="sxs-lookup"><span data-stu-id="83bee-120">*Figure 3. Personality menu on Mac*</span></span>

![Image illustrant le menu Caract?ristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="83bee-122">Impl?mentation</span><span class="sxs-lookup"><span data-stu-id="83bee-122">Implementation</span></span>

<span data-ttu-id="83bee-123">Pour consulter un exemple qui impl?mente un compl?ment de contenu, reportez-vous ? [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="83bee-123">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="83bee-124">Consid?rations relatives ? la prise en charge</span><span class="sxs-lookup"><span data-stu-id="83bee-124">Support considerations</span></span>
- <span data-ttu-id="83bee-125">V?rifiez si votre compl?ment Office fonctionne sur une [plateforme h?te Office sp?cifique](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability).</span><span class="sxs-lookup"><span data-stu-id="83bee-125">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="83bee-126">Certains compl?ments de contenu peuvent exiger que l?utilisateur accepte que le compl?ment lise et ?crive sur Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="83bee-126">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="83bee-127">Vous pouvez d?clarer le [niveau des autorisations](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) que vous souhaitez attribuer ? votre utilisateur dans le manifeste du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="83bee-127">You can declare what [level of permissions](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your use to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="83bee-128">Les compl?ments de contenu sont pris en charge dans Excel et PowerPoint dans Office 2013 et les versions ult?rieures.</span><span class="sxs-lookup"><span data-stu-id="83bee-128">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="83bee-129">Si vous ouvrez un compl?ment dans une version d?Office qui ne prend pas en charge les compl?ments web Office, le compl?ment s?affichera comme une image.</span><span class="sxs-lookup"><span data-stu-id="83bee-129">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="83bee-130">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="83bee-130">See also</span></span>
- [<span data-ttu-id="83bee-131">Disponibilit? des compl?ments Office sur les plateformes et les h?tes</span><span class="sxs-lookup"><span data-stu-id="83bee-131">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="83bee-132">Office UI Fabric dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="83bee-132">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="83bee-133">Mod?les de conception de l?exp?rience utilisateur pour les compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="83bee-133">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/ux-design-patterns)
- [<span data-ttu-id="83bee-134">Demande d?autorisations d?utilisation de l?API dans des compl?ments de contenu et de volet des t?ches</span><span class="sxs-lookup"><span data-stu-id="83bee-134">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
