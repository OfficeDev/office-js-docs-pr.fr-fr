---
title: Volets des tâches dans les compléments Office
description: Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cd8d9386fa9f154d611926add12e21f545e36351
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076303"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="81ca1-103">Volets des tâches dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="81ca1-103">Task panes in Office Add-ins</span></span>

<span data-ttu-id="81ca1-p101">Les volets des tâches sont des surfaces d’interface qui s’affichent généralement sur le côté droit de la fenêtre dans Word, PowerPoint, Excel et Outlook. Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données. Utilisez les volets des tâches lorsque vous n’avez pas besoin d’incorporer des fonctionnalités directement dans le document.</span><span class="sxs-lookup"><span data-stu-id="81ca1-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="81ca1-107">*Figure 1. Mise en page type du volet Office*</span><span class="sxs-lookup"><span data-stu-id="81ca1-107">*Figure 1. Typical task pane layout*</span></span>

![Illustration affichant une disposition classique du volet Des tâches avec des onglets de section en haut, le logo de la société et le nom de la société en bas à gauche, et une icône de paramètres en bas à droite.](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="81ca1-109">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="81ca1-109">Best practices</span></span>

|<span data-ttu-id="81ca1-110">À faire</span><span class="sxs-lookup"><span data-stu-id="81ca1-110">Do</span></span>|<span data-ttu-id="81ca1-111">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="81ca1-111">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="81ca1-112">Inclure le nom de votre complément dans le titre.</span><span class="sxs-lookup"><span data-stu-id="81ca1-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="81ca1-113">Ne pas ajouter le nom de votre société au titre.</span><span class="sxs-lookup"><span data-stu-id="81ca1-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="81ca1-114">Utiliser des noms descriptifs courts dans le titre.</span><span class="sxs-lookup"><span data-stu-id="81ca1-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="81ca1-115">N’ajoutez pas de chaînes telles que « add-in », « for Word » ou « for Office » au titre de votre application.</span><span class="sxs-lookup"><span data-stu-id="81ca1-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="81ca1-116">Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.</span><span class="sxs-lookup"><span data-stu-id="81ca1-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="81ca1-117">Inclure un élément de la marque tel que le BrandBar en bas de votre complément, sauf si votre complément doit être utilisé dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="81ca1-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||

## <a name="variants"></a><span data-ttu-id="81ca1-118">Variantes</span><span class="sxs-lookup"><span data-stu-id="81ca1-118">Variants</span></span>

<span data-ttu-id="81ca1-119">Les images suivantes montrent les différentes tailles de volet de tâches avec application Office ruban avec une résolution de 1 366 x 768.</span><span class="sxs-lookup"><span data-stu-id="81ca1-119">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution.</span></span> <span data-ttu-id="81ca1-120">Pour Excel, l’espace vertical supplémentaire est requis pour s’adapter à la barre de formule.</span><span class="sxs-lookup"><span data-stu-id="81ca1-120">For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="81ca1-121">*Figure 2. Tailles de volet des tâches du bureau Office 2016*</span><span class="sxs-lookup"><span data-stu-id="81ca1-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![Diagramme affichant les tailles du volet Des tâches du bureau à une résolution de 1 366 x 768.](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="81ca1-123">Excel - 320 x 455 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-123">Excel - 320x455 pixels</span></span>
- <span data-ttu-id="81ca1-124">PowerPoint - 320 x 531 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-124">PowerPoint - 320x531 pixels</span></span>
- <span data-ttu-id="81ca1-125">Word - 320 x 531 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-125">Word - 320x531 pixels</span></span>
- <span data-ttu-id="81ca1-126">Outlook - 348 x 535 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-126">Outlook - 348x535 pixels</span></span>

<br/>

<span data-ttu-id="81ca1-127">*Figure 3. Office tailles de volet de tâches*</span><span class="sxs-lookup"><span data-stu-id="81ca1-127">*Figure 3. Office task pane sizes*</span></span>

![Diagramme affichant les tailles du volet Des tâches à une résolution de 1 366 x 768.](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="81ca1-129">Excel - 350 x 378 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-129">Excel - 350x378 pixels</span></span>
- <span data-ttu-id="81ca1-130">PowerPoint - 348 x 391 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-130">PowerPoint - 348x391 pixels</span></span>
- <span data-ttu-id="81ca1-131">Word - 329 x 445 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-131">Word - 329x445 pixels</span></span>
- <span data-ttu-id="81ca1-132">Outlook (sur le web) - 320 x 570 pixels</span><span class="sxs-lookup"><span data-stu-id="81ca1-132">Outlook (on the web) - 320x570 pixels</span></span>

## <a name="personality-menu"></a><span data-ttu-id="81ca1-133">Menu Caractéristique</span><span class="sxs-lookup"><span data-stu-id="81ca1-133">Personality menu</span></span>

<span data-ttu-id="81ca1-p103">Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="81ca1-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="81ca1-136">Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.</span><span class="sxs-lookup"><span data-stu-id="81ca1-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="81ca1-137">*Figure 4. Menu Caractéristique sur Windows*</span><span class="sxs-lookup"><span data-stu-id="81ca1-137">*Figure 4. Personality menu on Windows*</span></span>

![Diagramme montrant le menu De la personnalité sur Windows bureau.](../images/personality-menu-win.png)

<span data-ttu-id="81ca1-139">Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace à 34 x 32 pixels, comme illustré.</span><span class="sxs-lookup"><span data-stu-id="81ca1-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="81ca1-140">*figure 5. Menu Caractéristique sur Mac*</span><span class="sxs-lookup"><span data-stu-id="81ca1-140">*Figure 5. Personality menu on Mac*</span></span>

![Diagramme montrant le menu Personnalité sur le bureau Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="81ca1-142">Implémentation</span><span class="sxs-lookup"><span data-stu-id="81ca1-142">Implementation</span></span>

<span data-ttu-id="81ca1-143">Pour consulter un exemple qui implémente un volet des tâches, reportez-vous à [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) sur GitHub.</span><span class="sxs-lookup"><span data-stu-id="81ca1-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="81ca1-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="81ca1-144">See also</span></span>

- [<span data-ttu-id="81ca1-145">Cœur de fabric dans les modules</span><span class="sxs-lookup"><span data-stu-id="81ca1-145">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="81ca1-146">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="81ca1-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
