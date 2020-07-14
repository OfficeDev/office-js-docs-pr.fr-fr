---
title: Vue d’ensemble de la plateforme de compléments pour Office | Microsoft Docs
description: Utilisez des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour étendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes.
ms.date: 02/13/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6b162a166bda0c988f5fbbaade3b0bef4b650984
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094070"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="04b3d-103">Vue d’ensemble de la plateforme de compléments pour Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="04b3d-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span><span class="sxs-lookup"><span data-stu-id="04b3d-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span></span> <span data-ttu-id="04b3d-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span><span class="sxs-lookup"><span data-stu-id="04b3d-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span></span> <span data-ttu-id="04b3d-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span><span class="sxs-lookup"><span data-stu-id="04b3d-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![Image d'une extensibilité de complément Office](../images/addins-overview.png)

<span data-ttu-id="04b3d-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span><span class="sxs-lookup"><span data-stu-id="04b3d-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span></span> <span data-ttu-id="04b3d-109">Use the Office Add-ins platform to:</span><span class="sxs-lookup"><span data-stu-id="04b3d-109">Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="04b3d-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span><span class="sxs-lookup"><span data-stu-id="04b3d-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span></span> <span data-ttu-id="04b3d-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span><span class="sxs-lookup"><span data-stu-id="04b3d-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="04b3d-112">**Créer de nouveaux objets interactifs et enrichis qui peuvent être incorporés dans des documents Office :** vous pouvez incorporer des cartes, des graphiques et des visualisations interactives que les utilisateurs peuvent ajouter à leurs feuilles de calcul Excel et présentations PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="04b3d-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="04b3d-113">En quoi les compléments Office sont-ils différents des compléments COM et VSTO ?</span><span class="sxs-lookup"><span data-stu-id="04b3d-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="04b3d-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span><span class="sxs-lookup"><span data-stu-id="04b3d-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span></span> <span data-ttu-id="04b3d-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span><span class="sxs-lookup"><span data-stu-id="04b3d-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span></span> <span data-ttu-id="04b3d-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span><span class="sxs-lookup"><span data-stu-id="04b3d-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span></span> <span data-ttu-id="04b3d-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span><span class="sxs-lookup"><span data-stu-id="04b3d-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![Image expliquant les raisons d'utiliser les compléments Office](../images/why.png)

<span data-ttu-id="04b3d-119">Les compléments Office offrent les avantages suivants par rapport aux compléments créés à l’aide de VBA, COM ou VSTO :</span><span class="sxs-lookup"><span data-stu-id="04b3d-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="04b3d-120">Cross-platform support.</span><span class="sxs-lookup"><span data-stu-id="04b3d-120">Cross-platform support.</span></span> <span data-ttu-id="04b3d-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span><span class="sxs-lookup"><span data-stu-id="04b3d-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="04b3d-122">Centralized deployment and distribution.</span><span class="sxs-lookup"><span data-stu-id="04b3d-122">Centralized deployment and distribution.</span></span> <span data-ttu-id="04b3d-123">Admins can deploy Office Add-ins centrally across an organization.</span><span class="sxs-lookup"><span data-stu-id="04b3d-123">Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="04b3d-124">Easy access via AppSource.</span><span class="sxs-lookup"><span data-stu-id="04b3d-124">Easy access via AppSource.</span></span> <span data-ttu-id="04b3d-125">You can make your solution available to a broad audience by submitting it to AppSource.</span><span class="sxs-lookup"><span data-stu-id="04b3d-125">You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="04b3d-126">Based on standard web technology.</span><span class="sxs-lookup"><span data-stu-id="04b3d-126">Based on standard web technology.</span></span> <span data-ttu-id="04b3d-127">You can use any library you like to build Office Add-ins.</span><span class="sxs-lookup"><span data-stu-id="04b3d-127">You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="04b3d-128">Composants d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-128">Components of an Office Add-in</span></span>

<span data-ttu-id="04b3d-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span><span class="sxs-lookup"><span data-stu-id="04b3d-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span></span> <span data-ttu-id="04b3d-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span><span class="sxs-lookup"><span data-stu-id="04b3d-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span></span> <span data-ttu-id="04b3d-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="04b3d-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="04b3d-132">Manifeste</span><span class="sxs-lookup"><span data-stu-id="04b3d-132">Manifest</span></span>

<span data-ttu-id="04b3d-133">Le manifeste est un fichier XML qui spécifie les paramètres et les fonctionnalités du complément, notamment :</span><span class="sxs-lookup"><span data-stu-id="04b3d-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="04b3d-134">Le nom d’affichage, la description, l’ID, la version et les paramètres régionaux par défaut du complément.</span><span class="sxs-lookup"><span data-stu-id="04b3d-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="04b3d-135">La façon dont le complément s’intègre à Office.</span><span class="sxs-lookup"><span data-stu-id="04b3d-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="04b3d-136">Le niveau d’autorisation et les conditions d’accès aux données pour le complément.</span><span class="sxs-lookup"><span data-stu-id="04b3d-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="04b3d-137">Application web</span><span class="sxs-lookup"><span data-stu-id="04b3d-137">Web app</span></span>

<span data-ttu-id="04b3d-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span><span class="sxs-lookup"><span data-stu-id="04b3d-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span></span> <span data-ttu-id="04b3d-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span><span class="sxs-lookup"><span data-stu-id="04b3d-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span></span> <span data-ttu-id="04b3d-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span><span class="sxs-lookup"><span data-stu-id="04b3d-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="04b3d-141">*Figure 2. Composants d’un complément Office Hello World*</span><span class="sxs-lookup"><span data-stu-id="04b3d-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Composants d’un complément Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="04b3d-143">Extension des clients Office et interaction avec ces clients</span><span class="sxs-lookup"><span data-stu-id="04b3d-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="04b3d-144">Les compléments Office offrent les possibilités suivantes dans une application Office hôte :</span><span class="sxs-lookup"><span data-stu-id="04b3d-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="04b3d-145">Étendre les fonctionnalités (toutes les applications Office)</span><span class="sxs-lookup"><span data-stu-id="04b3d-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="04b3d-146">Créer de nouveaux objets (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="04b3d-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="04b3d-147">Étendre les fonctionnalités d’Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-147">Extend Office functionality</span></span>

<span data-ttu-id="04b3d-148">Vous pouvez ajouter de nouvelles fonctionnalités aux applications Office via les éléments d’interface suivants :</span><span class="sxs-lookup"><span data-stu-id="04b3d-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="04b3d-149">Commandes de menu et boutons de ruban personnalisées (collectivement appelés « commandes de complément »)</span><span class="sxs-lookup"><span data-stu-id="04b3d-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="04b3d-150">Volets Office à insérer</span><span class="sxs-lookup"><span data-stu-id="04b3d-150">Insertable task panes</span></span>

<span data-ttu-id="04b3d-151">Les éléments d’interface personnalisés et les volets Office sont définis dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="04b3d-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="04b3d-152">Commandes de menu et boutons personnalisés</span><span class="sxs-lookup"><span data-stu-id="04b3d-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="04b3d-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span><span class="sxs-lookup"><span data-stu-id="04b3d-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span></span> <span data-ttu-id="04b3d-154">This makes it easy for users to access your add-in directly from their Office application.</span><span class="sxs-lookup"><span data-stu-id="04b3d-154">This makes it easy for users to access your add-in directly from their Office application.</span></span> <span data-ttu-id="04b3d-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span><span class="sxs-lookup"><span data-stu-id="04b3d-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="04b3d-156">*Figure 3. Commandes des compléments dans le ruban*</span><span class="sxs-lookup"><span data-stu-id="04b3d-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![Commandes de menu et boutons personnalisés](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="04b3d-158">Volets Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-158">Task panes</span></span>  

<span data-ttu-id="04b3d-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span><span class="sxs-lookup"><span data-stu-id="04b3d-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span></span> <span data-ttu-id="04b3d-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span><span class="sxs-lookup"><span data-stu-id="04b3d-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span></span> <span data-ttu-id="04b3d-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span><span class="sxs-lookup"><span data-stu-id="04b3d-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="04b3d-162">*Figure 4. Volet Office*</span><span class="sxs-lookup"><span data-stu-id="04b3d-162">*Figure 4. Task pane*</span></span>

![Utiliser les volets Office en plus des commandes de complément](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="04b3d-164">Extension des fonctionnalités Outlook</span><span class="sxs-lookup"><span data-stu-id="04b3d-164">Extend Outlook functionality</span></span>

<span data-ttu-id="04b3d-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span><span class="sxs-lookup"><span data-stu-id="04b3d-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span></span> <span data-ttu-id="04b3d-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span><span class="sxs-lookup"><span data-stu-id="04b3d-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="04b3d-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span><span class="sxs-lookup"><span data-stu-id="04b3d-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span></span> <span data-ttu-id="04b3d-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span><span class="sxs-lookup"><span data-stu-id="04b3d-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="04b3d-169">Pour accéder à une vue d’ensemble des compléments Outlook, reportez-vous à la rubrique [Présentation des compléments Outlook](../outlook/outlook-add-ins-overview.md).</span><span class="sxs-lookup"><span data-stu-id="04b3d-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="04b3d-170">Création d’objets dans des documents Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-170">Create new objects in Office documents</span></span>

<span data-ttu-id="04b3d-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span><span class="sxs-lookup"><span data-stu-id="04b3d-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span></span> <span data-ttu-id="04b3d-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span><span class="sxs-lookup"><span data-stu-id="04b3d-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="04b3d-173">*Figure 5. Complément de contenu*</span><span class="sxs-lookup"><span data-stu-id="04b3d-173">*Figure 5. Content add-in*</span></span>

![Incorporer des objets web appelés compléments de contenu](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="04b3d-175">API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-175">Office JavaScript APIs</span></span>

<span data-ttu-id="04b3d-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span><span class="sxs-lookup"><span data-stu-id="04b3d-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span></span> <span data-ttu-id="04b3d-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span><span class="sxs-lookup"><span data-stu-id="04b3d-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span></span> <span data-ttu-id="04b3d-178">There are also more extensive host-specific object models for Excel and Word.</span><span class="sxs-lookup"><span data-stu-id="04b3d-178">There are also more extensive host-specific object models for Excel and Word.</span></span> <span data-ttu-id="04b3d-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span><span class="sxs-lookup"><span data-stu-id="04b3d-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="04b3d-180">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="04b3d-180">Next steps</span></span>

<span data-ttu-id="04b3d-181">Pour une présentation en détails du développement des compléments Office, voir [Création de compléments Office](../overview/office-add-ins-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="04b3d-181">For a more detailed introduction to developing Office Add-ins, see [Building Office Add-ins](../overview/office-add-ins-fundamentals.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="04b3d-182">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="04b3d-182">See also</span></span>

- [<span data-ttu-id="04b3d-183">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-183">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="04b3d-184">Concepts de base pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-184">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="04b3d-185">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-185">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="04b3d-186">Concevoir des compléments Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-186">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="04b3d-187">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="04b3d-187">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="04b3d-188">Publish Office Add-ins</span><span class="sxs-lookup"><span data-stu-id="04b3d-188">Publish Office Add-ins</span></span>](../publish/publish.md)
