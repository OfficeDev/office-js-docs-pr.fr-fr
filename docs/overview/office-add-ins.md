---
title: Vue d?ensemble de la plateforme des compl?ments pour Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: f0f20371eee759a449773effaff1ce365e32bf48
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/25/2018
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="067fb-102">Vue d?ensemble de la plateforme de compl?ments pour Office</span><span class="sxs-lookup"><span data-stu-id="067fb-102">Office Add-ins platform overview</span></span>

<span data-ttu-id="067fb-p101">La plateforme des compl?ments Office permet de cr?er des solutions qui ?tendent des applications Office et interagissent avec du contenu dans des documents Office. Les compl?ments Office vous permettent d?utiliser des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour ?tendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes. Votre solution peut ?tre ex?cut?e dans Office sur plusieurs plateformes, notamment Office pour Windows, Office Online, Office pour Mac et Office pour iPad.</span><span class="sxs-lookup"><span data-stu-id="067fb-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>

<span data-ttu-id="067fb-p102">Les compl?ments Office offrent presque les m?mes possibilit?s qu?une page web dans un navigateur. Vous pouvez utiliser la plateforme des compl?ments Office pour :</span><span class="sxs-lookup"><span data-stu-id="067fb-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="067fb-p103">**Ajout de nouvelles fonctionnalit?s ? des clients Office :** vous pouvez importer des donn?es externes dans Office, automatiser des documents Office, exposer des fonctionnalit?s tierces dans des clients Office et bien plus encore. Par exemple, vous pouvez utiliser l?API Microsoft Graph pour ?tablir une connexion vers des donn?es qui am?liorent la productivit?.</span><span class="sxs-lookup"><span data-stu-id="067fb-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span> 
    
-  <span data-ttu-id="067fb-110">**Cr?er de nouveaux objets interactifs et enrichis qui peuvent ?tre incorpor?s dans des documents Office :** vous pouvez incorporer des cartes, des graphiques et des visualisations interactives que les utilisateurs peuvent ajouter ? leurs feuilles de calcul Excel et pr?sentations PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="067fb-110">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span> 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a><span data-ttu-id="067fb-111">En quoi les compl?ments Office sont-ils diff?rents des compl?ments COM et VSTO ?</span><span class="sxs-lookup"><span data-stu-id="067fb-111">How are Office Add-ins different than COM and VSTO add-ins?</span></span> 

<span data-ttu-id="067fb-p104">Les compl?ments COM ou VSTO sont des solutions d?int?gration ? Office ant?rieures qui s?ex?cutent uniquement sur Office pour Windows. Contrairement aux compl?ments COM, les compl?ments Office n?incluent pas de code ex?cut? sur l?appareil de l?utilisateur ou sur le client Office. Pour un compl?ment Office, l?application h?te, par exemple Excel, lit le manifeste du compl?ment et ins?re les commandes de menu et les boutons de ruban personnalis?s du compl?ment dans l?interface utilisateur. Lorsque cela est n?cessaire, elle charge le code JavaScript et HTML du compl?ment, qui est ex?cut? dans le contexte d?un navigateur dans un bac ? sable (sandbox).</span><span class="sxs-lookup"><span data-stu-id="067fb-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office for Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in?s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span> 

<span data-ttu-id="067fb-116">Les compl?ments Office offrent les avantages suivants par rapport aux compl?ments cr??s ? l?aide de VBA, COM ou VSTO :</span><span class="sxs-lookup"><span data-stu-id="067fb-116">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span> 

- <span data-ttu-id="067fb-p105">Prise en charge sur plusieurs plateformes. Les compl?ments Office s?ex?cutent dans Office pour Windows, Mac, iOS et Office Online.</span><span class="sxs-lookup"><span data-stu-id="067fb-p105">Cross-platform support. Office Add-ins run in Office for Windows, Mac, iOS, and Office Online.</span></span> 

- <span data-ttu-id="067fb-p106">Authentification unique (SSO) : les compl?ments Office s?int?grent facilement ? des comptes d?utilisateurs Office 365.</span><span class="sxs-lookup"><span data-stu-id="067fb-p106">Single sign-on (SSO). Office Add-ins integrate easily with users' Office 365 accounts.</span></span> 

- <span data-ttu-id="067fb-p107">D?ploiement et distribution centralis?s. Les administrateurs peuvent d?ployer des compl?ments Office de fa?on centralis?e dans une organisation.</span><span class="sxs-lookup"><span data-stu-id="067fb-p107">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span> 

- <span data-ttu-id="067fb-p108">Acc?s facile via AppSource. Vous pouvez mettre votre solution ? disposition d?un large public en l?envoyant ? AppSource.</span><span class="sxs-lookup"><span data-stu-id="067fb-p108">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span> 

- <span data-ttu-id="067fb-p109">S?appuie sur des technologies web standard. Vous pouvez utiliser n?importe quelle biblioth?que pour cr?er des compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="067fb-p109">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span> 

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="067fb-127">Composants d?un compl?ment Office</span><span class="sxs-lookup"><span data-stu-id="067fb-127">Components of an Office Add-in</span></span> 

<span data-ttu-id="067fb-p110">Un compl?ment Office inclut deux composants de base : un fichier manifeste XML et votre propre application web. Le manifeste d?finit diff?rents param?tres, y compris la fa?on dont votre compl?ment s?int?gre avec les clients Office. Votre application web doit ?tre h?berg?e sur un serveur web ou un service d?h?bergement web, tel que Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="067fb-p110">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

<span data-ttu-id="067fb-131">*Figure 1. Manifeste + page web = compl?ment Office*</span><span class="sxs-lookup"><span data-stu-id="067fb-131">*Figure 1. Manifest + webpage = an Office Add-in*</span></span>

![Manifeste + page web = compl?ment Office](../images/dk2-agave-overview-01.png)

### <a name="manifest"></a><span data-ttu-id="067fb-133">Manifeste</span><span class="sxs-lookup"><span data-stu-id="067fb-133">Manifest</span></span> 

<span data-ttu-id="067fb-134">Le manifeste est un fichier XML qui sp?cifie les param?tres et les fonctionnalit?s du compl?ment, notamment :</span><span class="sxs-lookup"><span data-stu-id="067fb-134">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span> 

- <span data-ttu-id="067fb-135">Le nom d?affichage, la description, l?ID, la version et les param?tres r?gionaux par d?faut du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="067fb-135">The add-in's display name, description, ID, version, and default locale.</span></span> 

- <span data-ttu-id="067fb-136">La fa?on dont le compl?ment s?int?gre ? Office.</span><span class="sxs-lookup"><span data-stu-id="067fb-136">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="067fb-137">Le niveau d?autorisation et les conditions d?acc?s aux donn?es pour le compl?ment.</span><span class="sxs-lookup"><span data-stu-id="067fb-137">The permission level and data access requirements for the add-in.</span></span> 

### <a name="web-app"></a><span data-ttu-id="067fb-138">Application web</span><span class="sxs-lookup"><span data-stu-id="067fb-138">Web app</span></span> 

<span data-ttu-id="067fb-p111">Le compl?ment Office le plus simple est compos? d?une page HTML statique qui est affich?e dans une application Office, mais qui n?interagit pas avec le document Office ou une autre ressource Internet. Toutefois, pour cr?er un compl?ment qui interagit avec des documents Office ou permet ? l?utilisateur d?interagir avec les ressources en ligne ? partir d?une application h?te Office, vous pouvez utiliser n?importe quelle technologie, aussi bien c?t? client que serveur, prise en charge par votre fournisseur d?h?bergement (par exemple, ASP.NET, PHP ou Node.js). Pour interagir avec des clients et des documents Office, vous pouvez utiliser les API JavaScript Office.js.</span><span class="sxs-lookup"><span data-stu-id="067fb-p111">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span> 

<span data-ttu-id="067fb-142">*Figure 2. Composants d?un compl?ment Office Hello World*</span><span class="sxs-lookup"><span data-stu-id="067fb-142">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Composants d?un compl?ment Hello World](../images/dk2-agave-overview-07.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="067fb-144">Extension des clients Office et interaction avec ces clients</span><span class="sxs-lookup"><span data-stu-id="067fb-144">Extending and interacting with Office clients</span></span> 

<span data-ttu-id="067fb-145">Les compl?ments Office offrent les possibilit?s suivantes dans une application Office h?te :</span><span class="sxs-lookup"><span data-stu-id="067fb-145">Office Add-ins can do the following within an Office host application:</span></span> 

-  <span data-ttu-id="067fb-146">?tendre les fonctionnalit?s (toutes les applications Office)</span><span class="sxs-lookup"><span data-stu-id="067fb-146">Extend functionality (any Office application)</span></span> 

-  <span data-ttu-id="067fb-147">Cr?er de nouveaux objets (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="067fb-147">Create new objects (Excel or PowerPoint)</span></span> 
 
### <a name="extend-office-functionality"></a><span data-ttu-id="067fb-148">?tendre les fonctionnalit?s d?Office</span><span class="sxs-lookup"><span data-stu-id="067fb-148">Extend Office functionality</span></span> 

<span data-ttu-id="067fb-149">Vous pouvez ajouter de nouvelles fonctionnalit?s aux applications Office via les ?l?ments d?interface suivants :</span><span class="sxs-lookup"><span data-stu-id="067fb-149">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="067fb-150">Commandes de menu et boutons de ruban personnalis?es (collectivement appel?s ? commandes de compl?ment ?)</span><span class="sxs-lookup"><span data-stu-id="067fb-150">Custom ribbon buttons and menu commands (collectively called ?add-in commands?)</span></span> 

-  <span data-ttu-id="067fb-151">Volets Office ? ins?rer</span><span class="sxs-lookup"><span data-stu-id="067fb-151">Insertable task panes</span></span> 

<span data-ttu-id="067fb-152">Les ?l?ments d?interface personnalis?s et les volets Office sont d?finis dans le manifeste du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="067fb-152">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="067fb-153">Commandes de menu et boutons personnalis?s</span><span class="sxs-lookup"><span data-stu-id="067fb-153">Custom buttons and menu commands</span></span>  

<span data-ttu-id="067fb-p112">Vous pouvez ajouter des ?l?ments de menu et des boutons de ruban personnalis? au ruban d?Office pour bureau Windows et Office Online. Les utilisateurs peuvent ainsi acc?der ? votre compl?ment directement ? partir de leur application Office. Les boutons de commande peuvent lancer diff?rentes actions, par exemple afficher un volet Office comportant du contenu HTML personnalis? ou ex?cuter une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="067fb-p112">You can add custom ribbon buttons and menu items to the ribbon in Office for Windows Desktop and Office Online. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="067fb-157">*Figure 3. Commandes de compl?ment en cours d?ex?cution dans Excel (version de bureau)*</span><span class="sxs-lookup"><span data-stu-id="067fb-157">*Figure 3. Add-in commands running in Excel Desktop*</span></span>

![Commandes de menu et boutons personnalis?s](../images/add-in-commands-overview.png)

#### <a name="task-panes"></a><span data-ttu-id="067fb-159">Volets Office</span><span class="sxs-lookup"><span data-stu-id="067fb-159">Task panes</span></span>  

<span data-ttu-id="067fb-p113">Vous pouvez utiliser des volets Office en plus des commandes de compl?ment pour permettre aux utilisateurs d?interagir avec votre solution. Les clients qui ne prennent pas en charge les commandes de compl?ment (Office 2013 et Office pour iPad) ex?cutent votre compl?ment sous la forme d?un volet Office. Les utilisateurs lancent les compl?ments de volet Office via le bouton **Mes compl?ments** situ? sous l?onglet **Insertion**.</span><span class="sxs-lookup"><span data-stu-id="067fb-p113">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office for iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span> 

<span data-ttu-id="067fb-163">*Figure 4. Volet Office*</span><span class="sxs-lookup"><span data-stu-id="067fb-163">*Figure 4. Task pane*</span></span>

![Volet de t?ches](../images/task-pane-overview.jpg)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="067fb-165">Extension des fonctionnalit?s Outlook</span><span class="sxs-lookup"><span data-stu-id="067fb-165">Extend Outlook functionality</span></span> 

<span data-ttu-id="067fb-p114">Les compl?ments Outlook peuvent d?velopper le ruban Office et s?afficher en regard d?un ?l?ment Outlook quand vous le visualisez ou le composez. Ils fonctionnent avec un message ?lectronique, une demande de r?union, une r?ponse ? une demande de r?union, une annulation de r?union ou un rendez-vous quand l?utilisateur visualise un ?l?ment re?u, r?pond ? un ?l?ment ou en cr?e un.</span><span class="sxs-lookup"><span data-stu-id="067fb-p114">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="067fb-p115">Les compl?ments Outlook peuvent acc?der ? des informations contextuelles ? partir de l??l?ment, telles qu?une adresse ou un ID de suivi, et utiliser ces donn?es pour acc?der ? d?autres informations sur le serveur ou provenant de services web pour cr?er des exp?riences utilisateur attrayantes. Dans la plupart des cas, un compl?ment Outlook peut ?tre ex?cut? sans modification sur les diff?rentes applications h?te prise en charge, notamment Outlook, Outlook pour Mac, Outlook Web App et Outlook Web App pour appareils, afin d?offrir une exp?rience homog?ne sur le bureau, en ligne, sur les tablettes et sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="067fb-p115">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and Outlook Web App for devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span> 

<span data-ttu-id="067fb-170">Pour acc?der ? une vue d?ensemble des compl?ments Outlook, reportez-vous ? la rubrique [Pr?sentation des compl?ments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="067fb-170">For an overview of Outlook add-ins, see [Outlook add-ins overview](https://docs.microsoft.com/en-us/outlook/add-ins/).</span></span> 

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="067fb-171">Cr?ation d?objets dans des documents Office</span><span class="sxs-lookup"><span data-stu-id="067fb-171">Create new objects in Office documents</span></span> 

<span data-ttu-id="067fb-p116">Vous pouvez incorporer des objets web, appel?s compl?ments de contenu, dans des documents Excel et PowerPoint. Ces compl?ments de contenu vous permettent d?int?grer des visualisations de donn?es web enrichies, du contenu multim?dia (comme un lecteur vid?o YouTube ou une galerie d?images) et d?autres types de contenu externe.</span><span class="sxs-lookup"><span data-stu-id="067fb-p116">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="067fb-174">*Figure 5. Compl?ment de contenu*</span><span class="sxs-lookup"><span data-stu-id="067fb-174">*Figure 5. Content add-in*</span></span>

![compl?ment de contenu](../images/dk2-agave-overview-05.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="067fb-176">API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="067fb-176">Office JavaScript APIs</span></span> 

<span data-ttu-id="067fb-p117">Les API JavaScript Office sont compos?es d?objets et de membres permettant de cr?er des compl?ments et d?interagir avec le contenu Office et les services web. Il existe un mod?le objet commun que se partagent Excel, Outlook, Word, PowerPoint, OneNote et Project. Il existe ?galement des mod?les objet plus complets et propres ? l?h?te pour Excel et Word. Ces API permettent d?acc?der ? des objets connus tels que des paragraphes et des classeurs, ce qui facilite la cr?ation de compl?ment pour un h?te sp?cifique.</span><span class="sxs-lookup"><span data-stu-id="067fb-p117">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="067fb-181">?tapes suivantes</span><span class="sxs-lookup"><span data-stu-id="067fb-181">Next steps</span></span> 

<span data-ttu-id="067fb-182">Pour en savoir plus sur la cr?ation de votre compl?ment Office, essayez notre [D?marrage rapide en 5 minutes](https://docs.microsoft.com/en-us/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="067fb-182">To learn more about how to start building your Office Add-in, try out our [5-minute Quickstarts](https://docs.microsoft.com/en-us/office/dev/add-ins/). You can start building add-ins right away using Visual Studio or any other editor.</span></span> <span data-ttu-id="067fb-183">Vous pouvez commencer ? cr?er des compl?ments imm?diatement ? l'aide de Visual Studio ou de tout autre ?diteur.</span><span class="sxs-lookup"><span data-stu-id="067fb-183">To learn more about how to start building your Office Add-in, try out our 5-minute Quickstarts. You can start building add-ins right away using Visual Studio or any other editor.</span></span> 

<span data-ttu-id="067fb-184">Pour commencer ? concevoir des solutions offrant des exp?riences utilisateur efficaces et attrayantes, consultez les [instructions de conception](../design/add-in-design.md) et les [meilleures pratiques](../concepts/add-in-development-best-practices.md) pour les compl?ments Office.</span><span class="sxs-lookup"><span data-stu-id="067fb-184">To start planning solutions that create effective and compelling user experiences, get familiar with the [design guidelines](../design/add-in-design.md) and [best practices](../concepts/add-in-development-best-practices.md) for Office Add-ins.</span></span>    
   
## <a name="see-also"></a><span data-ttu-id="067fb-185">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="067fb-185">See also</span></span>

- [<span data-ttu-id="067fb-186">Exemples de compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="067fb-186">Office Add-in samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples)
- [<span data-ttu-id="067fb-187">Pr?sentation de l?API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="067fb-187">Understanding the JavaScript API for Office</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="067fb-188">Disponibilit? des compl?ments Office sur les plateformes et les h?tes</span><span class="sxs-lookup"><span data-stu-id="067fb-188">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)


    
