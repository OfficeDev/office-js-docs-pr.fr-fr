---
title: Vue d’ensemble de la plateforme de compléments pour Office | Microsoft Docs
description: Utilisez des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour étendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes.
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 956e19a14cca1559c828265b2212c410f10b916b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076657"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="253ac-103">Vue d’ensemble de la plateforme de compléments pour Office</span><span class="sxs-lookup"><span data-stu-id="253ac-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="253ac-p101">La plateforme des compléments Office permet de créer des solutions qui étendent des applications Office et interagissent avec du contenu dans des documents Office. Les compléments Office vous permettent d’utiliser des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour étendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes. Votre solution peut être exécutée dans Office sur plusieurs plateformes, notamment Windows, Mac et iPad, ainsi que dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="253ac-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![L’application Office et un site web incorporé (complément) offrent des possibilités d’extensibilité infinies.](../images/addins-overview.png)

<span data-ttu-id="253ac-p102">Les compléments Office offrent presque les mêmes possibilités qu’une page web dans un navigateur. Vous pouvez utiliser la plateforme des compléments Office pour :</span><span class="sxs-lookup"><span data-stu-id="253ac-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="253ac-p103">**Ajout de nouvelles fonctionnalités à des clients Office :** vous pouvez importer des données externes dans Office, automatiser des documents Office, exposer des fonctionnalités tierces dans des clients Office et bien plus encore. Par exemple, vous pouvez utiliser l’API Microsoft Graph pour établir une connexion vers des données qui améliorent la productivité.</span><span class="sxs-lookup"><span data-stu-id="253ac-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="253ac-112">**Créer de nouveaux objets interactifs et enrichis qui peuvent être incorporés dans des documents Office :** vous pouvez incorporer des cartes, des graphiques et des visualisations interactives que les utilisateurs peuvent ajouter à leurs feuilles de calcul Excel et présentations PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="253ac-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="253ac-113">En quoi les compléments Office sont-ils différents des compléments COM et VSTO ?</span><span class="sxs-lookup"><span data-stu-id="253ac-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="253ac-p104">Les compléments COM ou VSTO sont des solutions d’intégration à Office antérieures qui s’exécutent uniquement dans Office pour Windows. Contrairement aux compléments COM, les compléments Office n’incluent pas de code exécuté sur l’appareil de l’utilisateur ni sur le client Office. Pour un complément Office, l’application (par exemple, Excel), lit le manifeste du complément et insère les commandes de menu et les boutons de ruban personnalisés du complément dans l’interface utilisateur. Lorsque cela est nécessaire, elle charge le code JavaScript et HTML du complément, qui est exécuté dans le contexte d’un navigateur dans un bac à sable (sandbox).</span><span class="sxs-lookup"><span data-stu-id="253ac-p104">COM or VSTO add-ins are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the application (for example, Excel), reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![Les raisons d’utiliser les compléments Office : multiplateforme, déploiement centralisé, accès facile via AppSource et basées sur des technologies web standard.](../images/why.png)

<span data-ttu-id="253ac-119">Les compléments Office offrent les avantages suivants par rapport aux compléments créés à l’aide de VBA, COM ou VSTO :</span><span class="sxs-lookup"><span data-stu-id="253ac-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="253ac-p105">Prise en charge sur plusieurs plateformes. Les compléments Office s’exécutent sur Office sur le web, Windows, Mac et iPad.</span><span class="sxs-lookup"><span data-stu-id="253ac-p105">Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="253ac-p106">Déploiement et distribution centralisés. Les administrateurs peuvent déployer des compléments Office de façon centralisée dans une organisation.</span><span class="sxs-lookup"><span data-stu-id="253ac-p106">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="253ac-p107">Accès facile via AppSource. Vous pouvez mettre votre solution à disposition d’un large public en l’envoyant à AppSource.</span><span class="sxs-lookup"><span data-stu-id="253ac-p107">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="253ac-p108">S’appuie sur des technologies web standard. Vous pouvez utiliser n’importe quelle bibliothèque pour créer des compléments Office.</span><span class="sxs-lookup"><span data-stu-id="253ac-p108">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="253ac-128">Composants d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="253ac-128">Components of an Office Add-in</span></span>

<span data-ttu-id="253ac-p109">Un complément Office inclut deux composants de base : un fichier manifeste XML et votre propre application web. Le manifeste définit différents paramètres, y compris la façon dont votre complément s’intègre avec les clients Office. Votre application web doit être hébergée sur un serveur web ou un service d’hébergement web, tel que Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="253ac-p109">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="253ac-132">Manifeste</span><span class="sxs-lookup"><span data-stu-id="253ac-132">Manifest</span></span>

<span data-ttu-id="253ac-133">Le manifeste est un fichier XML qui spécifie les paramètres et les fonctionnalités du complément, notamment :</span><span class="sxs-lookup"><span data-stu-id="253ac-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="253ac-134">Le nom d’affichage, la description, l’ID, la version et les paramètres régionaux par défaut du complément.</span><span class="sxs-lookup"><span data-stu-id="253ac-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="253ac-135">La façon dont le complément s’intègre à Office.</span><span class="sxs-lookup"><span data-stu-id="253ac-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="253ac-136">Le niveau d’autorisation et les conditions d’accès aux données pour le complément.</span><span class="sxs-lookup"><span data-stu-id="253ac-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="253ac-137">Application web</span><span class="sxs-lookup"><span data-stu-id="253ac-137">Web app</span></span>

<span data-ttu-id="253ac-p110">Le complément Office le plus simple est composé d’une page HTML statique qui est affichée dans une application Office, mais qui n’interagit pas avec le document Office ou une autre ressource Internet. Toutefois, pour créer un complément qui interagit avec des documents Office ou permet à l’utilisateur d’interagir avec les ressources en ligne à partir d’une application cliente Office, vous pouvez utiliser n’importe quelle technologie, aussi bien côté client que serveur, prise en charge par votre fournisseur d’hébergement (par exemple, ASP.NET, PHP ou Node.js). Pour interagir avec des clients et des documents Office, vous pouvez utiliser les API JavaScript Office.js.</span><span class="sxs-lookup"><span data-stu-id="253ac-p110">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office client application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="253ac-141">*Figure 2. Composants d’un complément Office Hello World*</span><span class="sxs-lookup"><span data-stu-id="253ac-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Composants d’un complément Hello World.](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="253ac-143">Extension des clients Office et interaction avec ces clients</span><span class="sxs-lookup"><span data-stu-id="253ac-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="253ac-144">Les compléments Office offrent les possibilités suivantes dans une application cliente Office :</span><span class="sxs-lookup"><span data-stu-id="253ac-144">Office Add-ins can do the following within an Office client application:</span></span>

-  <span data-ttu-id="253ac-145">Étendre les fonctionnalités (toutes les applications Office)</span><span class="sxs-lookup"><span data-stu-id="253ac-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="253ac-146">Créer de nouveaux objets (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="253ac-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="253ac-147">Étendre les fonctionnalités d’Office</span><span class="sxs-lookup"><span data-stu-id="253ac-147">Extend Office functionality</span></span>

<span data-ttu-id="253ac-148">Vous pouvez ajouter de nouvelles fonctionnalités aux applications Office via les éléments d’interface suivants :</span><span class="sxs-lookup"><span data-stu-id="253ac-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="253ac-149">Commandes de menu et boutons de ruban personnalisées (collectivement appelés « commandes de complément »)</span><span class="sxs-lookup"><span data-stu-id="253ac-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="253ac-150">Volets Office à insérer</span><span class="sxs-lookup"><span data-stu-id="253ac-150">Insertable task panes</span></span>

<span data-ttu-id="253ac-151">Les éléments d’interface personnalisés et les volets Office sont définis dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="253ac-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="253ac-152">Commandes de menu et boutons personnalisés</span><span class="sxs-lookup"><span data-stu-id="253ac-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="253ac-p111">Vous pouvez ajouter des éléments de menu et des boutons de ruban personnalisé au ruban d’Office sur le web et sur Windows. Les utilisateurs peuvent ainsi accéder à votre complément directement à partir de leur application Office. Les boutons de commande peuvent lancer différentes actions, par exemple afficher un volet Office comportant du contenu HTML personnalisé ou exécuter une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="253ac-p111">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and on Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="253ac-156">*Figure 3. Commandes des compléments dans le ruban*</span><span class="sxs-lookup"><span data-stu-id="253ac-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![Boutons personnalisés et commandes de menu.](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="253ac-158">Volets Office</span><span class="sxs-lookup"><span data-stu-id="253ac-158">Task panes</span></span>  

<span data-ttu-id="253ac-p112">Vous pouvez utiliser des volets Office en plus des commandes de complément pour permettre aux utilisateurs d’interagir avec votre solution. Les clients qui ne prennent pas en charge les commandes de complément (Office 2013 et Office sur iPad) exécutent votre complément sous la forme d’un volet Office. Les utilisateurs lancent les compléments de volet Office via le bouton **Mes compléments** situé dans l’onglet **Insertion**.</span><span class="sxs-lookup"><span data-stu-id="253ac-p112">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="253ac-162">*Figure 4. Volet Office*</span><span class="sxs-lookup"><span data-stu-id="253ac-162">*Figure 4. Task pane*</span></span>

![Utilisez des volets de tâches en plus des commandes de complément.](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="253ac-164">Extension des fonctionnalités Outlook</span><span class="sxs-lookup"><span data-stu-id="253ac-164">Extend Outlook functionality</span></span>

<span data-ttu-id="253ac-p113">Les add-ins Outlook peuvent étendre le ruban des applications Office et s'afficher contextuellement à côté d'un élément Outlook lorsque vous le visualisez ou le composez. Ils peuvent fonctionner avec un message électronique, une demande de réunion, une réponse à une réunion, l'annulation d'une réunion ou un rendez-vous lorsqu'un utilisateur consulte un élément reçu ou répond ou crée un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="253ac-p113">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="253ac-p114">Les compléments Outlook peuvent accéder aux informations contextuelles de l’élément, comme l’adresse ou l’ID de suivi, puis utiliser ces données pour accéder à des informations complémentaires sur le serveur et à partir des services web de façon à enrichir l’expérience utilisateur. Dans la plupart des cas, un complément Outlook s’exécute sans modification dans l’application Outlook afin d’offrir aux utilisateurs une expérience transparente sur le bureau, le web, les tablettes et les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="253ac-p114">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="253ac-169">Pour accéder à une vue d’ensemble des compléments Outlook, reportez-vous à la rubrique [Présentation des compléments Outlook](../outlook/outlook-add-ins-overview.md).</span><span class="sxs-lookup"><span data-stu-id="253ac-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="253ac-170">Création d’objets dans des documents Office</span><span class="sxs-lookup"><span data-stu-id="253ac-170">Create new objects in Office documents</span></span>

<span data-ttu-id="253ac-p115">Vous pouvez incorporer des objets web, appelés compléments de contenu, dans des documents Excel et PowerPoint. Ces compléments de contenu vous permettent d’intégrer des visualisations de données web enrichies, du contenu multimédia (comme un lecteur vidéo YouTube ou une galerie d’images) et d’autres types de contenu externe.</span><span class="sxs-lookup"><span data-stu-id="253ac-p115">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="253ac-173">*Figure 5. Complément de contenu*</span><span class="sxs-lookup"><span data-stu-id="253ac-173">*Figure 5. Content add-in*</span></span>

![Incorporer des objets web appelés compléments de contenu.](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="253ac-175">API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="253ac-175">Office JavaScript APIs</span></span>

<span data-ttu-id="253ac-p116">Les API JavaScript Office sont composées d’objets et de membres permettant de créer des compléments et d’interagir avec le contenu Office et les services web. Il existe un modèle objet commun que se partagent Excel, Outlook, Word, PowerPoint, OneNote et Project. Il existe également des modèles objet plus complets et propres à l’application pour Excel et Word. Ces API permettent d’accéder à des objets connus tels que des paragraphes et des classeurs, ce qui facilite la création de complément pour une application spécifique.</span><span class="sxs-lookup"><span data-stu-id="253ac-p116">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive application-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific application.</span></span>

## <a name="next-steps"></a><span data-ttu-id="253ac-180">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="253ac-180">Next steps</span></span>

<span data-ttu-id="253ac-181">Pour une présentation en détails du développement des compléments Office, voir [Développement de compléments Office](../develop/develop-overview.md).</span><span class="sxs-lookup"><span data-stu-id="253ac-181">For a more detailed introduction to developing Office Add-ins, see [Develop Office Add-ins](../develop/develop-overview.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="253ac-182">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="253ac-182">See also</span></span>

- [<span data-ttu-id="253ac-183">Concepts de base pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="253ac-183">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="253ac-184">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="253ac-184">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="253ac-185">Concevoir des compléments Office</span><span class="sxs-lookup"><span data-stu-id="253ac-185">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="253ac-186">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="253ac-186">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="253ac-187">Publier des compléments Office</span><span class="sxs-lookup"><span data-stu-id="253ac-187">Publish Office Add-ins</span></span>](../publish/publish.md)
- [<span data-ttu-id="253ac-188">Découvrez le programme pour les développeurs Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="253ac-188">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
