---
title: Vue d’ensemble de la plateforme de compléments pour Office | Microsoft Docs
description: Utilisez des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour étendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 480228c20b20de52a9e1224f6691696b5560986c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870316"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="550f6-103">Vue d’ensemble de la plateforme de compléments pour Office</span><span class="sxs-lookup"><span data-stu-id="550f6-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="550f6-p101">La plateforme des compléments Office permet de créer des solutions qui étendent des applications Office et interagissent avec du contenu dans des documents Office. Les compléments Office vous permettent d’utiliser des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour étendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes. Votre solution peut être exécutée dans Office sur plusieurs plateformes, notamment Office pour Windows, Office Online, Office pour Mac et Office pour iPad.</span><span class="sxs-lookup"><span data-stu-id="550f6-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>

<span data-ttu-id="550f6-p102">Les compléments Office offrent presque les mêmes possibilités qu’une page web dans un navigateur. Vous pouvez utiliser la plateforme des compléments Office pour :</span><span class="sxs-lookup"><span data-stu-id="550f6-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="550f6-p103">**Ajout de nouvelles fonctionnalités à des clients Office :** vous pouvez importer des données externes dans Office, automatiser des documents Office, exposer des fonctionnalités tierces dans des clients Office et bien plus encore. Par exemple, vous pouvez utiliser l’API Microsoft Graph pour établir une connexion vers des données qui améliorent la productivité.</span><span class="sxs-lookup"><span data-stu-id="550f6-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="550f6-111">**Créer de nouveaux objets interactifs et enrichis qui peuvent être incorporés dans des documents Office :** vous pouvez incorporer des cartes, des graphiques et des visualisations interactives que les utilisateurs peuvent ajouter à leurs feuilles de calcul Excel et présentations PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="550f6-111">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="550f6-112">En quoi les compléments Office sont-ils différents des compléments COM et VSTO ?</span><span class="sxs-lookup"><span data-stu-id="550f6-112">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="550f6-p104">Les compléments COM ou VSTO sont des solutions d’intégration à Office antérieures qui s’exécutent uniquement sur Office pour Windows. Contrairement aux compléments COM, les compléments Office n’incluent pas de code exécuté sur l’appareil de l’utilisateur ou sur le client Office. Pour un complément Office, l’application hôte, par exemple Excel, lit le manifeste du complément et insère les commandes de menu et les boutons de ruban personnalisés du complément dans l’interface utilisateur. Lorsque cela est nécessaire, elle charge le code JavaScript et HTML du complément, qui est exécuté dans le contexte d’un navigateur dans un bac à sable (sandbox).</span><span class="sxs-lookup"><span data-stu-id="550f6-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office for Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="550f6-117">Les compléments Office offrent les avantages suivants par rapport aux compléments créés à l’aide de VBA, COM ou VSTO :</span><span class="sxs-lookup"><span data-stu-id="550f6-117">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="550f6-p105">Prise en charge sur plusieurs plateformes. Les compléments Office s’exécutent dans Office pour Windows, Mac, iOS et Office Online.</span><span class="sxs-lookup"><span data-stu-id="550f6-p105">Cross-platform support. Office Add-ins run in Office for Windows, Mac, iOS, and Office Online.</span></span>

- <span data-ttu-id="550f6-p106">Déploiement et distribution centralisés. Les administrateurs peuvent déployer des compléments Office de façon centralisée dans une organisation.</span><span class="sxs-lookup"><span data-stu-id="550f6-p106">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="550f6-p107">Accès facile via AppSource. Vous pouvez mettre votre solution à disposition d’un large public en l’envoyant à AppSource.</span><span class="sxs-lookup"><span data-stu-id="550f6-p107">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="550f6-p108">S’appuie sur des technologies web standard. Vous pouvez utiliser n’importe quelle bibliothèque pour créer des compléments Office.</span><span class="sxs-lookup"><span data-stu-id="550f6-p108">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="550f6-126">Composants d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="550f6-126">Components of an Office Add-in</span></span>

<span data-ttu-id="550f6-p109">Un complément Office inclut deux composants de base : un fichier manifeste XML et votre propre application web. Le manifeste définit différents paramètres, y compris la façon dont votre complément s’intègre avec les clients Office. Votre application web doit être hébergée sur un serveur web ou un service d’hébergement web, tel que Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="550f6-p109">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

<span data-ttu-id="550f6-130">*Figure 1. Manifeste de complément (XML) + page web (HTML, JS) = un complément Office*</span><span class="sxs-lookup"><span data-stu-id="550f6-130">*Figure 1. Add-in manifest (XML) + webpage (HTML, JS) = an Office Add-in*</span></span>

![Manifeste + page web = complément Office](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a><span data-ttu-id="550f6-132">Manifeste</span><span class="sxs-lookup"><span data-stu-id="550f6-132">Manifest</span></span>

<span data-ttu-id="550f6-133">Le manifeste est un fichier XML qui spécifie les paramètres et les fonctionnalités du complément, notamment :</span><span class="sxs-lookup"><span data-stu-id="550f6-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="550f6-134">Le nom d’affichage, la description, l’ID, la version et les paramètres régionaux par défaut du complément.</span><span class="sxs-lookup"><span data-stu-id="550f6-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="550f6-135">La façon dont le complément s’intègre à Office.</span><span class="sxs-lookup"><span data-stu-id="550f6-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="550f6-136">Le niveau d’autorisation et les conditions d’accès aux données pour le complément.</span><span class="sxs-lookup"><span data-stu-id="550f6-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="550f6-137">Application web</span><span class="sxs-lookup"><span data-stu-id="550f6-137">Web app</span></span>

<span data-ttu-id="550f6-p110">Le complément Office le plus simple est composé d’une page HTML statique qui est affichée dans une application Office, mais qui n’interagit pas avec le document Office ou une autre ressource Internet. Toutefois, pour créer un complément qui interagit avec des documents Office ou permet à l’utilisateur d’interagir avec les ressources en ligne à partir d’une application hôte Office, vous pouvez utiliser n’importe quelle technologie, aussi bien côté client que serveur, prise en charge par votre fournisseur d’hébergement (par exemple, ASP.NET, PHP ou Node.js). Pour interagir avec des clients et des documents Office, vous pouvez utiliser les API JavaScript Office.js.</span><span class="sxs-lookup"><span data-stu-id="550f6-p110">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="550f6-141">*Figure 2. Composants d’un complément Office Hello World*</span><span class="sxs-lookup"><span data-stu-id="550f6-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Composants d’un complément Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="550f6-143">Extension des clients Office et interaction avec ces clients</span><span class="sxs-lookup"><span data-stu-id="550f6-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="550f6-144">Les compléments Office offrent les possibilités suivantes dans une application Office hôte :</span><span class="sxs-lookup"><span data-stu-id="550f6-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="550f6-145">Étendre les fonctionnalités (toutes les applications Office)</span><span class="sxs-lookup"><span data-stu-id="550f6-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="550f6-146">Créer de nouveaux objets (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="550f6-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="550f6-147">Étendre les fonctionnalités d’Office</span><span class="sxs-lookup"><span data-stu-id="550f6-147">Extend Office functionality</span></span>

<span data-ttu-id="550f6-148">Vous pouvez ajouter de nouvelles fonctionnalités aux applications Office via les éléments d’interface suivants :</span><span class="sxs-lookup"><span data-stu-id="550f6-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="550f6-149">Commandes de menu et boutons de ruban personnalisées (collectivement appelés « commandes de complément »)</span><span class="sxs-lookup"><span data-stu-id="550f6-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="550f6-150">Volets Office à insérer</span><span class="sxs-lookup"><span data-stu-id="550f6-150">Insertable task panes</span></span>

<span data-ttu-id="550f6-151">Les éléments d’interface personnalisés et les volets Office sont définis dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="550f6-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="550f6-152">Commandes de menu et boutons personnalisés</span><span class="sxs-lookup"><span data-stu-id="550f6-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="550f6-p111">Vous pouvez ajouter des éléments de menu et des boutons de ruban personnalisé au ruban d’Office pour bureau Windows et Office Online. Les utilisateurs peuvent ainsi accéder à votre complément directement à partir de leur application Office. Les boutons de commande peuvent lancer différentes actions, par exemple afficher un volet Office comportant du contenu HTML personnalisé ou exécuter une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="550f6-p111">You can add custom ribbon buttons and menu items to the ribbon in Office for Windows Desktop and Office Online. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="550f6-156">*Figure 3. Commandes des compléments dans le ruban*</span><span class="sxs-lookup"><span data-stu-id="550f6-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![Commandes de menu et boutons personnalisés](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="550f6-158">Volets Office</span><span class="sxs-lookup"><span data-stu-id="550f6-158">Task panes</span></span>  

<span data-ttu-id="550f6-p112">Vous pouvez utiliser des volets Office en plus des commandes de complément pour permettre aux utilisateurs d’interagir avec votre solution. Les clients qui ne prennent pas en charge les commandes de complément (Office 2013 et Office pour iPad) exécutent votre complément sous la forme d’un volet Office. Les utilisateurs lancent les compléments de volet Office via le bouton **Mes compléments** situé sous l’onglet **Insertion**.</span><span class="sxs-lookup"><span data-stu-id="550f6-p112">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office for iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span> 

<span data-ttu-id="550f6-162">*Figure 4. Volet Office*</span><span class="sxs-lookup"><span data-stu-id="550f6-162">*Figure 4. Task pane*</span></span>

![Utiliser les volets Office en plus des commandes de complément](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="550f6-164">Extension des fonctionnalités Outlook</span><span class="sxs-lookup"><span data-stu-id="550f6-164">Extend Outlook functionality</span></span>

<span data-ttu-id="550f6-p113">Les compléments Outlook peuvent développer le ruban Office et s’afficher en regard d’un élément Outlook quand vous le visualisez ou le composez. Ils fonctionnent avec un message électronique, une demande de réunion, une réponse à une demande de réunion, une annulation de réunion ou un rendez-vous quand l’utilisateur visualise un élément reçu, répond à un élément ou en crée un.</span><span class="sxs-lookup"><span data-stu-id="550f6-p113">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="550f6-p114">Les compléments Outlook peuvent accéder à des informations contextuelles à partir de l’élément, telles qu’une adresse ou un ID de suivi, et utiliser ces données pour accéder à d’autres informations sur le serveur ou provenant de services web pour créer des expériences utilisateur attrayantes. Dans la plupart des cas, un complément Outlook peut être exécuté sans modification sur les différentes applications hôte prise en charge, notamment Outlook, Outlook pour Mac, Outlook Web App et Outlook Web App pour appareils, afin d’offrir une expérience homogène sur le bureau, en ligne, sur les tablettes et sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="550f6-p114">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and Outlook Web App for devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span> 

<span data-ttu-id="550f6-169">Pour accéder à une vue d’ensemble des compléments Outlook, reportez-vous à la rubrique [Présentation des compléments Outlook](/outlook/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="550f6-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](/outlook/add-ins/).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="550f6-170">Création d’objets dans des documents Office</span><span class="sxs-lookup"><span data-stu-id="550f6-170">Create new objects in Office documents</span></span>

<span data-ttu-id="550f6-p115">Vous pouvez incorporer des objets web, appelés compléments de contenu, dans des documents Excel et PowerPoint. Ces compléments de contenu vous permettent d’intégrer des visualisations de données web enrichies, du contenu multimédia (comme un lecteur vidéo YouTube ou une galerie d’images) et d’autres types de contenu externe.</span><span class="sxs-lookup"><span data-stu-id="550f6-p115">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="550f6-173">*Figure 5. Complément de contenu*</span><span class="sxs-lookup"><span data-stu-id="550f6-173">*Figure 5. Content add-in*</span></span>

![Incorporer des objets web appelés compléments de contenu](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="550f6-175">API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="550f6-175">Office JavaScript APIs</span></span>

<span data-ttu-id="550f6-p116">Les API JavaScript Office sont composées d’objets et de membres permettant de créer des compléments et d’interagir avec le contenu Office et les services web. Il existe un modèle objet commun que se partagent Excel, Outlook, Word, PowerPoint, OneNote et Project. Il existe également des modèles objet plus complets et propres à l’hôte pour Excel et Word. Ces API permettent d’accéder à des objets connus tels que des paragraphes et des classeurs, ce qui facilite la création de complément pour un hôte spécifique.</span><span class="sxs-lookup"><span data-stu-id="550f6-p116">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="550f6-180">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="550f6-180">Next steps</span></span>

<span data-ttu-id="550f6-181">Pour plus d’informations sur la façon de commencer à créer votre complément Office, essayez nos [ressources de démarrage rapide en 5 minutes](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="550f6-181">To learn more about how to start building your Office Add-in, try out our [5-minute Quick Starts](/office/dev/add-ins/).</span></span> <span data-ttu-id="550f6-182">Vous pouvez commencer à créer des compléments instant en utilisant Visual Studio ou tout autre éditeur.</span><span class="sxs-lookup"><span data-stu-id="550f6-182">You can start building add-ins right away using Visual Studio or any other editor.</span></span> 

<span data-ttu-id="550f6-183">Pour commencer à concevoir des solutions offrant des expériences utilisateur efficaces et attrayantes, consultez les [instructions de conception](../design/add-in-design.md) et les [meilleures pratiques](../concepts/add-in-development-best-practices.md) pour les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="550f6-183">To start planning solutions that create effective and compelling user experiences, get familiar with the [design guidelines](../design/add-in-design.md) and [best practices](../concepts/add-in-development-best-practices.md) for Office Add-ins.</span></span>    

## <a name="see-also"></a><span data-ttu-id="550f6-184">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="550f6-184">See also</span></span>

- [<span data-ttu-id="550f6-185">Exemples de compléments Office</span><span class="sxs-lookup"><span data-stu-id="550f6-185">Office Add-in samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
- [<span data-ttu-id="550f6-186">Présentation de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="550f6-186">Understanding the JavaScript API for Office</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="550f6-187">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="550f6-187">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
