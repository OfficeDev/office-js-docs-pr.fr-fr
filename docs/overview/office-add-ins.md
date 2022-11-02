---
title: Vue d’ensemble de la plateforme de compléments pour Office
description: Utilisez des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour étendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes.
ms.date: 04/14/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 5a780fcc1f863fb6803e2f719fc27338d4a6c366
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810112"
---
# <a name="office-add-ins-platform-overview"></a>Vue d’ensemble de la plateforme de compléments pour Office

You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Outlook, Excel, Word, PowerPoint, OneNote, and Project. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.

![L’application Office et un site web incorporé (complément) offrent des possibilités d’extensibilité infinies.](../images/addins-overview.png)

Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:

- **Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose functionality from Microsoft and others in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.

- **Créer de nouveaux objets interactifs et enrichis qui peuvent être incorporés dans des documents Office :** vous pouvez incorporer des cartes, des graphiques et des visualisations interactives que les utilisateurs peuvent ajouter à leurs feuilles de calcul Excel et présentations PowerPoint.

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a>En quoi les compléments Office sont-ils différents des compléments COM et VSTO ?

COM or VSTO add-ins are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the application (for example, Excel), reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.

![Les raisons d’utiliser les compléments Office : multiplateforme, déploiement centralisé, accès facile via AppSource et basées sur des technologies web standard.](../images/why.png)

Les compléments Office offrent les avantages suivants par rapport aux compléments créés à l’aide de VBA, COM ou VSTO.

- Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.

- Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.

- Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.

- Based on standard web technology. You can use any library you like to build Office Add-ins.

## <a name="components-of-an-office-add-in"></a>Composants d’un complément Office

An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.

### <a name="manifest"></a>Manifeste

Le manifeste est un fichier XML qui spécifie les paramètres et les fonctionnalités du complément, notamment :

- Le nom d’affichage, la description, l’ID, la version et les paramètres régionaux par défaut du complément.

- La façon dont le complément s’intègre à Office.  

- Le niveau d’autorisation et les conditions d’accès aux données pour le complément.

### <a name="web-app"></a>Application web

The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office client application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.

![Composants d’un complément Hello World.](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>Extension des clients Office et interaction avec ces clients

Les compléments Office offrent les possibilités suivantes dans une application cliente Office.

- Étendre les fonctionnalités (toutes les applications Office)

- Créer de nouveaux objets (Excel ou PowerPoint)

### <a name="extend-office-functionality"></a>Étendre les fonctionnalités d’Office

Vous pouvez ajouter de nouvelles fonctionnalités aux applications Office via les éléments d’interface suivants :  

- Commandes de menu et boutons de ruban personnalisées (collectivement appelés « commandes de complément »)

- Volets Office à insérer

Les éléments d’interface personnalisés et les volets Office sont définis dans le manifeste du complément.  

#### <a name="custom-buttons-and-menu-commands"></a>Commandes de menu et boutons personnalisés  

You can add custom ribbon buttons and menu items to the ribbon in Office on the web and on Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.  

![Boutons personnalisés et commandes de menu.](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>Volets Office  

You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.

![Utilisez des volets de tâches en plus des commandes de complément.](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>Extension des fonctionnalités Outlook

Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.

Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook application to provide a seamless experience on the desktop, web, and tablet and mobile devices.

Pour accéder à une vue d’ensemble des compléments Outlook, reportez-vous à la rubrique [Présentation des compléments Outlook](../outlook/outlook-add-ins-overview.md).

### <a name="create-new-objects-in-office-documents"></a>Création d’objets dans des documents Office

You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.

![Incorporer des objets web appelés compléments de contenu.](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>API JavaScript pour Office

Les API JavaScript Office sont composées d’objets et de membres permettant de créer des compléments et d’interagir avec le contenu Office et les services web. Il existe un modèle objet commun partagé par Excel, Outlook, Word, PowerPoint, OneNote et Project. Il existe également des modèles objet spécifiques à l’application plus complets pour Excel et Word. Ces API permettent d’accéder à des objets connus tels que des paragraphes et des classeurs, ce qui facilite la création d’un complément pour une application spécifique.

## <a name="next-steps"></a>Étapes suivantes

Pour une présentation en détails du développement des compléments Office, voir [Développement de compléments Office](../develop/develop-overview.md).

## <a name="see-also"></a>Voir aussi

- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publier des compléments Office](../publish/publish.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
