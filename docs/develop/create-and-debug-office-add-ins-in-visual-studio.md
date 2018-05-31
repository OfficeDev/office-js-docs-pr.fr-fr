---
title: Créer et déboguer des compléments Office dans Visual Studio
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 3e4fbcd3919be0d5510b36ae77a6e3706eab9689
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437604"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="35ff6-102">Créer et déboguer des compléments Office dans Visual Studio</span><span class="sxs-lookup"><span data-stu-id="35ff6-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="35ff6-p101">Cet article explique comment utiliser Visual Studio pour créer votre premier complément Office. Les étapes décrites dans cet article concernent Visual Studio 2015. Si vous utilisez une autre version de Visual Studio, les procédures peuvent légèrement varier.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="35ff6-106">Si vous débutez avec les compléments pour OneNote, reportez-vous à [Créer votre premier complément OneNote](../onenote/onenote-add-ins-getting-started.md).</span><span class="sxs-lookup"><span data-stu-id="35ff6-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="35ff6-107">Créer un projet de complément Office dans Visual Studio</span><span class="sxs-lookup"><span data-stu-id="35ff6-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="35ff6-p102">Pour commencer, vérifiez que les [outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx) sont installés et que vous disposez d’une version de Microsoft Office. Vous pouvez participer au [programme pour les développeurs Office 365](https://developer.microsoft.com/en-us/office/dev-program), ou suivre ces instructions pour obtenir la [version la plus récente](../develop/install-latest-office-version.md).</span><span class="sxs-lookup"><span data-stu-id="35ff6-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/en-us/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>


1. <span data-ttu-id="35ff6-110">Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="35ff6-111">Dans la liste des types de projets sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments web**, puis sélectionnez un des projets de compléments.</span><span class="sxs-lookup"><span data-stu-id="35ff6-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>  
    
3. <span data-ttu-id="35ff6-112">Nommez le projet, puis cliquez sur **OK** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-112">Name the project, and then choose  **OK** to create the project.</span></span>
    
4. <span data-ttu-id="35ff6-p103">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. La page par défaut Home.html s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The default Home.html page opens in Visual Studio.</span></span>
    
<span data-ttu-id="35ff6-115">Dans Visual Studio 2015, certains des modèles de projet de complément ont été mis à jour pour refléter des fonctionnalités supplémentaires :</span><span class="sxs-lookup"><span data-stu-id="35ff6-115">In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:</span></span>


- <span data-ttu-id="35ff6-p104">Des compléments de contenu peuvent apparaître dans le corps des documents Access et PowerPoint, en plus des feuilles de calcul Excel. Vous pouvez également choisir l’option de projet de base pour créer un projet de complément ayant un contenu élémentaire avec code de démarrage minimal, ou l’option Projet de visualisation de documents (pour Access et Excel seulement) afin de créer un complément dont le contenu est plus complet, qui inclut un code de démarrage pour visualiser et se lier à des données.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p104">Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.</span></span>
    
- <span data-ttu-id="35ff6-118">Les compléments Outlook comprennent des options permettant non seulement d’inclure votre complément dans un message électronique ou un rendez-vous, mais aussi d’indiquer si le complément est disponible lorsqu’un message électronique ou un rendez-vous est composé et lu.</span><span class="sxs-lookup"><span data-stu-id="35ff6-118">Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.</span></span>
    

> [!NOTE]
> <span data-ttu-id="35ff6-p105">Dans Visual Studio, la plupart des options sont compréhensibles par leurs descriptions, sauf la case à cocher **Message électronique**. Cochez cette case si vous souhaitez créer un complément Outlook qui apparaît non seulement avec les éléments de messagerie, mais aussi avec les demandes de réunion, les réponses et les annulations.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p105">In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.</span></span>

<span data-ttu-id="35ff6-121">Lorsque vous avez terminé l’Assistant, Visual Studio crée une solution qui contient deux projets.</span><span class="sxs-lookup"><span data-stu-id="35ff6-121">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span>



|<span data-ttu-id="35ff6-122">**Projet**</span><span class="sxs-lookup"><span data-stu-id="35ff6-122">**Project**</span></span>|<span data-ttu-id="35ff6-123">**Description**</span><span class="sxs-lookup"><span data-stu-id="35ff6-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="35ff6-124">Projet de complément</span><span class="sxs-lookup"><span data-stu-id="35ff6-124">Add-in project</span></span>|<span data-ttu-id="35ff6-p106">Contient seulement un fichier de manifeste XML, qui contient tous les paramètres qui décrivent votre complément. Ces paramètres aident l’hôte Office à déterminer quand votre complément doit être activé et où il doit apparaître. Visual Studio génère le contenu de ce fichier pour vous afin que vous puissiez exécuter le projet et utiliser immédiatement votre complément. Vous pouvez modifier ces paramètres à tout moment à l’aide de l’éditeur de manifeste.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p106">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="35ff6-129">Projet d’application web</span><span class="sxs-lookup"><span data-stu-id="35ff6-129">Web application project</span></span>|<span data-ttu-id="35ff6-p107">Contient les pages de contenu de votre complément, notamment tous les fichiers et références de fichiers dont vous avez besoin pour développer des pages HTML et JavaScript compatibles avec Office. Pendant que vous développez votre complément, Visual Studio héberge l’application web sur votre serveur IIS local. Lorsque vous êtes prêt à la publier, vous devez trouver un serveur pour héberger ce projet.Pour en savoir plus sur les projets d’applications web ASP.NET, voir [Projets web ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="35ff6-p107">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="35ff6-133">Modifier les paramètres de votre complément</span><span class="sxs-lookup"><span data-stu-id="35ff6-133">Modify your add-in settings</span></span>


<span data-ttu-id="35ff6-p108">Pour modifier les paramètres de votre complément, modifiez le fichier manifeste XML du projet. Dans l’**Explorateur de solutions**, développez le nœud de projet du complément et le dossier contenant le manifeste XML, puis sélectionnez le manifeste XML. Vous pouvez pointer sur n’importe quel élément du fichier pour afficher une info-bulle qui décrit l’objectif de l’élément. Pour plus d’informations sur le fichier manifeste, voir l’article sur le [manifeste XML de compléments Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="35ff6-p108">To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="35ff6-138">Développer le contenu de votre complément</span><span class="sxs-lookup"><span data-stu-id="35ff6-138">Develop the contents of your add-in</span></span>


<span data-ttu-id="35ff6-139">Alors que le projet de complément vous permet de modifier les paramètres qui décrivent le complément, l’application web fournit le contenu qui apparaît dans le complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-139">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="35ff6-p109">Le projet d’application web contient une page HTML par défaut et le fichier JavaScript que vous pouvez utiliser pour commencer. Il contient également un fichier JavaScript commun à toutes les pages que vous ajoutez à votre projet. Ces fichiers sont pratiques car ils contiennent des références à d’autres bibliothèques JavaScript, notamment l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p109">The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> 

<span data-ttu-id="35ff6-p110">Au fur et à mesure que votre complément devient plus complexe, vous pouvez ajouter d’autres fichiers HTML et JavaScript. Vous pouvez utiliser le contenu des fichiers HTML et JavaScript par défaut comme exemples des types de références que vous pouvez ajouter à d’autres pages de votre projet pour les faire fonctionner avec votre complément. Le tableau suivant décrit les fichiers HTML et JavaScript par défaut.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p110">As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.</span></span>



|<span data-ttu-id="35ff6-146">**Fichier**</span><span class="sxs-lookup"><span data-stu-id="35ff6-146">**File**</span></span>|<span data-ttu-id="35ff6-147">**Description**</span><span class="sxs-lookup"><span data-stu-id="35ff6-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="35ff6-148">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="35ff6-148">**Home.html**</span></span>|<span data-ttu-id="35ff6-p111">Situé dans le dossier  **de base** du projet ; il s’agit de la page HTML par défaut du complément. Cette page apparaît en tant que première page du complément lorsqu’elle est activée dans un élément de rendez-vous, de message électronique ou de document. Ce fichier est utile, car il contient toutes les références de fichiers dont vous avez besoin pour commencer. Lorsque vous êtes prêt à créer votre premier complément, ajoutez votre code HTML à ce fichier.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p111">Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.</span></span>|
|<span data-ttu-id="35ff6-153">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="35ff6-153">**Home.js**</span></span>|<span data-ttu-id="35ff6-p112">Situé dans le dossier  **de base** du projet ; il s’agit du fichier JavaScript associé à la page Home.js. Vous pouvez placer tout code propre au comportement de la page Home.html dans le fichier Home.js. Ce dernier contient un exemple de code pour vous aider à commencer.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p112">Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="35ff6-157">**App.js**</span><span class="sxs-lookup"><span data-stu-id="35ff6-157">**App.js**</span></span>|<span data-ttu-id="35ff6-p113">Situé dans le dossier  **Complément** du projet ; il s’agit du fichier JavaScript par défaut de l’ensemble du complément. Vous pouvez placer tout code commun au comportement de plusieurs pages de votre application dans le fichier App.js. Ce dernier contient un exemple de code pour vous aider à commencer.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p113">Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.</span></span>|

> [!NOTE]
> <span data-ttu-id="35ff6-p114">Vous n’êtes pas obligé d’utiliser ces fichiers. N’hésitez pas à ajouter d’autres fichiers au projet et à les utiliser à la place. Si vous souhaitez voir apparaître un autre fichier HTML comme page initiale du complément, ouvrez l’éditeur de manifeste et définissez la propriété **SourceLocation** sur le nom du fichier.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p114">You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>


## <a name="debug-your-add-in"></a><span data-ttu-id="35ff6-164">Déboguer votre complément</span><span class="sxs-lookup"><span data-stu-id="35ff6-164">Debug your add-in</span></span>


<span data-ttu-id="35ff6-165">Lorsque vous êtes prêt à démarrer votre complément, vérifiez les propriétés liées à la génération et au débogage, puis démarrez la solution.</span><span class="sxs-lookup"><span data-stu-id="35ff6-165">When you are ready to start your add-in, review build and debug related properties, and then start the solution.</span></span>


### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="35ff6-166">Réviser les propriétés de génération et de débogage</span><span class="sxs-lookup"><span data-stu-id="35ff6-166">Review the build and debug properties</span></span>

<span data-ttu-id="35ff6-p115">Avant de démarrer la solution, assurez-vous que Visual Studio va ouvrir l’application hôte souhaitée. Cette information apparaît dans les pages de propriétés du projet avec d’autres propriétés liées à la génération et au débogage du complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p115">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>


### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="35ff6-169">Pour ouvrir les pages de propriétés d’un projet</span><span class="sxs-lookup"><span data-stu-id="35ff6-169">To open the property pages of a project</span></span>


1. <span data-ttu-id="35ff6-170">Dans l’ **Explorateur de solutions**, choisissez le nom du projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-170">In  **Solution Explorer**, choose the project name.</span></span>
    
2. <span data-ttu-id="35ff6-171">Dans la barre de menus, choisissez  **Affichage**,  **Fenêtre Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-171">On the menu bar, choose  **View**,  **Properties Window**.</span></span>
    
<span data-ttu-id="35ff6-172">Le tableau suivant décrit les propriétés du projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-172">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="35ff6-173">**Propriété**</span><span class="sxs-lookup"><span data-stu-id="35ff6-173">**Property**</span></span>|<span data-ttu-id="35ff6-174">**Description**</span><span class="sxs-lookup"><span data-stu-id="35ff6-174">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="35ff6-175">**Action de démarrage**</span><span class="sxs-lookup"><span data-stu-id="35ff6-175">**Start Action**</span></span>|<span data-ttu-id="35ff6-176">Indique si votre complément doit être débogué dans un client de bureau Office ou dans un client Office Online dans le navigateur spécifié.</span><span class="sxs-lookup"><span data-stu-id="35ff6-176">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="35ff6-177">**Document de démarrage** (compléments de contenu et du volet Office uniquement)</span><span class="sxs-lookup"><span data-stu-id="35ff6-177">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="35ff6-178">Spécifie le document à ouvrir lors du démarrage du projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-178">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="35ff6-179">**Projet Web**</span><span class="sxs-lookup"><span data-stu-id="35ff6-179">**Web Project**</span></span>|<span data-ttu-id="35ff6-180">Spécifie le nom du projet web associé au complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-180">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="35ff6-181">**Adresse de messagerie** (compléments Outlook uniquement)</span><span class="sxs-lookup"><span data-stu-id="35ff6-181">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="35ff6-182">Spécifie l’adresse de messagerie du compte d’utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="35ff6-182">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="35ff6-183">**URL EWS** (compléments Outlook uniquement)</span><span class="sxs-lookup"><span data-stu-id="35ff6-183">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="35ff6-184">URL de service web Exchange (par exemple : https://www.contoso.com/ews/exchange.aspx).</span><span class="sxs-lookup"><span data-stu-id="35ff6-184">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="35ff6-185">**URL OWA** (compléments Outlook uniquement)</span><span class="sxs-lookup"><span data-stu-id="35ff6-185">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="35ff6-186">URL d’application web Outlook (par exemple : https://www.contoso.com/owa).</span><span class="sxs-lookup"><span data-stu-id="35ff6-186">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="35ff6-187">**Nom d’utilisateur** (compléments Outlook uniquement)</span><span class="sxs-lookup"><span data-stu-id="35ff6-187">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="35ff6-188">Spécifie le nom de votre compte d’utilisateur dans Exchange Server ou Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="35ff6-188">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="35ff6-189">**Fichier du projet**</span><span class="sxs-lookup"><span data-stu-id="35ff6-189">**Project File**</span></span>|<span data-ttu-id="35ff6-190">Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-190">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="35ff6-191">**Dossier du projet**</span><span class="sxs-lookup"><span data-stu-id="35ff6-191">**Project Folder**</span></span>|<span data-ttu-id="35ff6-192">Emplacement du fichier de projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-192">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="35ff6-193">Utiliser un document existant pour déboguer le complément (compléments de contenu et du volet Office uniquement)</span><span class="sxs-lookup"><span data-stu-id="35ff6-193">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>


<span data-ttu-id="35ff6-p116">Vous pouvez ajouter des documents au projet de complément. Si vous disposez d’un document qui contient des données de test que vous souhaitez utiliser avec votre application, Visual Studio ouvre ce document lorsque vous commencez le projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p116">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="35ff6-196">Pour utiliser un document existant pour déboguer le complément</span><span class="sxs-lookup"><span data-stu-id="35ff6-196">To use an existing document to debug the add-in</span></span>


1. <span data-ttu-id="35ff6-197">Dans l’ **Explorateur de solutions**, choisissez le dossier du projet de complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-197">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="35ff6-198">Choisissez le projet de complément et non le projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="35ff6-198">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="35ff6-199">Dans le menu **Projet**, choisissez **Ajouter un élément existant**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-199">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="35ff6-200">Dans la boîte de dialogue  **Ajouter un élément existant**, recherchez et sélectionnez le document que vous souhaitez ajouter.</span><span class="sxs-lookup"><span data-stu-id="35ff6-200">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="35ff6-201">Choisissez le bouton  **Ajouter** pour ajouter le document à votre projet.</span><span class="sxs-lookup"><span data-stu-id="35ff6-201">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="35ff6-202">Dans l’ **Explorateur de solutions**, ouvrez le menu contextuel du projet, puis choisissez  **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-202">In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.</span></span>
    
    <span data-ttu-id="35ff6-203">Les pages des propriétés relatives au projet s’affichent.</span><span class="sxs-lookup"><span data-stu-id="35ff6-203">The property pages for the project appear.</span></span>
    
6. <span data-ttu-id="35ff6-204">Dans la liste  **Document de démarrage**, choisissez le document que vous avez ajouté au projet, puis cliquez sur le bouton  **OK** pour fermer les pages de propriétés.</span><span class="sxs-lookup"><span data-stu-id="35ff6-204">In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.</span></span>
    

### <a name="start-the-solution"></a><span data-ttu-id="35ff6-205">Démarrer la solution</span><span class="sxs-lookup"><span data-stu-id="35ff6-205">Start the solution</span></span>


<span data-ttu-id="35ff6-p117">Visual Studio génère automatiquement la solution lorsque vous la démarrez. Vous pouvez la démarrer à partir de la barre de  **Menu** en choisissant **Débogage**,  **Démarrer**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p117">Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**.</span></span> 


> [!NOTE]
> <span data-ttu-id="35ff6-p118">Si le débogage de script n’est pas activé dans Internet Explorer, vous ne pourrez pas démarrer le débogueur dans Visual Studio. Pour activer le débogage de script, ouvrez la boîte de dialogue **Options Internet**, choisissez l’onglet **Avancé**, puis désélectionnez les cases à cocher **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p118">If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.</span></span>

<span data-ttu-id="35ff6-210">Visual Studio génère le projet et effectue les actions suivantes.</span><span class="sxs-lookup"><span data-stu-id="35ff6-210">Visual Studio builds the project and does the following:</span></span>


1. <span data-ttu-id="35ff6-p119">Il crée une copie du fichier manifeste XML et l’ajoute au répertoire  _ProjectName_\Output. L’application hôte utilise cette copie lorsque vous démarrez Visual Studio et déboguez l’application.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p119">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="35ff6-213">Il crée un ensemble d’entrées dans le Registre de votre ordinateur qui permettent au complément d’apparaître dans l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="35ff6-213">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="35ff6-214">Il génère le projet d’application web, puis le déploie sur le serveur web IIS local (http://localhost).</span><span class="sxs-lookup"><span data-stu-id="35ff6-214">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="35ff6-215">Visual Studio effectue ensuite les actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="35ff6-215">Next, Visual Studio does the following:</span></span>


1. <span data-ttu-id="35ff6-216">Il modifie l'élément [emplacement source](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) du fichier manifeste XML en remplaçant le jeton ~remoteAppUrl par l'adresse complète de la page de démarrage (par exemple, http://localhost/MyAgave.html).</span><span class="sxs-lookup"><span data-stu-id="35ff6-216">Modifies the SourceLocation element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="35ff6-217">Il démarre le projet d’application web dans IIS Express.</span><span class="sxs-lookup"><span data-stu-id="35ff6-217">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="35ff6-218">Il ouvre l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="35ff6-218">Opens the host application.</span></span> 
    
<span data-ttu-id="35ff6-p120">Visual Studio n’affiche pas les erreurs de validation dans la fenêtre  **OUTPUT** lorsque vous générez le projet. Visual Studio signale au fur et à mesure les erreurs et les avertissements dans la fenêtre **ERRORLIST**. Visual Studio signale également les erreurs de validation en affichant des traits de soulignement ondulés (appelés aussi zigzags) de différentes couleurs dans le code et l’éditeur de texte. Ces marques sont là pour vous indiquer les problèmes détectés par Visual Studio dans votre code. Pour plus d’informations, voir la page relative au [code et à l’éditeur de texte](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). Pour plus d’informations sur l’activation ou la désactivation de la validation, voir les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="35ff6-p120">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- <span data-ttu-id="35ff6-225">[Options, Éditeur de texte, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="35ff6-225">[Options, Text Editor, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)</span></span>
    
- <span data-ttu-id="35ff6-226">[Procédure : définir des options de validation pour l’édition HTML dans Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="35ff6-226">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="35ff6-227">[Validation, CSS, Éditeur de texte, boîte de dialogue Options](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="35ff6-227">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="35ff6-228">Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="35ff6-228">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a><span data-ttu-id="35ff6-229">Afficher un complément dans Excel, Word ou Project, et avancer pas à pas dans votre code</span><span class="sxs-lookup"><span data-stu-id="35ff6-229">Show an add-in in Excel, Word, or Project and step through your code</span></span>


<span data-ttu-id="35ff6-p121">Si vous définissez la propriété  **Document de démarrage** du projet de complément sur Excel ou Word, Visual Studio crée un document et le complément apparaît. Si vous définissez la propriété **Document de démarrage** du projet de complément afin d’utiliser un document existant, Visual Studio ouvre le document, mais vous devez insérer manuellement le complément. Si vous définissez la propriété **Document de démarrage** sur **Microsoft Project**, vous devez également insérer le complément manuellement.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p121">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually. If you set the **Start Document** to **Microsoft Project**, you also have to insert the add-in manually.</span></span>


### <a name="to-show-an-office-add-in-in-excel-or-word"></a><span data-ttu-id="35ff6-233">Pour afficher une Complément Office dans Excel ou Word</span><span class="sxs-lookup"><span data-stu-id="35ff6-233">To show an Office Add-in in Excel or Word</span></span>


1. <span data-ttu-id="35ff6-234">Dans Excel ou Word, dans l’onglet  **Insertion**, choisissez  **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-234">In Excel or Word, on the  **Insert** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="35ff6-235">Dans la liste qui apparaît, choisissez votre complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-235">In the list that appears, choose your add-in.</span></span>
    

### <a name="to-show-an-office-add-in-in-project"></a><span data-ttu-id="35ff6-236">Pour afficher une Complément Office dans Project</span><span class="sxs-lookup"><span data-stu-id="35ff6-236">To show an Office Add-in in Project</span></span>


1. <span data-ttu-id="35ff6-237">Dans Project, dans l’onglet  **Projet**, choisissez  **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-237">In Project, on the  **Project** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="35ff6-238">Dans la liste qui apparaît, choisissez votre complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-238">In the list that appears, choose your add-in.</span></span>
    
<span data-ttu-id="35ff6-p122">Dans Visual Studio, vous pouvez définir des points d’interruption, puis pendant l’interaction avec votre complément et l’exécution pas à pas du code de vos fichiers HTML, JavaScript et C# ou VB.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p122">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="35ff6-241">Afficher le complément Outlook dans Outlook et avancer pas à pas dans votre code</span><span class="sxs-lookup"><span data-stu-id="35ff6-241">Show the Outlook add-in in Outlook and step through your code</span></span>


<span data-ttu-id="35ff6-242">Pour voir le complément dans Outlook, ouvrez un message électronique ou un élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="35ff6-242">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="35ff6-p123">Outlook active le complément pour l’élément à condition que les critères d’activation soient respectés. La barre complément apparaît en haut de la fenêtre de l’inspecteur ou du volet de lecture, et votre complément Outlook apparaît sous la forme d’un bouton dans la barre du complément. Si votre complément est doté d’une commande, un bouton apparaît dans le ruban (soit dans l’onglet par défaut, soit dans un onglet personnalisé indiqué), et le complément n’apparaît pas dans la barre complément.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p123">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="35ff6-246">Pour voir votre complément Outlook, cliquez sur le bouton correspondant.</span><span class="sxs-lookup"><span data-stu-id="35ff6-246">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="35ff6-p124">Dans Visual Studio, vous pouvez définir des points d’interruption, puis pendant l’interaction avec votre complément Outlook et l’exécution pas à pas du code de vos fichiers HTML, JavaScript et C# ou VB.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p124">In Visual Studio, you can set break-points. Then, as you interact with your Outlook add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span> 

<span data-ttu-id="35ff6-p125">Vous pouvez également modifier votre code et vérifier les effets de ces modifications dans votre complément Outlook sans devoir fermer le Complément Office ni redémarrer le projet. Dans Outlook, ouvrez simplement le menu contextuel du complément Outlook, puis choisissez **Recharger**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p125">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="35ff6-251">Modifier le code et continuer le débogage du complément sans redémarrer le projet</span><span class="sxs-lookup"><span data-stu-id="35ff6-251">Modify code and continue to debug the add-in without having to start the project again</span></span>


<span data-ttu-id="35ff6-p126">Vous pouvez modifier votre code et vérifier les effets de ces modifications dans votre complément sans avoir à fermer l’application hôte et à redémarrer le projet. Après avoir modifié votre code, ouvrez le menu contextuel du complément, puis choisissez  **Recharger**. Quand vous rechargez le complément, il est déconnecté du débogueur Visual Studio. Vous pouvez constater les effets de vos modifications, mais vous ne pouvez pas parcourir pas à pas le code tant que vous n’attachez pas le débogueur Visual Studio à tous les processus Iexplore.exe disponibles.</span><span class="sxs-lookup"><span data-stu-id="35ff6-p126">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**. When you reload the add-in it becomes disconnected with the Visual Studio debugger. Therefore, you can view the effects of your change, but you cannot step through your code again until you attach the Visual Studio debugger to all of the available Iexplore.exe processes.</span></span>


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a><span data-ttu-id="35ff6-256">Pour attacher le débogueur Visual Studio à tous les processus Iexplore.exe disponibles</span><span class="sxs-lookup"><span data-stu-id="35ff6-256">To attach the Visual Studio debugger to all of the available Iexplore.exe processes</span></span>


1. <span data-ttu-id="35ff6-257">Dans Visual Studio, choisissez  **DÉBOGUER**,  **Attacher au processus**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-257">In Visual Studio, choose  **DEBUG**,  **Attach to Process**.</span></span>
    
2. <span data-ttu-id="35ff6-258">Dans la boîte de dialogue  **Attacher au processus**, choisissez tous les processus  **Iexplore.exe** disponibles, puis sélectionnez le bouton **Attacher**.</span><span class="sxs-lookup"><span data-stu-id="35ff6-258">In the  **Attach to Process** dialog box, choose all of the available **Iexplore.exe** processes, and then choose the **Attach** button.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="35ff6-259">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="35ff6-259">Next steps</span></span>

- [<span data-ttu-id="35ff6-260">Déploiement et publication de votre complément Office</span><span class="sxs-lookup"><span data-stu-id="35ff6-260">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
