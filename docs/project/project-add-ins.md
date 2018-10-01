---
title: Compléments du volet Office pour Project
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 2aa8a88878082357949935305b9d39d203f5fb5d
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348162"
---
# <a name="task-pane-add-ins-for-project"></a><span data-ttu-id="bd9f8-102">Compléments du volet de tâches pour Project</span><span class="sxs-lookup"><span data-stu-id="bd9f8-102">Task pane add-ins for Project</span></span>

<span data-ttu-id="bd9f8-103">Project Standard 2013 et Project Professional 2013 (version 15.1 ou ultérieure) comprennent tous les deux la prise en charge des compléments de volet Office. Vous pouvez exécuter les compléments de volet Office généraux qui sont développés pour Word 2013 ou Excel 2013.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-103">Project Standard 2013 and Project Professional 2013 (version 15.1 or higher) both include support for task pane add-ins. You can run general task pane add-ins that are developed for Word 2013 or Excel 2013.</span></span> <span data-ttu-id="bd9f8-104">Vous pouvez également développer des compléments personnalisés qui gèrent les événements de sélection de projet et intégrer des tâches, ressources, affichage et autres données au niveau des cellules dans un projet avec les listes SharePoint, les compléments SharePoint, les composants Web Part, les services Web et les applications d’entreprise.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-104">Project Standard 2013 and Project Professional 2013 both include support for task pane add-ins. You can run general task pane add-ins that are developed for Word 2013 or Excel 2013. You can also develop custom add-ins that handle selection events in Project and integrate task, resource, view, and other cell-level data in a project with SharePoint lists, SharePoint Add-ins, Web Parts, web services, and enterprise applications.</span></span>

> [!NOTE]
> <span data-ttu-id="bd9f8-p102">Le [téléchargement du kit de développement logiciel (SDK) de Project 2013](https://www.microsoft.com/download/details.aspx?id=30435%20) inclut des exemples de compléments qui montrent comment utiliser le modèle objet du complément pour Project et le service OData pour la création de rapports de données dans Project Server 2013. Après avoir extrait et installé le SDK, accédez au sous-dossier `\Samples\Apps\`.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p102">The [Project 2013 SDK download](https://www.microsoft.com/download/details.aspx?id=30435%20) includes sample add-ins that show how to use the add-in object model for Project, and how to use the OData service for reporting data in Project Server 2013. When you extract and install the SDK, see the `\Samples\Apps\` subdirectory.</span></span>

<span data-ttu-id="bd9f8-107">Pour une introduction aux compléments Office, reportez-vous à [Vue d’ensemble de la plateforme des compléments pour Office](../overview/office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-107">For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).</span></span>

## <a name="add-in-scenarios-for-project"></a><span data-ttu-id="bd9f8-108">Scénarios de compléments pour Project</span><span class="sxs-lookup"><span data-stu-id="bd9f8-108">Add-in scenarios for Project</span></span>

<span data-ttu-id="bd9f8-p103">Les gestionnaires de projet peuvent utiliser les compléments du volet Office dans Project pour les aider dans la gestion de leurs activités. Au lieu de quitter Project et d’ouvrir une nouvelle application pour rechercher les informations qu’ils utilisent fréquemment, les gestionnaires de projet peuvent accéder directement à ces informations à partir de Project. Le contenu d’un complément du volet Office peut être contextuel, basé sur la tâche sélectionnée, la ressource, la vue ou d’autres données dans une cellule de diagramme de Gantt, une vue Utilisation des tâches ou une vue Utilisation des ressources.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p103">Project managers can use Project task pane add-ins to help with project management activities. Instead of leaving Project and opening another application to search for frequently used information, project managers can directly access the information within Project. The content in a task pane add-in can be context-sensitive, based on the selected task, resource, view, or other data in a cell in a Gantt chart, task usage view, or resource usage view.</span></span>

> [!NOTE]
> <span data-ttu-id="bd9f8-112">Avec Project Professionnel 2013, vous pouvez développer des compléments du volet Office qui accèdent aux installations locales de Project Server 2013 et Project Online, ainsi qu’aux versions locales ou en ligne de SharePoint 2013. Project Standard 2013 ne prend pas en charge l’intégration directe aux données Project Server ou aux listes de tâches SharePoint synchronisées avec Project Server.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-112">With Project Professional 2013, you can develop task pane add-ins that access on-premises installations of Project Server 2013, Project Online, and on-premises or online SharePoint 2013.Project Standard 2013 does not support direct integration with Project Server data or SharePoint task lists that are synchronized with Project Server.</span></span>

<span data-ttu-id="bd9f8-113">Les scénarios des compléments pour Project comprennent les fonctionnalités suivantes :</span><span class="sxs-lookup"><span data-stu-id="bd9f8-113">Add-in scenarios for Project include the following:</span></span>

-  <span data-ttu-id="bd9f8-p104">**Planification de projet** Affichez des données de projets associés pouvant avoir une influence sur la planification. Un complément du volet Office peut intégrer les données d’autres projets dans Project Server 2013. Par exemple, vous pouvez afficher un ensemble de projets du service avec les dates importantes, ou afficher les données d’autres projets sur la base d’un champ personnalisé sélectionné.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p104">**Project scheduling** View data from related projects that can affect scheduling. A task pane add-in can integrate relevant data from other projects in Project Server 2013. For example, you can view the departmental collection of projects and milestone dates, or view specified data from other projects that are based on a selected custom field.</span></span>
    
-  <span data-ttu-id="bd9f8-117">**Gestion des ressources** Affichez un ensemble complet des ressources dans Project Server 2013, ou un sous-ensemble basé sur les compétences spécifiques, notamment les coûts et la disponibilité, afin de sélectionner les ressources appropriées.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-117">**Resource management** View the complete resource pool in Project Server 2013 or a subset based on specified skills, including cost data and resource availability, to help select appropriate resources.</span></span>
    
-  <span data-ttu-id="bd9f8-p105">**État des tâches et approbation** Utilisez une application Web dans un complément du volet Office pour mettre à jour ou afficher les données d’une application ERP (enterprise resource planning) externe, d’un système de feuille de temps, ou d’une application de comptabilité. Vous pouvez également créer un composant WebPart d’approbation des états personnalisé utilisable à la fois dans une application Project Web et Project Professionnel 2013.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p105">**Statusing and approvals** Use a web application in a task pane add-in to update or view data from an external enterprise resource planning (ERP) application, timesheet system, or accounting application. Or, create a custom status approval Web Part that can be used within both Project Web App and Project Professional 2013.</span></span>
    
-  <span data-ttu-id="bd9f8-p106">**Communication avec l’équipe** Communiquez avec les ressources et les membres de l’équipe directement à partir d’un complément du volet Office, dans le contexte d’un projet. Vous pouvez également gérer facilement un ensemble de notes contextuelles pour vous aider dans votre projet.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p106">**Team communication** Communicate with team members and resources directly from a task pane add-in, within the context of a project. Or, easily maintain a set of context-sensitive notes for yourself as you work in a project.</span></span>
    
-  <span data-ttu-id="bd9f8-p107">**Pack de travail** Recherchez des types spécifiques de modèles de projets dans les bibliothèques SharePoint et les librairies, et dans les collections de modèles en ligne. Par exemple, trouvez des modèles pour la construction de vos projets et ajoutez-les à votre collection de modèles Project.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p107">**Work packages** Search for specified kinds of project templates within SharePoint libraries and online template collections. For example, find templates for construction projects and add them to your Project template collection.</span></span>
    
-  <span data-ttu-id="bd9f8-p108">**Éléments associés** Affichez des métadonnées, des documents et des messages associés aux tâches spécifiques dans un plan de projet. Par exemple, vous pouvez utiliser Project Professionnel 2013 pour gérer un projet importé à partir d’une liste de tâches SharePoint, et continuer de synchroniser la liste des tâches avec les changements apportés au projet. Un complément du volet Office peut afficher des champs supplémentaires ou des métadonnées que Project n’a pas importés avec les tâches de la liste SharePoint.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p108">**Related items** View metadata, documents, and messages that are related to specific tasks in a project plan. For example, you can use Project Professional 2013 to manage a project that was imported from a SharePoint task list, and still synchronize the task list with changes in the project. A task pane add-in can show additional fields or metadata that Project did not import for tasks in the SharePoint list.</span></span>
    
-  <span data-ttu-id="bd9f8-p109">**Utilisation des modèles objet Project Server** Utilisez le GUID d’une tâche sélectionnée avec les méthodes dans l’interface PSI (Project Server Interface) ou le modèle objet côté client (CSOM) de Project Server. Par exemple, l’application Web d'un complément peut lire et mettre à jour les données d’états d’une tâche et d’une ressource sélectionnées, ou s’intégrer à une application de feuille de temps externe.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p109">**Use the Project Server object models** Use the GUID of a selected task with methods in the Project Server Interface (PSI) or the client-side object model (CSOM) of Project Server. For example, the web application for an add-in can read and update the statusing data of a selected task and resource, or integrate with an external timesheet application.</span></span>
    
-  <span data-ttu-id="bd9f8-p110">**Obtenir les données de rapports** Utilisez REST (Representational State Transfer), JavaScript, ou les requêtes LINQ pour trouver les informations associées à une tâche ou ressource sélectionnée dans le service OData pour les tableaux de rapports d'application Project Web. Les requêtes qui utilisent le service OData peuvent être créées avec une installation de Project Server 2013 en ligne ou locale.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p110">**Get reporting data** Use Representational State Transfer (REST), JavaScript, or LINQ queries to find related information for a selected task or resource in the OData service for reporting tables in Project Web App. Queries that use the OData service can be done with an online or an on-premises installation of Project Server 2013.</span></span>
    
    <span data-ttu-id="bd9f8-131">Par exemple, reportez-vous à [Créer un complément Project qui utilise REST avec un service OData Project Server local](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-131">For example, see [Create a Project add-in that uses REST with an on-premises Project Server OData  service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span></span>
    
## <a name="developing-project-add-ins"></a><span data-ttu-id="bd9f8-132">Développement de compléments pour Project</span><span class="sxs-lookup"><span data-stu-id="bd9f8-132">Developing Project add-ins</span></span>

<span data-ttu-id="bd9f8-p111">La bibliothèque JavaScript pour les compléments Project comprend des extensions de l’alias de l’espace de nom **Office** qui permet aux développeurs d’accéder aux propriétés de l’application Project, ainsi qu’aux tâches, ressources et vues dans un projet. Les extensions de la bibliothèque JavaScript du fichier Project-15.js sont utilisées dans un complément Project créé avec Visual Studio 2015. Les fichiers Office.js, Office.debug.js, Project-15.js, Project-15.debug.js et autres fichiers associés sont également fournis dans le téléchargement du Kit de développement logiciel (SDK) Project 2013.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p111">The JavaScript library for Project add-ins includes extensions of the  **Office** namespace alias that enable developers to access properties of the Project application and tasks, resources, and views in a project. The JavaScript library extensions in the Project-15.js file are used in a Project add-in created with Visual Studio 2015. The Office.js, Office.debug.js, Project-15.js, Project-15.debug.js, and related files are also provided in the Project 2013 SDK download.</span></span>

<span data-ttu-id="bd9f8-p112">Pour créer un complément, vous pouvez utiliser un éditeur de texte simple afin de créer une page Web HTML avec les fichiers JavaScript associés, les fichiers CSS et les requêtes REST. Outre une page HTML ou une application Web, le complément nécessite un fichier manifeste XML pour la configuration. Project peut utiliser un fichier manifeste qui inclut un attribut  **type** spécifié comme **TaskPaneExtension**. Le fichier manifeste peut être utilisé par plusieurs applications clientes Office 2013, ou vous pouvez créer un fichier manifeste spécifique pour Project 2013. Pour plus d’informations, voir la section  _Notions fondamentales de développement_ dans [Vue d’ensemble de la plateforme des compléments pour Office](../overview/office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p112">To create an add-in, you can use a simple text editor to create an HTML webpage and related JavaScript files, CSS files, and REST queries. In addition to an HTML page or a web application, an add-in requires an XML manifest file for configuration. Project can use a manifest file that includes a  **type** attribute that is specified as **TaskPaneExtension**. The manifest file can be used by multiple Office 2013 client applications, or you can create a manifest file that is specific for Project 2013. For more information, see the  _Development basics_ section in [Office Add-ins platform overview](../overview/office-add-ins.md).</span></span>

<span data-ttu-id="bd9f8-p113">Pour les applications complexes personnalisées, et pour un débogage plus facile, nous vous recommandons d’utiliser Visual Studio 2015 pour développer des sites Web pour les compléments. Visual Studio 2015 contient des modèles pour les projets de compléments qui permettent de choisir le type de complément (volet Office, contenu ou messagerie) et l’application hôte (Project, Word, Excel ou Outlook). Pour obtenir un exemple qui intègre des données de Project Online, reportez-vous à l’article relatif à la [connexion d’un complément du volet Office Project à PWA](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx) dans le blog Project Programmability sur MSDN.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p113">For complex custom applications, and for easier debugging, we recommend that you use Visual Studio 2015 to develop websites for add-ins. Visual Studio 2015 include templates for add-in projects, where you can choose the kind of add-in (task pane, content, or mail) and the host application (Project, Word, Excel, or Outlook).  For an example that integrates with data from Project Online, see [Connecting a Project task pane add-in to PWA](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx) in the Project Programmability blog on MSDN.</span></span>

<span data-ttu-id="bd9f8-143">Lorsque vous installez le Kit de développement logiciel (SDK) de Project 2013, le sous-répertoire `\Samples\Apps\` inclut les exemples de compléments suivants :</span><span class="sxs-lookup"><span data-stu-id="bd9f8-143">When you install the Project 2013 SDK download, the  `\Samples\Apps\` subdirectory includes the following sample add-ins:</span></span>


-  <span data-ttu-id="bd9f8-p114">**Bing Search :** le fichier manifeste BingSearch.xml pointe vers la page de recherche Bing pour les périphériques mobiles. Comme l’application Web Bing existe déjà sur Internet, le complément de recherche Bing n’utilise pas d’autres fichiers de code source ou le modèle objet de complément pour Project.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p114">**Bing Search:** The BingSearch.xml manifest file points to the Bing search page for mobile devices. Because the Bing web app already exists on the Internet, the Bing Search add-in does not use other source code files or the add-in object model for Project.</span></span>
    
-  <span data-ttu-id="bd9f8-p115">**Projet Test MO :** le fichier manifeste JSOM_SimpleOMCalls.xml et le fichier JSOM_Call.html constituent, ensemble, un exemple de test du modèle objet et de la fonctionnalité de complément dans Project 2013. Le fichier HTML fait référence au fichier JSOM_Sample.js, qui contient des fonctions JavaScript qui utilisent le fichier Office.js et le fichier Project-15.js pour les fonctionnalités de base. Le téléchargement du SDK contient tous les fichiers de code source nécessaires et le fichier manifeste XML pour le complément Projet Test MO. Le développement et l’installation de l’exemple Projet Test MO sont décrits dans [Créer votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p115">**Project OM Test:** The JSOM_SimpleOMCalls.xml manifest file and the JSOM_Call.html file are, together, an example that tests the object model and add-in functionality in Project 2013. The HTML file references the JSOM_Sample.js file, which has JavaScript functions that use the Office.js file and the Project-15.js file for the primary functionality. The SDK download includes all of the necessary source code files and the manifest XML file for the Project OM Test add-in. The development and installation of the Project OM Test sample is described in [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>
    
-  <span data-ttu-id="bd9f8-p116">**HelloProject_OData :** solution Visual Studio pour Project Professionnel 2013 qui résume les données du projet actuel, telles que les coûts, le travail et le pourcentage accompli, et les compare avec la moyenne de tous les projets publiés dans l’instance d'application Project Web où le projet actif est stocké. Le développement, l’installation et le test de cet exemple qui utilise le protocole REST avec le service **ProjectData** dans l'application Project Web, sont décrits dans [Créer un complément Project qui utilise REST avec un service OData Project Server local](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p116">**HelloProject_OData:** This is a Visual Studio solution for Project Professional 2013 that summarizes data from the active project, such as cost, work, and percent complete, and compares that with the average for all published projects in the Project Web App instance where the active project is stored. The development, installation, and testing of the sample, which uses the REST protocol with the **ProjectData** service in Project Web App, is described in [Create a Project add-in that uses REST with an on-premises Project Server OData service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span></span>
    

### <a name="creating-an-add-in-manifest-file"></a><span data-ttu-id="bd9f8-152">Création d’un fichier manifeste de complément</span><span class="sxs-lookup"><span data-stu-id="bd9f8-152">Creating an add-in manifest file</span></span>


<span data-ttu-id="bd9f8-153">Le fichier manifeste spécifie l’URL de la page Web du complément ou l’application Web, le type de complément (volet Office pour Project), les URL facultatives de contenus pour d’autres langues ou paramètres régionaux, et d’autres propriétés.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-153">The manifest file specifies the URL of the add-in webpage or web application, the kind of add-in (task pane for Project), optional URLs of content for other languages and locales, and other properties.</span></span>


### <a name="procedure-1-to-create-the-add-in-manifest-file-for-bing-search"></a><span data-ttu-id="bd9f8-p117">Procédure 1. Créer le fichier manifeste du complément pour Bing Search</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p117">Procedure 1. To create the add-in manifest file for Bing Search</span></span>


- <span data-ttu-id="bd9f8-p118">Créez un fichier XML dans un répertoire local. Le fichier XML inclut l’élément  **OfficeApp**, et ses éléments enfants, qui sont décrits dans [Manifeste XML des compléments Office](../develop/add-in-manifests.md). Par exemple, créez un fichier nommé BingSearch.xml qui contient le code XML suivant.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p118">Create an XML file in a local directory. The XML file includes the  **OfficeApp** element and child elements, which are described in the [Office Add-ins XML manifest](../develop/add-in-manifests.md). For example, create a file named BingSearch.xml that contains the following XML.</span></span>
    
    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
                xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp">
      <Id>1234-5678</Id>
      <Version>15.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-us</DefaultLocale>
      <DisplayName DefaultValue="Bing Search">
      </DisplayName>
      <Description DefaultValue="Search selected data on Bing">
      </Description>
      <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
      </IconUrl>
      <Capabilities>
        <Capability Name="Project"/>
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="http://m.bing.com">
        </SourceLocation>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

- <span data-ttu-id="bd9f8-159">Les éléments suivants sont requis dans le manifeste du complément :</span><span class="sxs-lookup"><span data-stu-id="bd9f8-159">Following are the required elements in the add-in manifest:</span></span>
  - <span data-ttu-id="bd9f8-160">Dans l’élément  **OfficeApp**, l’attribut `xsi:type="TaskPaneApp"` spécifie que le complément est de type volet Office.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-160">In the  **OfficeApp** element, the `xsi:type="TaskPaneApp"` attribute specifies that the add-in is a task pane type.</span></span>
  - <span data-ttu-id="bd9f8-161">L’élément  **Id** est un UUID et doit être unique.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-161">The  **Id** element is a UUID and must be unique.</span></span>
  - <span data-ttu-id="bd9f8-p119">L’élément  **Version** indique la version du complément. L’élément **ProviderName** correspond au nom de l’entreprise ou du développeur qui fournit le complément. L’élément **DefaultLocale** spécifie les paramètres régionaux par défaut pour les chaînes du manifeste.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p119">The  **Version** element is the version of the add-in. The **ProviderName** element is the name of the company or developer who provides the add-in. The **DefaultLocale** element specifies the default locale for the strings in the manifest.</span></span>
  - <span data-ttu-id="bd9f8-p120">L’élément  **DisplayName** correspond au nom qui s’affiche dans la liste déroulante **Complément du volet Office** de l’onglet **AFFICHAGE**, dans le ruban de Project 2013. La valeur du nom peut contenir jusqu’à 32 caractères.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p120">The  **DisplayName** element is the name that shows in the **Task Pane Add-in** drop-down list in the **VIEW** tab of the ribbon in Project 2013. The value can contain up to 32 characters.</span></span>
  - <span data-ttu-id="bd9f8-p121">L’élément  **Description** contient la description du complément pour les paramètres régionaux par défaut. La valeur peut contenir jusqu’à 2 000 caractères.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p121">The  **Description** element contains the add-in description for the default locale. The value can contain up to 2000 characters.</span></span>
  - <span data-ttu-id="bd9f8-169">L’élément  **Capabilities** contient un ou plusieurs éléments enfants **Capability** qui spécifient l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-169">The  **Capabilities** element contains one or more **Capability** child elements that specify the host application.</span></span>
  - <span data-ttu-id="bd9f8-p122">L’élément  **DefaultSettings** inclut l’élément **SourceLocation**, qui spécifie le chemin d’accès d’un fichier HTML sur un partage de fichiers ou l’URL d’une page Web que le complément utilise. Un complément du volet Office ignore l’élément **RequestedHeight** et l’élément **RequestedWidth**.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p122">The  **DefaultSettings** element includes the **SourceLocation** element, which specifies the path of an HTML file on a file share or the URL of a webpage that the add-in uses. A task pane add-in ignores the **RequestedHeight** element and the **RequestedWidth** element.</span></span>
  - <span data-ttu-id="bd9f8-p123">L’élément **IconUrl** est facultatif. Il peut être une icône sur un partage de fichiers ou l’URL d’une icône dans une application Web.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p123">The  **IconUrl** element is optional. It can be an icon on a file share or the URL of an icon in a web application.</span></span>
    
- <span data-ttu-id="bd9f8-p124">(Facultatif) Ajoutez des éléments  **Override** qui ont des valeurs pour les autres paramètres régionaux. Par exemple, le manifeste suivant fournit des éléments **Override** pour les valeurs françaises de **DisplayName**,  **Description**,  **IconUrl** et **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p124">(Optional) Add  **Override** elements that have values for other locales. For example, the following manifest provides **Override** elements for French values of **DisplayName**,  **Description**,  **IconUrl**, and  **SourceLocation**.</span></span>
    
    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
                xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp">
      <Id>1234-5678</Id>
      <Version>15.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-us</DefaultLocale>
      <DisplayName DefaultValue="Bing Search">
        <Override Locale="fr-fr" Value="Bing Search"/>
      </DisplayName>
      <Description DefaultValue="Search selected data on Bing">
        <Override Locale="fr-fr" Value="Search selected data on Bing"></Override>
      </Description>
      <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
        <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
      </IconUrl>
      <Capabilities>
        <Capability Name="Project"/>
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="http://m.bing.com">
          <Override Locale="fr-fr" Value="http://m.bing.com"/>
        </SourceLocation>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```


## <a name="installing-project-add-ins"></a><span data-ttu-id="bd9f8-176">Installation de compléments Project</span><span class="sxs-lookup"><span data-stu-id="bd9f8-176">Installing Project add-ins</span></span>


<span data-ttu-id="bd9f8-p125">Dans Project 2013, vous pouvez installer des compléments comme solutions autonomes sur un partage de fichiers ou dans un catalogue de compléments privé. Vous pouvez également consulter et acheter des compléments dans AppSource.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p125">In Project 2013, you can install add-ins as stand-alone solutions on a file share, or in a private add-in catalog. You can also review and purchase add-ins in AppSource.</span></span>

<span data-ttu-id="bd9f8-p126">Un partage de fichiers peut contenir plusieurs fichiers manifestes XML de complément et sous-répertoires. Vous pouvez ajouter ou supprimer des catalogues et des emplacements de répertoire Manifest à l’aide de l’onglet  **Catalogues de compléments approuvés** dans la boîte de dialogue **Centre de gestion de la confidentialité** dans Project 2013. Pour afficher un complément dans Project, l’élément **SourceLocation** dans un manifeste doit pointer vers un site Web existant ou un fichier source HTML.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p126">There can be multiple add-in manifest XML files and subdirectories in a file share. You can add or remove manifest directory locations and catalogs by using the  **Trusted Add-in Catalogs** tab in the **Trust Center** dialog box in Project 2013. To show an add-in in Project, the **SourceLocation** element in a manifest must point to an existing website or HTML source file.</span></span>


> [!NOTE]
> <span data-ttu-id="bd9f8-p127">Internet Explorer 9 (ou version ultérieure) doit être installé, sans obligation de le définir comme navigateur par défaut. Les compléments Office nécessitent des composants d’Internet Explorer 9. Le navigateur par défaut peut être Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, ou une version ultérieure de ces navigateurs.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p127">Internet Explorer 9 (or later) must be installed, but does not have to be the default browser. Office Add-ins require components in Internet Explorer 9. The default browser can be Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.</span></span>

<span data-ttu-id="bd9f8-p128">Dans la procédure 2, le complément Bing Search est installé sur l’ordinateur local où Project 2013 est installé. Toutefois, comme l’infrastructure du complément n’utilise pas directement les chemins de fichiers locaux tels que  `C:\Project\AppManifests`, vous pouvez créer un partage de fichiers sur l’ordinateur local. Si vous préférez, vous pouvez créer un partage de fichiers sur un ordinateur à distance.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p128">In Procedure 2, the Bing Search add-in is installed on the local computer where Project 2013 is installed. However, because the add-in infrastructure does not directly use local file paths such as  `C:\Project\AppManifests`, you can create a network share on the local computer. If you prefer, you can create a file share on a remote computer.</span></span>


### <a name="procedure-2-to-install-the-bing-search-add-in"></a><span data-ttu-id="bd9f8-p129">Procédure 2. Installer le complément Bing Search</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p129">Procedure 2. To install the Bing Search add-in</span></span>


1. <span data-ttu-id="bd9f8-p130">Créez un répertoire local pour les fichiers manifestes des compléments. Par exemple, créez un répertoire qui s’appelle  `C:\Project\AppManifests`.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p130">Create a local directory for add-in manifests. For example, create the  `C:\Project\AppManifests` directory.</span></span>
    
2. <span data-ttu-id="bd9f8-192">Partagez le répertoire  `C:\Project\AppManifests` comme AppManifests, pour que le chemin du partage de fichiers sur le réseau devienne  `\\ServerName\AppManifests`.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-192">Share the  `C:\Project\AppManifests` directory asAppManifests, so the network path to the file share becomes  `\\ServerName\AppManifests`.</span></span>
    
3. <span data-ttu-id="bd9f8-193">Copiez le fichier manifeste BingSearch.xml dans le répertoire  `C:\Project\AppManifests`.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-193">Copy the BingSearch.xml manifest file to the  `C:\Project\AppManifests` directory.</span></span>
    
4. <span data-ttu-id="bd9f8-194">Dans Project 2013, ouvrez la boîte de dialogue  **Options de Project**, choisissez  **Centre de gestion de la confidentialité**, puis choisissez **Paramètres du centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-194">In Project 2013, open the  **Project Options** dialog box, choose **Trust Center**, and then choose  **Trust Center Settings**.</span></span>
    
5. <span data-ttu-id="bd9f8-195">Dans la boîte de dialogue  **Centre de gestion de la confidentialité**, dans le volet de gauche, choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-195">In the  **Trust Center** dialog box, in the left pane, choose **Trusted Add-in Catalogs**.</span></span>
    
6. <span data-ttu-id="bd9f8-196">Dans le volet  **Catalogues de compléments approuvés** (voir la figure 1), ajoutez le chemin `\\ServerName\AppManifests` dans la zone de texte **URL du catalogue**, choisissez  **Ajouter un catalogue**, puis choisissez **OK**.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-196">In the  **Trusted Add-in Catalogs** pane (see Figure 1), add the `\\ServerName\AppManifests` path in the **Catalog Url** text box, choose **Add Catalog**, and then choose  **OK**.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="bd9f8-p131">La figure 1 présente deux partages de fichiers et une URL hypothétique associée à un catalogue privé dans la liste **Adresse du catalogue approuvé**. Un seul partage de fichiers peut être défini comme partage par défaut et un seul catalogue d’URL peut être défini comme catalogue par défaut. Par exemple, si vous définissez `\\Server2\AppManifests` comme valeur par défaut, Project désélectionne la case à cocher **Par défaut** pour `\\ServerName\AppManifests`. Si vous changez la sélection par défaut, vous pouvez choisir **Effacer** pour supprimer des compléments installés, puis redémarrer Project. Si vous ajoutez un complément au partage de fichier par défaut ou au catalogue SharePoint alors que Project est ouvert, redémarrez Project.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p131">Figure 1 shows two file shares and one hypothetical URL for a private catalog in the  **Trusted Catalog Address** list. Only one file share can be the default file share and only one catalog URL can be the default catalog. For example, if you set `\\Server2\AppManifests` as the default, Project clears the **Default** check box for `\\ServerName\AppManifests`.If you change the default selection, you can choose  **Clear** to remove installed add-ins, and then restart Project. If you add an add-in to the default file share or SharePoint catalog while Project is open, you should restart Project.</span></span>

    <span data-ttu-id="bd9f8-201">*Figure 1. Utilisation du centre de gestion de la confidentialité pour ajouter des catalogues de manifestes de complément*</span><span class="sxs-lookup"><span data-stu-id="bd9f8-201">*Figure 1. Using the Trust Center to add catalogs of add-in manifests*</span></span>

    ![Utilisation du Centre de gestion de la confidentialité pour ajouter des manifestes d’application](../images/pj15-agave-overview-trust-centers.png)

7. <span data-ttu-id="bd9f8-p132">Dans le ruban **Project**, choisissez le menu déroulant  **Compléments Office**, puis choisissez **Afficher tout**. Dans la boîte de dialogue **Insérer un complément**, choisissez **DOSSIER PARTAGÉ** (voir la figure 2).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p132">On the  **Project** ribbon, choose the **Office Add-ins** drop-down menu, and then choose **See All**. In the  **Insert Add-in** dialog box, choose **SHARED FOLDER** (see Figure 2).</span></span>
    
    <span data-ttu-id="bd9f8-205">*Figure 2. Démarrage d’un complément se trouvant sur un partage de fichiers*</span><span class="sxs-lookup"><span data-stu-id="bd9f8-205">*Figure 2. Starting an add-in that is on a file share*</span></span>

    ![Démarrage d’une application Office dans un partage de fichiers](../images/pj15-agave-overview-start-agave-apps.png)

8. <span data-ttu-id="bd9f8-207">Sélectionnez le complément Bing Search, puis choisissez  **Insérer**.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-207">Select the Bing Search add-in, and then choose  **Insert**.</span></span>
    
    <span data-ttu-id="bd9f8-p133">Le complément Bing Search affiche un volet Office comme dans la figure 3. Vous pouvez redimensionner manuellement le volet Office et utiliser le complément Bing Search.</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p133">The Bing Search add-in shows in a task pane, as in Figure 3. You can manually resize the task pane, and use the Bing Search add-in.</span></span>

    <span data-ttu-id="bd9f8-210">*Figure 3. Utilisation du complément Recherche Bing*</span><span class="sxs-lookup"><span data-stu-id="bd9f8-210">*Figure 3. Using the Bing Search add-in*</span></span>

    ![Utilisation de l’application de recherche Bing](../images/pj15-agave-overview-bing-search.png)


## <a name="distributing-project-add-ins"></a><span data-ttu-id="bd9f8-212">Distribution de compléments Project</span><span class="sxs-lookup"><span data-stu-id="bd9f8-212">Distributing Project add-ins</span></span>


<span data-ttu-id="bd9f8-p134">Vous pouvez distribuer des compléments via un partage de fichiers, un catalogue de compléments dans une bibliothèque SharePoint ou dans AppSource. Pour plus d’informations, reportez-vous à [Publier votre complément Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="bd9f8-p134">You can distribute add-ins through a file share, an add-in catalog in a SharePoint library, or AppSource. For more information, see [Publish your Office Add-in](../publish/publish.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="bd9f8-215">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bd9f8-215">See also</span></span>

- [<span data-ttu-id="bd9f8-216">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="bd9f8-216">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="bd9f8-217">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="bd9f8-217">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="bd9f8-218">Interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="bd9f8-218">JavaScript API for Office</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
- [<span data-ttu-id="bd9f8-219">Créez votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte</span><span class="sxs-lookup"><span data-stu-id="bd9f8-219">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [<span data-ttu-id="bd9f8-220">Créer un complément Project qui utilise REST avec un service OData Project Server local</span><span class="sxs-lookup"><span data-stu-id="bd9f8-220">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
- [<span data-ttu-id="bd9f8-221">Connexion d’un complément du volet Office Project à PWA</span><span class="sxs-lookup"><span data-stu-id="bd9f8-221">Connecting a Project task pane add-in to PWA</span></span>](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx)
- [<span data-ttu-id="bd9f8-222">Téléchargement du Kit de développement logiciel (SDK) de Project 2013</span><span class="sxs-lookup"><span data-stu-id="bd9f8-222">Project 2013 SDK download</span></span>](https://www.microsoft.com/download/details.aspx?id=30435%20)
    
