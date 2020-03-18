---
title: Créer un complément Project qui utilise REST avec un service OData Project Server local
description: Découvrez comment créer un complément du volet Office pour Project Professional 2013 qui compare les données de coût et de travail du projet actif avec les moyennes de tous les projets de l’instance Project Web App actuelle.
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 6e53796f586e5ffa375f10cd877fc950b47203b4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720994"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a><span data-ttu-id="b4d03-103">Créer un complément Project qui utilise REST avec un service OData Project Server local</span><span class="sxs-lookup"><span data-stu-id="b4d03-103">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>

<span data-ttu-id="b4d03-104">Cet article décrit comment créer un complément de volet Office pour Project Professionnel 2013, qui compare les données de coût et les données de travail du projet actif avec les moyennes de tous les projets de l’instance actuelle de Project Web App.</span><span class="sxs-lookup"><span data-stu-id="b4d03-104">This article describes how to build a task pane add-in for Project Professional 2013 that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance.</span></span> <span data-ttu-id="b4d03-105">Le complément utilise REST avec la bibliothèque jQuery pour accéder au service de création de rapports OData **ProjectData** dans Project Server 2013.</span><span class="sxs-lookup"><span data-stu-id="b4d03-105">The add-in uses REST with the jQuery library to access the **ProjectData** OData reporting service in Project Server 2013.</span></span>

<span data-ttu-id="b4d03-106">Le code de cet article est basé sur un exemple développé par Saurabh Sanghvi et Arvind Iyer, Microsoft Corporation.</span><span class="sxs-lookup"><span data-stu-id="b4d03-106">The code in this article is based on a sample developed by Saurabh Sanghvi and Arvind Iyer, Microsoft Corporation.</span></span>

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a><span data-ttu-id="b4d03-107">Conditions requises pour la création d’un complément du volet Office qui lit les données de rapport Project Server</span><span class="sxs-lookup"><span data-stu-id="b4d03-107">Prerequisites for creating a task pane add-in that reads Project Server reporting data</span></span>

<span data-ttu-id="b4d03-108">Les conditions requises pour la création d’un complément de volet de tâches de projet qui lit le service **ProjectData** d’une instance Project Web App dans une installation locale de project Server 2013 sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="b4d03-108">The following are the prerequisites for creating a Project task pane add-in that reads the **ProjectData** service of a Project Web App instance in an on-premises installation of Project Server 2013:</span></span>

- <span data-ttu-id="b4d03-p102">Assurez-vous d’avoir installé les mises à jour Windows et les Service Packs les plus récents sur votre ordinateur de développement local. Le système d’exploitation peut être Windows 7, Windows 8, Windows Server 2008 ou Windows Server 2012.</span><span class="sxs-lookup"><span data-stu-id="b4d03-p102">Ensure that you have installed the most recent service packs and Windows updates on your local development computer. The operating system can be Windows 7, Windows 8, Windows Server 2008, or Windows Server 2012.</span></span>

- <span data-ttu-id="b4d03-111">Project Professionnel 2013 est nécessaire pour la connexion à Project Web App.</span><span class="sxs-lookup"><span data-stu-id="b4d03-111">Project Professional 2013 is required to connect with Project Web App.</span></span> <span data-ttu-id="b4d03-112">Project Professionnel 2013 doit être installé sur l’ordinateur de développement pour pouvoir activer le débogage **F5** avec Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="b4d03-112">The development computer must have Project Professional 2013 installed to enable **F5** debugging with Visual Studio.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b4d03-113">Project Standard 2013 peut également héberger des compléments de volet Office, mais ne peut pas se connecter à Project Web App.</span><span class="sxs-lookup"><span data-stu-id="b4d03-113">Project Standard 2013 can also host task pane add-ins, but cannot log on to Project Web App.</span></span>

- <span data-ttu-id="b4d03-114">Visual Studio 2015 avec Outils de développement Office pour Visual Studio comprend des modèles permettant de créer des Compléments Office et SharePoint. Assurez-vous que vous avez installé la version la plus récente des outils de développement Office. Consultez la section  _Outils_ de la page relative aux [téléchargements de compléments Office et SharePoint](https://developer.microsoft.com/office/docs)</span><span class="sxs-lookup"><span data-stu-id="b4d03-114">Visual Studio 2015 with Office Developer Tools for Visual Studio includes templates for creating Office and SharePoint Add-ins. Ensure that you have installed the most recent version of Office Developer Tools; see the  _Tools_ section of the [Office Add-ins and SharePoint downloads](https://developer.microsoft.com/office/docs).</span></span>

- <span data-ttu-id="b4d03-115">Les procédures et les exemples de code dans cet article accèdent au service **ProjectData** de Project Server 2013 dans un domaine local.</span><span class="sxs-lookup"><span data-stu-id="b4d03-115">The procedures and code examples in this article access the **ProjectData** service of Project Server 2013 in a local domain.</span></span> <span data-ttu-id="b4d03-116">Les méthodes jQuery de cet article ne fonctionnent pas avec Project sur le Web.</span><span class="sxs-lookup"><span data-stu-id="b4d03-116">The jQuery methods in this article do not work with Project on the web.</span></span>

    <span data-ttu-id="b4d03-117">Vérifiez que le service **ProjectData** est accessible à partir de votre ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="b4d03-117">Verify that the **ProjectData** service is accessible from your development computer.</span></span>

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a><span data-ttu-id="b4d03-p105">Procédure 1. Pour vérifier que le service ProjectData est accessible</span><span class="sxs-lookup"><span data-stu-id="b4d03-p105">Procedure 1. To verify that the ProjectData service is accessible</span></span>

1. <span data-ttu-id="b4d03-p106">Pour permettre à votre navigateur d’afficher directement les données XML à partir d’une requête REST, désactivez le mode Lecture du flux. Pour plus d’informations sur la façon d’y parvenir dans Internet Explorer, voir la procédure 1, étape 4 dans [Interrogation des flux OData pour les données de création de rapports Project](/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="b4d03-p106">To enable your browser to directly show the XML data from a REST query, turn off the feed reading view. For information about how to do this in Internet Explorer, see Procedure 1, step 4 in [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

2. <span data-ttu-id="b4d03-122">Interrogez le service **ProjectData** à l’aide de votre navigateur avec l’URL suivante : \*\* http://ServerName /ProjectServerName/_API/ProjectData\*\*.</span><span class="sxs-lookup"><span data-stu-id="b4d03-122">Query the **ProjectData** service by using your browser with the following URL: **http://ServerName /ProjectServerName /_api/ProjectData**.</span></span> <span data-ttu-id="b4d03-123">Par exemple, si l’instance Project Web App est `http://MyServer/pwa`, le navigateur affiche les résultats suivants :</span><span class="sxs-lookup"><span data-stu-id="b4d03-123">For example, if the Project Web App instance is  `http://MyServer/pwa`, the browser shows the following results:</span></span>

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

3. <span data-ttu-id="b4d03-p108">Vous pouvez être amené à fournir vos informations d’identification réseau pour voir les résultats. Si le navigateur affiche un message similaire à « Erreur 403, accès refusé », cela signifie que vous n’avez pas d’autorisation d’ouverture de session pour cette instance de Project Web App, ou qu’il existe un problème réseau qui nécessite une aide de la part d’un administrateur.</span><span class="sxs-lookup"><span data-stu-id="b4d03-p108">You may have to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you do not have logon permission for that Project Web App instance, or there is a network problem that requires administrative help.</span></span>

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a><span data-ttu-id="b4d03-126">Utilisation de Visual Studio pour créer un complément du volet Office pour Project</span><span class="sxs-lookup"><span data-stu-id="b4d03-126">Using Visual Studio to create a task pane add-in for Project</span></span>

<span data-ttu-id="b4d03-127">Outils de développement Office pour Visual Studio comprend un modèle pour les compléments du volet Office pour Project 2013.</span><span class="sxs-lookup"><span data-stu-id="b4d03-127">Office Developer Tools for Visual Studio includes a template for task pane add-ins for Project 2013.</span></span> <span data-ttu-id="b4d03-128">Si vous créez une solution nommée **HelloProjectOData**, la solution contient les deux projets Visual Studio suivants :</span><span class="sxs-lookup"><span data-stu-id="b4d03-128">If you create a solution named **HelloProjectOData**, the solution contains the following two Visual Studio projects:</span></span>

- <span data-ttu-id="b4d03-129">Le projet de complément prend le nom de la solution.</span><span class="sxs-lookup"><span data-stu-id="b4d03-129">The add-in project takes the name of the solution.</span></span> <span data-ttu-id="b4d03-130">Il inclut le fichier manifeste XML du complément et cible .NET Framework 4.5.</span><span class="sxs-lookup"><span data-stu-id="b4d03-130">It includes the XML manifest file for the add-in and targets the .NET Framework 4.5.</span></span> <span data-ttu-id="b4d03-131">La procédure 3 montre les étapes à suivre pour modifier le manifeste pour le complément **HelloProjectOData** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-131">Procedure 3 shows the steps to modify the manifest for the **HelloProjectOData** add-in.</span></span>

- <span data-ttu-id="b4d03-132">Le projet Web est nommé **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-132">The web project is named **HelloProjectODataWeb**.</span></span> <span data-ttu-id="b4d03-133">Il comprend les pages web, les fichiers JavaScript, les fichiers CSS, les images, les références et les fichiers de configuration du contenu web dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="b4d03-133">It includes the webpages, JavaScript files, CSS files, images, references, and configuration files for the web content in the task pane.</span></span> <span data-ttu-id="b4d03-134">Le projet cible .NET Framework 4.</span><span class="sxs-lookup"><span data-stu-id="b4d03-134">The web project targets the .NET Framework 4.</span></span> <span data-ttu-id="b4d03-135">Les procédures 4 et 5 montrent comment modifier les fichiers du projet web pour créer les fonctionnalités du complément **HelloProjectOData**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-135">Procedure 4 and Procedure 5 show how to modify the files in the web project to create the functionality of the **HelloProjectOData** add-in.</span></span>

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a><span data-ttu-id="b4d03-136">Procédure 2.</span><span class="sxs-lookup"><span data-stu-id="b4d03-136">Procedure 2.</span></span> <span data-ttu-id="b4d03-137">Pour créer le complément HelloProjectOData pour Project</span><span class="sxs-lookup"><span data-stu-id="b4d03-137">To create the HelloProjectOData add-in for Project</span></span>

1. <span data-ttu-id="b4d03-138">Exécutez Visual Studio 2015 en tant qu’administrateur, puis sélectionnez **nouveau projet** sur la page de démarrage.</span><span class="sxs-lookup"><span data-stu-id="b4d03-138">Run Visual Studio 2015 as an administrator, and then select **New Project** on the Start page.</span></span>

2. <span data-ttu-id="b4d03-139">Dans la boîte de dialogue **nouveau projet** , développez les nœuds **modèles**, **Visual C#** et **Office/SharePoint** , puis sélectionnez \* \* Compléments Office \* \*.</span><span class="sxs-lookup"><span data-stu-id="b4d03-139">In the **New Project** dialog box, expand the **Templates**, **Visual C#**, and **Office/SharePoint** nodes, and then select \*\* Office Add-ins\*\*.</span></span> <span data-ttu-id="b4d03-140">Sélectionnez **.NET Framework 4.5.2** dans la liste déroulante Framework cible en haut du volet central, puis sélectionnez **complément Office** (voir la capture d’écran suivante).</span><span class="sxs-lookup"><span data-stu-id="b4d03-140">Select **.NET Framework 4.5.2** in the target framework drop-down list at the top of the center pane, and then select **Office Add-in** (see the next screenshot).</span></span>

3. <span data-ttu-id="b4d03-141">Pour placer les deux projets Visual Studio dans le même répertoire, sélectionnez **Créer le répertoire pour la solution**, puis accédez à l’emplacement de votre choix.</span><span class="sxs-lookup"><span data-stu-id="b4d03-141">To place both of the Visual Studio projects in the same directory, select **Create directory for solution**, and then browse to the location you want.</span></span>

4. <span data-ttu-id="b4d03-142">Dans le champ **nom** , tapez helloprojectodata, puis choisissez **OK**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-142">In the **Name** field, typeHelloProjectOData, and then choose **OK**.</span></span>

    <span data-ttu-id="b4d03-143">*Figure 1. Création d’un complément Office*</span><span class="sxs-lookup"><span data-stu-id="b4d03-143">*Figure 1. Creating an Office Add-in*</span></span>

    ![Création d’un complément Office](../images/pj15-hello-project-o-data-creating-app.png)

5. <span data-ttu-id="b4d03-145">Dans la boîte de dialogue **Choisir le type de complément**, sélectionnez **Volet Office** et choisissez **Suivant** (voir la capture d’écran suivante).</span><span class="sxs-lookup"><span data-stu-id="b4d03-145">In the **Choose the add-in type** dialog box, select **Task pane** and choose **Next** (see the next screenshot).</span></span>

    <span data-ttu-id="b4d03-146">*Figure 2. Choix du type de complément à créer*</span><span class="sxs-lookup"><span data-stu-id="b4d03-146">*Figure 2. Choosing the type of add-in to create*</span></span>

    ![Choix du type de complément à créer](../images/pj15-hello-project-o-data-choose-project.png)

6. <span data-ttu-id="b4d03-148">Dans la boîte de dialogue **Choisir les applications hôtes**, désélectionnez toutes les cases, sauf la case **Project** (voir la capture d’écran suivante) et cliquez sur **Terminer**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-148">In the **Choose the host applications** dialog box, clear all check boxes except the **Project** check box (see the next screenshot) and choose **Finish**.</span></span>

    <span data-ttu-id="b4d03-149">*Figure 3. Choix de l’application hôte*</span><span class="sxs-lookup"><span data-stu-id="b4d03-149">*Figure 3. Choosing the host application*</span></span>

    ![Sélection d’un projet comme application hôte unique](../images/create-office-add-in.png)

    <span data-ttu-id="b4d03-151">Visual Studio crée le projet **HelloProjectOdata** et le projet **HelloProjectODataWeb** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-151">Visual Studio creates the **HelloProjectOdata** project and the **HelloProjectODataWeb** project.</span></span>

<span data-ttu-id="b4d03-152">Le dossier **AddIn** (voir la capture d’écran suivante) contient le fichier App.css pour les styles CSS personnalisés.</span><span class="sxs-lookup"><span data-stu-id="b4d03-152">The **AddIn** folder (see the next screenshot) contains the App.css file for custom CSS styles.</span></span> <span data-ttu-id="b4d03-153">Dans le sous-dossier **Home**, le fichier Home.html contient des références aux fichiers CSS et aux fichiers JavaScript utilisés par le complément, et le contenu HTML5 pour le complément.</span><span class="sxs-lookup"><span data-stu-id="b4d03-153">In the **Home** subfolder , the Home.html file contains references to the CSS files and the JavaScript files that the add-in uses, and the HTML5 content for the add-in.</span></span> <span data-ttu-id="b4d03-154">Par ailleurs, le fichier Home.js est pour votre code JavaScript personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b4d03-154">Also, the Home.js file is for your custom JavaScript code.</span></span> <span data-ttu-id="b4d03-155">Le dossier **Scripts** inclut les fichiers de bibliothèque jQuery.</span><span class="sxs-lookup"><span data-stu-id="b4d03-155">The **Scripts** folder includes the jQuery library files.</span></span> <span data-ttu-id="b4d03-156">Le sous-dossier **Office** comprend les bibliothèques JavaScript telles que office.js et project-15.js, ainsi que les bibliothèques de langage pour les chaînes standard dans les compléments Office. Dans le dossier **Content**, le fichier Office.css contient les styles par défaut pour tous les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="b4d03-156">The **Office** subfolder includes the JavaScript libraries such as office.js and project-15.js, plus the language libraries for standard strings in the Office Add-ins. In the **Content** folder, the Office.css file contains the default styles for all of the Office Add-ins.</span></span>

<span data-ttu-id="b4d03-157">*Figure 4. Affichage des fichiers de projet web par défaut dans l’Explorateur de solutions*</span><span class="sxs-lookup"><span data-stu-id="b4d03-157">*Figure 4. Viewing the default web project files in Solution Explorer*</span></span>

![Affichage des fichiers de projet web dans l’Explorateur de solutions](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

<span data-ttu-id="b4d03-159">Le manifeste pour le projet **HelloProjectOData** est le fichier HelloProjectOData. Xml.</span><span class="sxs-lookup"><span data-stu-id="b4d03-159">The manifest for the **HelloProjectOData** project is the HelloProjectOData.xml file.</span></span> <span data-ttu-id="b4d03-160">Vous pouvez éventuellement modifier le manifeste pour ajouter une description du complément, une référence à une icône, des informations pour des langues supplémentaires et d’autres paramètres.</span><span class="sxs-lookup"><span data-stu-id="b4d03-160">You can optionally modify the manifest to add a description of the add-in, a reference to an icon, information for additional languages, and other settings.</span></span> <span data-ttu-id="b4d03-161">La procédure 3 modifie simplement le nom d’affichage et la description du complément, puis ajoute une icône.</span><span class="sxs-lookup"><span data-stu-id="b4d03-161">Procedure 3 simply modifies the add-in display name and description, and adds an icon.</span></span>

<span data-ttu-id="b4d03-162">Pour plus d’informations sur le manifeste, reportez-vous à la rubrique [Manifeste XML des compléments Office](../develop/add-in-manifests.md) et [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../develop/add-in-manifests.md#see-also).</span><span class="sxs-lookup"><span data-stu-id="b4d03-162">For more information about the manifest, see [Office Add-ins XML manifest](../develop/add-in-manifests.md) and [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md#see-also).</span></span>

### <a name="procedure-3-to-modify-the-add-in-manifest"></a><span data-ttu-id="b4d03-p116">Procédure 3. Pour modifier le manifeste du complément</span><span class="sxs-lookup"><span data-stu-id="b4d03-p116">Procedure 3. To modify the add-in manifest</span></span>

1. <span data-ttu-id="b4d03-165">Dans Visual Studio, ouvrez le fichier HelloProjectOData.xml.</span><span class="sxs-lookup"><span data-stu-id="b4d03-165">In Visual Studio, open the HelloProjectOData.xml file.</span></span>

2. <span data-ttu-id="b4d03-166">Le nom d’affichage par défaut est le nom du projet Visual Studio (« HelloProjectOData »).</span><span class="sxs-lookup"><span data-stu-id="b4d03-166">The default display name is the name of the Visual Studio project ("HelloProjectOData").</span></span> <span data-ttu-id="b4d03-167">Par exemple, remplacez la valeur par défaut de l’élément **DisplayName** par "Hello ProjectData".</span><span class="sxs-lookup"><span data-stu-id="b4d03-167">For example, change the default value of the **DisplayName** element to"Hello ProjectData".</span></span>

3. <span data-ttu-id="b4d03-p118">La description par défaut est également « HelloProjectOData ». Par exemple, remplacez la valeur par défaut de l’élément Description par "Test REST queries of the ProjectData service".</span><span class="sxs-lookup"><span data-stu-id="b4d03-p118">The default description is also "HelloProjectOData". For example, change the default value of the Description element to "Test REST queries of the ProjectData service".</span></span>

4. <span data-ttu-id="b4d03-170">Ajoutez une icône à afficher dans la liste déroulante **Compléments Office** sous l’onglet **PROJET** du ruban.</span><span class="sxs-lookup"><span data-stu-id="b4d03-170">Add an icon to show in the **Office Add-ins** drop-down list on the **PROJECT** tab of the ribbon.</span></span> <span data-ttu-id="b4d03-171">Vous pouvez ajouter un fichier d’icône dans la solution Visual Studio ou utiliser une URL pour une icône.</span><span class="sxs-lookup"><span data-stu-id="b4d03-171">You can add an icon file in the Visual Studio solution or use a URL for an icon.</span></span> 

<span data-ttu-id="b4d03-172">Les étapes suivantes montrent comment ajouter un fichier d’icône à la solution Visual Studio :</span><span class="sxs-lookup"><span data-stu-id="b4d03-172">The following steps show how to add an icon file to the Visual Studio solution:</span></span>

1. <span data-ttu-id="b4d03-173">Dans l' **Explorateur de solutions**, accédez au dossier nommé images.</span><span class="sxs-lookup"><span data-stu-id="b4d03-173">In **Solution Explorer**, go to the folder named Images.</span></span>

2. <span data-ttu-id="b4d03-174">Pour pouvoir être affichée dans la liste déroulante **Compléments Office**, l’icône doit avoir une taille de 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="b4d03-174">To be displayed in the **Office Add-ins** drop-down list, the icon must be 32 x 32 pixels.</span></span> <span data-ttu-id="b4d03-175">Par exemple, installez le Kit de développement logiciel (SDK) de Project 2013, puis sélectionnez le dossier **Images** et ajoutez le fichier suivant à partir du Kit de développement logiciel (SDK) : `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span><span class="sxs-lookup"><span data-stu-id="b4d03-175">For example, install the Project 2013 SDK, and then choose the **Images** folder and add the following file from the SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span></span>

    <span data-ttu-id="b4d03-176">Vous pouvez également utiliser votre propre icône 32 x 32. Sinon, copiez l’image suivante dans un fichier nommé NewIcon.png, puis ajoutez ce fichier dans le dossier `HelloProjectODataWeb\Images` :</span><span class="sxs-lookup"><span data-stu-id="b4d03-176">Alternately, use your own 32 x 32 icon; or, copy the following image to a file named NewIcon.png, and then add that file to the  `HelloProjectODataWeb\Images` folder:</span></span>

    ![Icône de l’application HelloProjectOData](../images/pj15-hello-project-data-new-icon.jpg)

3. <span data-ttu-id="b4d03-178">Dans le manifeste HelloProjectOData. xml, ajoutez un élément **IconUrl** sous l’élément **Description** , où la valeur de l’URL de l’icône est le chemin d’accès relatif au fichier d’icône 32x32.</span><span class="sxs-lookup"><span data-stu-id="b4d03-178">In the HelloProjectOData.xml manifest, add an **IconUrl** element below the **Description** element, where the value of the icon URL is the relative path to the 32x32 icon file.</span></span> <span data-ttu-id="b4d03-179">Par exemple, ajoutez la ligne suivante : **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-179">For example, add the following line: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**.</span></span> <span data-ttu-id="b4d03-180">Le fichier manifeste HelloProjectOData. xml contient désormais les éléments suivants (votre valeur **ID** sera différente) :</span><span class="sxs-lookup"><span data-stu-id="b4d03-180">The HelloProjectOData.xml manifest file now contains the following (your **Id** value will be different):</span></span>

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82</Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a><span data-ttu-id="b4d03-181">Création du contenu HTML pour le complément HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="b4d03-181">Creating the HTML content for the HelloProjectOData add-in</span></span>

<span data-ttu-id="b4d03-182">Le complément **HelloProjectOData** est un exemple qui inclut le débogage et la sortie d’erreur ; Il n’est pas destiné à une utilisation de production.</span><span class="sxs-lookup"><span data-stu-id="b4d03-182">The **HelloProjectOData** add-in is a sample that includes debugging and error output; it is not intended for production use.</span></span> <span data-ttu-id="b4d03-183">Avant de commencer à coder le contenu HTML, concevez l’interface utilisateur et l’expérience utilisateur du complément, et définissez également les fonctions JavaScript qui interagissent avec le code HTML.</span><span class="sxs-lookup"><span data-stu-id="b4d03-183">Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code.</span></span> <span data-ttu-id="b4d03-184">Pour plus d’informations, voir[Instructions de conception pour les compléments Office](../design/add-in-design.md).</span><span class="sxs-lookup"><span data-stu-id="b4d03-184">For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span></span> 

<span data-ttu-id="b4d03-185">Le volet Office affiche le nom complet du complément en haut, qui est la valeur de l’élément **DisplayName** dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="b4d03-185">The task pane shows the add-in display name at the top, which is the value of the **DisplayName** element in the manifest.</span></span> <span data-ttu-id="b4d03-186">L’élément **body** du fichier HelloProjectOData.html contient les autres éléments d’interface utilisateur, comme suit :</span><span class="sxs-lookup"><span data-stu-id="b4d03-186">The **body** element in the HelloProjectOData.html file contains the other UI elements, as follows:</span></span>

- <span data-ttu-id="b4d03-187">Un sous-titre indique la fonctionnalité générale ou le type de l’opération, par exemple **ODATA REST QUERY**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-187">A subtitle indicates the general functionality or type of operation, for example, **ODATA REST QUERY**.</span></span>

- <span data-ttu-id="b4d03-188">Le bouton **obtenir le point de terminaison ProjectData** appelle la `setOdataUrl` fonction pour obtenir le point de terminaison du service **ProjectData** , puis l’affiche dans une zone de texte.</span><span class="sxs-lookup"><span data-stu-id="b4d03-188">The **Get ProjectData Endpoint** button calls the `setOdataUrl` function to get the endpoint of the **ProjectData** service, and display it in a text box.</span></span> <span data-ttu-id="b4d03-189">Si Project n’est pas connecté à Project Web App, le complément appelle un gestionnaire d’erreur afin d’afficher un message d’erreur dans une fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="b4d03-189">If Project is not connected with Project Web App, the add-in calls an error handler to display a pop-up error message.</span></span>

- <span data-ttu-id="b4d03-190">Le bouton **Comparer tous les projets** est désactivé jusqu’à ce que le complément obtienne un point de terminaison OData valide.</span><span class="sxs-lookup"><span data-stu-id="b4d03-190">The **Compare All Projects** button is disabled until the add-in gets a valid OData endpoint.</span></span> <span data-ttu-id="b4d03-191">Lorsque vous sélectionnez le bouton, la `retrieveOData` fonction est appelée, ce qui fait appel à une requête REST pour obtenir les données de coût et de travail du projet à partir du service **ProjectData** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-191">When you select the button, it calls the `retrieveOData` function, which uses a REST query to get project cost and work data from the **ProjectData** service.</span></span>

- <span data-ttu-id="b4d03-p126">Un tableau affiche les valeurs moyennes relatives au coût du projet, au coût réel, au travail et au pourcentage achevé. Le tableau compare également les valeurs actuelles du projet actif à la moyenne. Si la valeur actuelle est supérieure à la moyenne de tous les projets, elle est affichée en rouge. Si la valeur actuelle est inférieure à la moyenne, la valeur est affichée en vert. Si la valeur actuelle n’est pas disponible, le tableau affiche **NA** en bleu.</span><span class="sxs-lookup"><span data-stu-id="b4d03-p126">A table displays the average values for project cost, actual cost, work, and percent complete. The table also compares the current active project values with the average. If the current value is greater than the average for all projects, the value is displayed as red. If the current value is less than the average, the value is displayed as green. If the current value is not available, the table displays a blue **NA**.</span></span>

    <span data-ttu-id="b4d03-197">La `retrieveOData` fonction appelle la `parseODataResult` fonction, qui calcule et affiche les valeurs du tableau.</span><span class="sxs-lookup"><span data-stu-id="b4d03-197">The `retrieveOData` function calls the `parseODataResult` function, which calculates and displays values for the table.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b4d03-198">Dans cet exemple, les données de coût et de travail du projet actif sont dérivées des valeurs publiées.</span><span class="sxs-lookup"><span data-stu-id="b4d03-198">In this example, cost and work data for the active project are derived from the published values.</span></span> <span data-ttu-id="b4d03-199">Si vous modifiez les valeurs dans Project, le service **ProjectData** ne dispose pas des modifications tant que le projet n’est pas publié.</span><span class="sxs-lookup"><span data-stu-id="b4d03-199">If you change values in Project, the **ProjectData** service does not have the changes until the project is published.</span></span>

### <a name="procedure-4-to-create-the-html-content"></a><span data-ttu-id="b4d03-200">Procédure 4.</span><span class="sxs-lookup"><span data-stu-id="b4d03-200">Procedure 4.</span></span> <span data-ttu-id="b4d03-201">Pour créer du contenu HTML</span><span class="sxs-lookup"><span data-stu-id="b4d03-201">To create the HTML content</span></span>

1. <span data-ttu-id="b4d03-202">Dans l’élément **Head** du fichier Home. html, ajoutez tous les éléments de **lien** supplémentaires pour les fichiers CSS utilisés par votre complément.</span><span class="sxs-lookup"><span data-stu-id="b4d03-202">In the **head** element of the Home.html file, add any additional **link** elements for CSS files that your add-in uses.</span></span> <span data-ttu-id="b4d03-203">Le modèle de projet Visual Studio inclut un lien pour le fichier App.css que vous pouvez utiliser pour des styles CSS personnalisés.</span><span class="sxs-lookup"><span data-stu-id="b4d03-203">The Visual Studio project template includes a link for the App.css file that you can use for custom CSS styles.</span></span>

2. <span data-ttu-id="b4d03-204">Ajoutez des éléments **script** supplémentaires pour les bibliothèques JavaScript utilisées par votre complément.</span><span class="sxs-lookup"><span data-stu-id="b4d03-204">Add any additional **script** elements for JavaScript libraries that your add-in uses.</span></span> <span data-ttu-id="b4d03-205">Le modèle de projet inclut des liens pour les fichiers jQuery- _[version]_. js, Office. js et MicrosoftAjax. js dans le dossier **scripts** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-205">The project template includes links for the jQuery- _[version]_.js, office.js, and MicrosoftAjax.js files in the **Scripts** folder.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b4d03-p131">Avant de déployer le complément, remplacez la référence à office.js et celle à jQuery par la référence au réseau de distribution de contenu. Cette dernière permet d’accéder à la version la plus récente et d’obtenir de meilleures performances.</span><span class="sxs-lookup"><span data-stu-id="b4d03-p131">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

    <span data-ttu-id="b4d03-208">Le complément **HelloProjectOData** utilise également le fichier SurfaceErrors. js, qui affiche les erreurs dans un message contextuel.</span><span class="sxs-lookup"><span data-stu-id="b4d03-208">The **HelloProjectOData** add-in also uses the SurfaceErrors.js file, which displays errors in a pop-up message.</span></span> <span data-ttu-id="b4d03-209">Vous pouvez copier le code à partir de la section _programmation fiable_ de la section [créer votre premier complément du volet Office pour le projet 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), puis ajouter un fichier SurfaceErrors. js dans le dossier **Scripts\Office** du projet **HelloProjectODataWeb** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-209">You can copy the code from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), and then add a SurfaceErrors.js file in the **Scripts\Office** folder of the **HelloProjectODataWeb** project.</span></span>

    <span data-ttu-id="b4d03-210">Voici le code HTML mis à jour pour l’élément **Head** , avec la ligne supplémentaire pour le fichier SurfaceErrors. js :</span><span class="sxs-lookup"><span data-stu-id="b4d03-210">Following is the updated HTML code for the **head** element, with the additional line for the SurfaceErrors.js file:</span></span>

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. <span data-ttu-id="b4d03-211">Dans l’élément **Body** , supprimez le code existant du modèle, puis ajoutez le code de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b4d03-211">In the **body** element, delete the existing code from the template, and then add the code for the user interface.</span></span> <span data-ttu-id="b4d03-212">Si un élément doit être rempli avec des données ou manipulé par une instruction jQuery, l’élément doit inclure un attribut  **id** unique.</span><span class="sxs-lookup"><span data-stu-id="b4d03-212">If an element is to be filled with data or manipulated by a jQuery statement, the element must include a unique **id** attribute.</span></span> <span data-ttu-id="b4d03-213">Dans le code suivant, les attributs **ID** pour les éléments **Button**, **span**et **TD** (définition de cellule de table) que les fonctions jQuery utilisent sont affichés en gras.</span><span class="sxs-lookup"><span data-stu-id="b4d03-213">In the following code, the **id** attributes for the **button**, **span**, and **td** (table cell definition) elements that jQuery functions use are shown in bold font.</span></span>

   <span data-ttu-id="b4d03-214">Le code HTML suivant ajoute une image graphique, pouvant être un logo d’entreprise.</span><span class="sxs-lookup"><span data-stu-id="b4d03-214">The following HTML adds a graphic image, which could be a company logo.</span></span> <span data-ttu-id="b4d03-215">Vous pouvez utiliser un logo de votre choix ou copier le fichier NewLogo. png à partir du téléchargement du kit de développement logiciel (SDK) de Project 2013, puis utiliser l' `HelloProjectODataWeb\Images` Explorateur de **solutions** pour ajouter le fichier au dossier.</span><span class="sxs-lookup"><span data-stu-id="b4d03-215">You can use a logo of your choice, or copy the NewLogo.png file from the Project 2013 SDK download, and then use **Solution Explorer** to add the file to the `HelloProjectODataWeb\Images` folder.</span></span>

    ```HTML
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
                <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
            <table class="infoTable" aria-readonly="True" style="width: 100%;">
                <tr>
                    <td class="heading_leftCol"></td>
                    <td class="heading_midCol"><strong>Average</strong></td>
                    <td class="heading_rightCol"><strong>Current</strong></td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Work</strong></td>
                    <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project % Complete</strong></td>
                    <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
                </tr>
            </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    ```

## <a name="creating-the-javascript-code-for-the-add-in"></a><span data-ttu-id="b4d03-216">Création du code JavaScript pour le complément</span><span class="sxs-lookup"><span data-stu-id="b4d03-216">Creating the JavaScript code for the add-in</span></span>

<span data-ttu-id="b4d03-217">Le modèle pour un complément de volet de tâches de projet inclut le code d’initialisation par défaut qui est conçu pour illustrer les actions Get et set de base des données dans un document pour un complément Office 2013 standard.</span><span class="sxs-lookup"><span data-stu-id="b4d03-217">The template for a Project task pane add-in includes default initialization code that is designed to demonstrate basic get and set actions for data in a document for a typical Office 2013 add-in.</span></span> <span data-ttu-id="b4d03-218">Étant donné que le projet 2013 ne prend pas en charge les actions qui écrivent **HelloProjectOData** dans le projet actif et que le complément `getSelectedDataAsync` HelloProjectOData n’utilise pas la méthode, vous pouvez `Office.initialize` supprimer le script dans la `setData` fonction et `getData` supprimer la fonction et la fonction dans le fichier HelloProjectOData. js par défaut.</span><span class="sxs-lookup"><span data-stu-id="b4d03-218">Because Project 2013 does not support actions that write to the active project, and the **HelloProjectOData** add-in does not use the `getSelectedDataAsync` method, you can delete the script within the `Office.initialize` function, and delete the `setData` function and `getData` function in the default HelloProjectOData.js file.</span></span>

<span data-ttu-id="b4d03-219">JavaScript comprend des constantes globales pour la requête REST et des variables globales qui sont utilisées dans plusieurs fonctions.</span><span class="sxs-lookup"><span data-stu-id="b4d03-219">The JavaScript includes global constants for the REST query and global variables that are used in several functions.</span></span> <span data-ttu-id="b4d03-220">Le bouton **obtenir le point de terminaison ProjectData** appelle la `setOdataUrl` fonction, qui initialise les variables globales et détermine si le projet est connecté à Project Web App.</span><span class="sxs-lookup"><span data-stu-id="b4d03-220">The **Get ProjectData Endpoint** button calls the `setOdataUrl` function, which initializes the global variables and determines whether Project is connected with Project Web App.</span></span>

<span data-ttu-id="b4d03-221">Le reste du fichier HelloProjectOData. js comprend deux fonctions : la `retrieveOData` fonction est appelée lorsque l’utilisateur sélectionne **comparer tous les projets**; la `parseODataResult` fonction calcule les moyennes, puis remplit le tableau de comparaison avec les valeurs mises en forme pour la couleur et les unités.</span><span class="sxs-lookup"><span data-stu-id="b4d03-221">The remainder of the HelloProjectOData.js file includes two functions: the `retrieveOData` function is called when the user selects **Compare All Projects**; and the `parseODataResult` function calculates averages and then populates the comparison table with values that are formatted for color and units.</span></span>

### <a name="procedure-5-to-create-the-javascript-code"></a><span data-ttu-id="b4d03-222">Procédure 5.</span><span class="sxs-lookup"><span data-stu-id="b4d03-222">Procedure 5.</span></span> <span data-ttu-id="b4d03-223">Pour créer du code JavaScript</span><span class="sxs-lookup"><span data-stu-id="b4d03-223">To create the JavaScript code</span></span>

1. <span data-ttu-id="b4d03-224">Supprimez tout le code dans le fichier HelloProjectOData. js par défaut, puis ajoutez les variables `**`globales et la fonction Office. Initialize.</span><span class="sxs-lookup"><span data-stu-id="b4d03-224">Delete all code in the default HelloProjectOData.js file, and then add the global variables and `**`Office.initialize\` function.</span></span> <span data-ttu-id="b4d03-225">Les noms de variable qui sont tous des majuscules signifient qu’il s’agit de constantes ; elles sont ensuite utilisées avec la variable **_pwa** pour créer la requête Rest dans cet exemple.</span><span class="sxs-lookup"><span data-stu-id="b4d03-225">Variable names that are all capitals imply that they are constants; they are later used with the **_pwa** variable to create the REST query in this example.</span></span>

    ```js
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    ```

2. <span data-ttu-id="b4d03-226">Ajouter `setOdataUrl` des fonctions connexes.</span><span class="sxs-lookup"><span data-stu-id="b4d03-226">Add `setOdataUrl` and related functions.</span></span> <span data-ttu-id="b4d03-227">La `setOdataUrl` fonction appelle `getProjectGuid` et `getDocumentUrl` initialise les variables globales.</span><span class="sxs-lookup"><span data-stu-id="b4d03-227">The `setOdataUrl` function calls `getProjectGuid` and `getDocumentUrl` to initialize the global variables.</span></span> <span data-ttu-id="b4d03-228">Dans la [méthode getProjectFieldAsync](/javascript/api/office/office.document), la fonction anonyme pour le paramètre _callback_ active le bouton **comparer tous les projets** à l' `removeAttr` aide de la méthode dans la bibliothèque jQuery, puis affiche l’URL du service **ProjectData** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-228">In the [getProjectFieldAsync method](/javascript/api/office/office.document), the anonymous function for the  _callback_ parameter enables the **Compare All Projects** button by using the `removeAttr` method in the jQuery library, and then displays the URL of the **ProjectData** service.</span></span> <span data-ttu-id="b4d03-229">Si Project n’est pas connecté à Project Web App, la fonction génère une erreur, ce qui entraîne l’affichage d’un message d’erreur dans une fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="b4d03-229">If Project is not connected with Project Web App, the function throws an error, which displays a pop-up error message.</span></span> <span data-ttu-id="b4d03-230">Le fichier SurfaceErrors. js inclut la `throwError` méthode.</span><span class="sxs-lookup"><span data-stu-id="b4d03-230">The SurfaceErrors.js file includes the `throwError` method.</span></span>

   > [!NOTE]
   > <span data-ttu-id="b4d03-231">Si vous exécutez Visual Studio sur l’ordinateur Project Server, utilisez le débogage **F5**, supprimez le commentaire de code après la ligne qui initialise la variable globale **_pwa**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-231">If you run Visual Studio on the Project Server computer, to use **F5** debugging, uncomment the code after the line that initializes the **_pwa** global variable.</span></span> <span data-ttu-id="b4d03-232">Pour activer l’utilisation de `ajax` la méthode jQuery lors du débogage sur l’ordinateur Project Server, vous devez `localhost` définir la valeur de l’URL PWA. Si vous exécutez Visual Studio sur un ordinateur distant, l' `localhost` URL n’est pas obligatoire.</span><span class="sxs-lookup"><span data-stu-id="b4d03-232">To enable using the jQuery `ajax` method when debugging on the Project Server computer, you must set the `localhost` value for the PWA URL.If you run Visual Studio on a remote computer, the  `localhost` URL is not required.</span></span> <span data-ttu-id="b4d03-233">Before you deploy the add-in, comment out that code.</span><span class="sxs-lookup"><span data-stu-id="b4d03-233">Before you deploy the add-in, comment out that code.</span></span>

    ```js
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);

                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;

                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    ```

3. <span data-ttu-id="b4d03-234">Ajoutez la `retrieveOData` fonction, qui concatène les valeurs de la requête REST, puis appelle la `ajax` fonction dans jQuery pour obtenir les données demandées à partir du service **ProjectData** .</span><span class="sxs-lookup"><span data-stu-id="b4d03-234">Add the `retrieveOData` function, which concatenates values for the REST query and then calls the `ajax` function in jQuery to get the requested data from the **ProjectData** service.</span></span> <span data-ttu-id="b4d03-235">La variable **support. cors** active le partage de ressources entre origines (cors) avec `ajax` la fonction.</span><span class="sxs-lookup"><span data-stu-id="b4d03-235">The **support.cors** variable enables cross-origin resource sharing (CORS) with the `ajax` function.</span></span> <span data-ttu-id="b4d03-236">Si l’instruction **support. cors** est manquante ou est définie sur **false**, la `ajax` fonction **ne renvoie aucune** erreur de transport.</span><span class="sxs-lookup"><span data-stu-id="b4d03-236">If the **support.cors** statement is missing or is set to **false**, the `ajax` function returns a **No transport** error.</span></span>

   > [!NOTE]
   > <span data-ttu-id="b4d03-p142">Le code suivant fonctionne avec une installation locale de Project Server 2013. Pour Project sur le web, vous pouvez utiliser OAuth pour l’authentification basée sur le jeton. Pour plus d’informations, voir [Résolutions des limites de stratégie d’origine identique dans les compléments Office](../develop/addressing-same-origin-policy-limitations.md).</span><span class="sxs-lookup"><span data-stu-id="b4d03-p142">The following code works with an on-premises installation of Project Server 2013. For Project on the web, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

   <span data-ttu-id="b4d03-240">Dans l' `ajax` appel, vous pouvez utiliser soit le paramètre _headers_ , soit le paramètre _beforeSend_ .</span><span class="sxs-lookup"><span data-stu-id="b4d03-240">In the `ajax` call, you can use either the _headers_ parameter or the _beforeSend_ parameter.</span></span> <span data-ttu-id="b4d03-241">Le paramètre _Complete_ est une fonction anonyme de sorte qu’elle se trouve dans la même étendue que `retrieveOData`les variables dans.</span><span class="sxs-lookup"><span data-stu-id="b4d03-241">The _complete_ parameter is an anonymous function so that it is in the same scope as the variables in `retrieveOData`.</span></span> <span data-ttu-id="b4d03-242">La fonction pour le paramètre _Complete_ affiche les résultats `odataText` dans le contrôle et appelle `parseODataResult` également la méthode pour analyser et afficher la réponse JSON.</span><span class="sxs-lookup"><span data-stu-id="b4d03-242">The function for the  _complete_ parameter displays results in the `odataText` control and also calls the `parseODataResult` method to parse and display the JSON response.</span></span> <span data-ttu-id="b4d03-243">Le paramètre _Error_ spécifie la `getProjectDataErrorHandler` fonction nommée, qui écrit un message d’erreur `odataText` dans le contrôle et utilise `throwError` également la méthode pour afficher un message contextuel.</span><span class="sxs-lookup"><span data-stu-id="b4d03-243">The _error_ parameter specifies the named `getProjectDataErrorHandler` function, which writes an error message to the `odataText` control and also uses the `throwError` method to display a pop-up message.</span></span>

    ```js
    // Functions to get and parse the Project Server reporting data./

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project on the web.
        $.support.cors = true;

        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;

                // xhr.responseText is the result from an XmlHttpRequest, which
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);

                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);

                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }

    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
        throwError(errorCode, errorMessage);
    }
    ```

4. <span data-ttu-id="b4d03-244">Ajoutez la `parseODataResult` méthode, qui désérialise et traite la réponse JSON du service OData.</span><span class="sxs-lookup"><span data-stu-id="b4d03-244">Add the `parseODataResult` method, which deserializes and processes the JSON response from the OData service.</span></span> <span data-ttu-id="b4d03-245">La `parseODataResult` méthode calcule les valeurs moyennes des données de coût et de travail avec une précision d’une ou deux décimales, met en forme les valeurs avec la couleur correcte **$** et ajoute une unité **%**(, **HR**ou), puis affiche les valeurs dans les cellules de tableau spécifiées.</span><span class="sxs-lookup"><span data-stu-id="b4d03-245">The `parseODataResult` method calculates average values of the cost and work data to an accuracy of one or two decimal places, formats values with the correct color and adds a unit ( **$**, **hrs**, or **%**), and then displays the values in specified table cells.</span></span>

   <span data-ttu-id="b4d03-246">Si le GUID du projet actif correspond à la `ProjectId` valeur, la `myProjectIndex` variable est définie sur l’index du projet.</span><span class="sxs-lookup"><span data-stu-id="b4d03-246">If the GUID of the active project matches the `ProjectId` value, the `myProjectIndex` variable is set to the project index.</span></span> <span data-ttu-id="b4d03-247">Si `myProjectIndex` indique que le projet actif est publié sur Project Server, `parseODataResult` la méthode met en forme et affiche les données de coût et de travail pour ce projet.</span><span class="sxs-lookup"><span data-stu-id="b4d03-247">If `myProjectIndex` indicates the active project is published on Project Server, the `parseODataResult` method formats and displays cost and work data for that project.</span></span> <span data-ttu-id="b4d03-248">Si le projet actif n’est pas publié, les valeurs pour le projet actif sont sous la forme **N/A** (en bleu).</span><span class="sxs-lookup"><span data-stu-id="b4d03-248">If the active project is not published, values for the active project are displayed as a blue **NA**.</span></span>

    ```js
    // Calculate the average values of actual cost, cost, work, and percent complete
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;

        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

        // Display averages in the table, with the correct units.
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";

        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
            Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }

            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }

            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }

            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"

            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"

            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"

            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
    ```

## <a name="testing-the-helloprojectodata-add-in"></a><span data-ttu-id="b4d03-249">Test du complément HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="b4d03-249">Testing the HelloProjectOData add-in</span></span>

<span data-ttu-id="b4d03-250">Pour tester et déboguer le complément **HelloProjectOData** avec Visual Studio 2015, Project Professionnel 2013 doit être installé sur l’ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="b4d03-250">To test and debug the **HelloProjectOData** add-in with Visual Studio 2015, Project Professional 2013 must be installed on the development computer.</span></span> <span data-ttu-id="b4d03-251">Pour permettre différents scénarios de test, assurez-vous que vous pouvez choisir si Project ouvre les fichiers sur l’ordinateur local ou s’il se connecte à Project Web App.</span><span class="sxs-lookup"><span data-stu-id="b4d03-251">To enable different test scenarios, ensure that you can choose whether Project opens for files on the local computer or connects with Project Web App.</span></span> <span data-ttu-id="b4d03-252">Par exemple, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="b4d03-252">For example, do the following steps:</span></span>

1. <span data-ttu-id="b4d03-253">Sous l’onglet **FICHIER** du ruban, choisissez l’onglet **Informations** en mode Backstage, puis choisissez **Gérer les comptes**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-253">On the **FILE** tab on the ribbon, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.</span></span>

2. <span data-ttu-id="b4d03-254">Dans la boîte de dialogue **comptes Project Web App** , la liste **comptes disponibles** peut avoir plusieurs comptes Project Web App en plus du compte de l' **ordinateur** local.</span><span class="sxs-lookup"><span data-stu-id="b4d03-254">In the **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account.</span></span> <span data-ttu-id="b4d03-255">Dans la section **Lors du démarrage**, sélectionnez  **Choisir un compte**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-255">In the **When starting** section, select **Choose an account**.</span></span>

3. <span data-ttu-id="b4d03-256">Fermez Project afin que Visual Studio puisse le démarrer pour le débogage du complément.</span><span class="sxs-lookup"><span data-stu-id="b4d03-256">Close Project so that Visual Studio can start it for debugging the add-in.</span></span>

<span data-ttu-id="b4d03-257">Voici les tests de base préconisés :</span><span class="sxs-lookup"><span data-stu-id="b4d03-257">Basic tests should include the following:</span></span>

- <span data-ttu-id="b4d03-258">Exécutez le complément à partir de Visual Studio, puis ouvrez un projet publié à partir de Project Web App, qui contient des données de coût et de travail.</span><span class="sxs-lookup"><span data-stu-id="b4d03-258">Run the add-in from Visual Studio, and then open a published project from Project Web App that contains cost and work data.</span></span> <span data-ttu-id="b4d03-259">Vérifiez que le complément affiche le point de terminaison **ProjectData** et affiche correctement les données de coût et de travail dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="b4d03-259">Verify that the add-in displays the **ProjectData** endpoint and correctly displays the cost and work data in the table.</span></span> <span data-ttu-id="b4d03-260">Vous pouvez utiliser la sortie du contrôle **odataText** pour vérifier la requête REST et d’autres informations.</span><span class="sxs-lookup"><span data-stu-id="b4d03-260">You can use the output in the **odataText** control to check the REST query and other information.</span></span>

- <span data-ttu-id="b4d03-261">Réexécutez le complément pour choisir le profil de l’ordinateur local dans la boîte de dialogue **Connexion** quand Project démarre.</span><span class="sxs-lookup"><span data-stu-id="b4d03-261">Run the add-in again, where you choose the local computer profile in the **Login** dialog box when Project starts.</span></span> <span data-ttu-id="b4d03-262">Ouvrez un fichier .mpp local, puis testez le complément.</span><span class="sxs-lookup"><span data-stu-id="b4d03-262">Open a local .mpp file, and then test the add-in.</span></span> <span data-ttu-id="b4d03-263">Vérifiez que le complément affiche un message d’erreur quand vous essayez d’obtenir le point de terminaison **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-263">Verify that the add-in displays an error message when you try to get the **ProjectData** endpoint.</span></span>

- <span data-ttu-id="b4d03-264">Réexécutez le complément pour créer un projet qui comporte des tâches avec des données de coût et de travail.</span><span class="sxs-lookup"><span data-stu-id="b4d03-264">Run the add-in again, where you create a project that has tasks with cost and work data.</span></span> <span data-ttu-id="b4d03-265">Vous pouvez enregistrer le projet dans Project Web App mais ne le publiez pas.</span><span class="sxs-lookup"><span data-stu-id="b4d03-265">You can save the project to Project Web App, but don't publish it.</span></span> <span data-ttu-id="b4d03-266">Vérifiez que le complément affiche les données de Project Server et **NA** pour le projet actuel.</span><span class="sxs-lookup"><span data-stu-id="b4d03-266">Verify that the add-in displays data from Project Server, but **NA** for the current project.</span></span>

### <a name="procedure-6-to-test-the-add-in"></a><span data-ttu-id="b4d03-p151">Procédure 6. Pour tester le complément</span><span class="sxs-lookup"><span data-stu-id="b4d03-p151">Procedure 6. To test the add-in</span></span>

1. <span data-ttu-id="b4d03-p152">Exécutez Project Professionnel 2013, connectez-vous à Project Web App, puis créez un projet de test. Affectez des tâches aux ressources locales ou à des ressources d’entreprise, définissez diverses valeurs de pourcentage achevé pour certaines tâches, puis publiez le projet. Quittez Project, ce qui permet à Visual Studio de démarrer Project pour le débogage du complément.</span><span class="sxs-lookup"><span data-stu-id="b4d03-p152">Run Project Professional 2013, connect with Project Web App, and then create a test project. Assign tasks to local resources or to enterprise resources, set various values of percent complete on some tasks, and then publish the project. Quit Project, which enables Visual Studio to start Project for debugging the add-in.</span></span>

2. <span data-ttu-id="b4d03-272">Dans Visual Studio, appuyez sur **F5**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-272">In Visual Studio, press **F5**.</span></span> <span data-ttu-id="b4d03-273">Connectez-vous à Project Web App, puis ouvrez le projet que vous avez créé à l’étape précédente.</span><span class="sxs-lookup"><span data-stu-id="b4d03-273">Log on to Project Web App, and then open the project that you created in the previous step.</span></span> <span data-ttu-id="b4d03-274">Vous pouvez ouvrir le projet en mode lecture seule ou en mode d’édition.</span><span class="sxs-lookup"><span data-stu-id="b4d03-274">You can open the project in read-only mode or in edit mode.</span></span>

3. <span data-ttu-id="b4d03-275">Sous l’onglet **projet** du ruban, dans la liste déroulante **Compléments Office** , sélectionnez **Hello ProjectData** (voir figure 5).</span><span class="sxs-lookup"><span data-stu-id="b4d03-275">On the **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData** (see Figure 5).</span></span> <span data-ttu-id="b4d03-276">Le bouton **Comparer tous les projets** devrait être désactivé.</span><span class="sxs-lookup"><span data-stu-id="b4d03-276">The **Compare All Projects** button should be disabled.</span></span>

    <span data-ttu-id="b4d03-277">*Figure 5. Démarrage du complément HelloProjectOData*</span><span class="sxs-lookup"><span data-stu-id="b4d03-277">*Figure 5. Starting the HelloProjectOData add-in*</span></span>

    ![Test de l’application HelloProjectOData](../images/pj15-hello-project-data-test-the-app.png)

4. <span data-ttu-id="b4d03-279">Dans le volet Office **Hello ProjectData**, sélectionnez **Obtenir le point de terminaison ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-279">In the **Hello ProjectData** task pane, select **Get ProjectData Endpoint**.</span></span> <span data-ttu-id="b4d03-280">La ligne **projectDataEndPoint** doit afficher l’URL du service **ProjectData** , et le bouton **comparer tous les projets** doit être activé (voir figure 6).</span><span class="sxs-lookup"><span data-stu-id="b4d03-280">The **projectDataEndPoint** line should show the URL of the **ProjectData** service, and the **Compare All Projects** button should be enabled (see Figure 6).</span></span>

5. <span data-ttu-id="b4d03-281">Sélectionnez **comparer tous les projets**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-281">Select **Compare All Projects**.</span></span> <span data-ttu-id="b4d03-282">Le complément peut s’arrêter pendant qu’il récupère des données à partir du service **ProjectData** , puis afficher les valeurs moyennes et actuelles mises en forme dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="b4d03-282">The add-in may pause while it retrieves data from the **ProjectData** service, and then it should display the formatted average and current values in the table.</span></span>

    <span data-ttu-id="b4d03-283">*Figure 6. Affichage des résultats de la requête REST*</span><span class="sxs-lookup"><span data-stu-id="b4d03-283">*Figure 6. Viewing results of the REST query*</span></span>

    ![Affichage des résultats de la requête REST](../images/pj15-hello-project-data-rest-results.png)

6. <span data-ttu-id="b4d03-285">Examinez la sortie dans la zone de texte.</span><span class="sxs-lookup"><span data-stu-id="b4d03-285">Examine output in the text box.</span></span> <span data-ttu-id="b4d03-286">Il doit afficher le chemin d’accès au document, la requête REST, les informations d’État et l’interface JSON résultant des appels vers **Ajax** et **parseODataResult**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-286">It should show the document path, REST query, status information, and JSON results from the calls to **ajax** and **parseODataResult**.</span></span> <span data-ttu-id="b4d03-287">La sortie aide à comprendre, créer et déboguer du code `parseODataResult` dans la méthode `projCost += Number(res.d.results[i].ProjectCost);`telle que.</span><span class="sxs-lookup"><span data-stu-id="b4d03-287">The output helps to understand, create, and debug code in the `parseODataResult` method such as `projCost += Number(res.d.results[i].ProjectCost);`.</span></span>

    <span data-ttu-id="b4d03-288">Voici un exemple de sortie avec des sauts de ligne et des espaces ajoutés au texte pour plus de clarté, pour trois projets dans une instance de Project Web App :</span><span class="sxs-lookup"><span data-stu-id="b4d03-288">Following is an example of the output with line breaks and spaces added to the text for clarity, for three projects in a Project Web App instance:</span></span>

    ```json
    Document path: <>\WinProj test1

    REST query:
    http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
        &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost

    textStatus: success
    ContentType: application/json;odata=verbose;charset=utf-8
    Status: 200

    ResponseText:
    {"d":{"results":[
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "type":"ReportingData.Project"},
        "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
        "ProjectActualCost":"0.000000",
        "ProjectCost":"0.000000",
        "ProjectName":"Task list created in PWA",
        "ProjectPercentCompleted":0,
        "ProjectWork":"16.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
        "ProjectActualCost":"700.000000",
        "ProjectCost":"2400.000000",
        "ProjectName":"WinProj test 2",
        "ProjectPercentCompleted":29,
        "ProjectWork":"48.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
        "ProjectActualCost":"1900.000000",
        "ProjectCost":"5200.000000",
        "ProjectName":"WinProj test1",
        "ProjectPercentCompleted":37,
        "ProjectWork":"104.000000"}
    ]}}
    ```

7. <span data-ttu-id="b4d03-289">Arrêtez le débogage (appuyez sur **Maj + F5**), puis appuyez à nouveau sur **F5** pour exécuter une nouvelle instance de Project.</span><span class="sxs-lookup"><span data-stu-id="b4d03-289">Stop debugging (press **Shift + F5**), and then press **F5** again to run a new instance of Project.</span></span> <span data-ttu-id="b4d03-290">Dans la boîte de dialogue **Connexion**, choisissez le profil  **Ordinateur** local et non Project Web App.</span><span class="sxs-lookup"><span data-stu-id="b4d03-290">In the **Login** dialog box, choose the local **Computer** profile, not Project Web App.</span></span> <span data-ttu-id="b4d03-291">Créez ou ouvrez un fichier Project .mpp local, ouvrez le volet Office **Hello ProjectData**, puis cliquez sur  **Obtenir le point de terminaison ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-291">Create or open a local project .mpp file, open the **Hello ProjectData** task pane, and then select **Get ProjectData Endpoint**.</span></span> <span data-ttu-id="b4d03-292">Le complément doit indiquer une absence de **connexion !**</span><span class="sxs-lookup"><span data-stu-id="b4d03-292">The add-in should show a **No connection!**</span></span> <span data-ttu-id="b4d03-293">erreur (voir figure 7) et le bouton **comparer tous les projets** doit rester désactivé.</span><span class="sxs-lookup"><span data-stu-id="b4d03-293">error (see Figure 7), and the **Compare All Projects** button should remain disabled.</span></span>

   <span data-ttu-id="b4d03-294">*Figure 7. Utilisation du complément sans connexion à Project Web App*</span><span class="sxs-lookup"><span data-stu-id="b4d03-294">*Figure 7. Using the add-in without a Project web app connection*</span></span>

   ![Utilisation de l’application sans connexion à Project Web App](../images/pj15-hello-project-data-no-connection.png)

8. <span data-ttu-id="b4d03-296">Arrêtez le débogage, puis appuyez à nouveau sur **F5**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-296">Stop debugging, and then press **F5** again.</span></span> <span data-ttu-id="b4d03-297">Connectez-vous à Project Web App, puis créez un projet qui contient des données de coût et de travail.</span><span class="sxs-lookup"><span data-stu-id="b4d03-297">Log on to Project Web App, and then create a project that contains cost and work data.</span></span> <span data-ttu-id="b4d03-298">Vous pouvez enregistrer le projet mais pas le publier.</span><span class="sxs-lookup"><span data-stu-id="b4d03-298">You can save the project, but don't publish it.</span></span>

   <span data-ttu-id="b4d03-299">Dans le volet Office **Hello ProjectData** , lorsque vous sélectionnez **comparer tous les projets**, vous devez voir une **na** bleue pour les champs de la colonne **actuelle** (voir figure 8).</span><span class="sxs-lookup"><span data-stu-id="b4d03-299">In the **Hello ProjectData** task pane, when you select **Compare All Projects**, you should see a blue **NA** for fields in the **Current** column (see Figure 8).</span></span>

   <span data-ttu-id="b4d03-300">*Figure 8. Comparaison d’un projet non publié à d’autres projets*</span><span class="sxs-lookup"><span data-stu-id="b4d03-300">*Figure 8. Comparing an unpublished project with other projects*</span></span>

   ![Comparaison d’un projet non publié à d’autres](../images/pj15-hello-project-data-not-published.png)

<span data-ttu-id="b4d03-p160">Même si votre complément fonctionne correctement dans les tests précédents, il existe d’autres tests à exécuter. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="b4d03-p160">Even if your add-in is working correctly in the previous tests, there are other tests that should be run. For example:</span></span>

- <span data-ttu-id="b4d03-304">À partir de Project Web App, ouvrez un projet qui ne dispose pas de données de coût ou de travail pour les tâches.</span><span class="sxs-lookup"><span data-stu-id="b4d03-304">Open a project from Project Web App that has no cost or work data for the tasks.</span></span> <span data-ttu-id="b4d03-305">Vous devriez voir des valeurs égales à zéro dans les champs de la colonne **Actif**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-305">You should see values of zero in the fields in the **Current** column.</span></span>

- <span data-ttu-id="b4d03-306">Testez un projet qui n’a pas de tâches.</span><span class="sxs-lookup"><span data-stu-id="b4d03-306">Test a project that has no tasks.</span></span>

- <span data-ttu-id="b4d03-p162">Si vous modifiez le complément et que vous le publiez, vous devez réexécuter des tests similaires avec le complément publié. Pour d’autres considérations, voir [Étapes suivantes](#next-steps).</span><span class="sxs-lookup"><span data-stu-id="b4d03-p162">If you modify the add-in and publish it, you should run similar tests again with the published add-in. For other considerations, see [Next steps](#next-steps).</span></span>

> [!NOTE]
> <span data-ttu-id="b4d03-309">Il existe des limites à la quantité de données qui peuvent être renvoyées dans une requête du service **ProjectData** ; la quantité de données varie en fonction de l’entité.</span><span class="sxs-lookup"><span data-stu-id="b4d03-309">There are limits to the amount of data that can be returned in one query of the **ProjectData** service; the amount of data varies by entity.</span></span> <span data-ttu-id="b4d03-310">Par exemple, le `Projects` jeu d’entités a une limite par défaut de 100 projets par requête, `Risks` mais le jeu d’entités a une limite par défaut de 200.</span><span class="sxs-lookup"><span data-stu-id="b4d03-310">For example, the `Projects` entity set has a default limit of 100 projects per query, but the `Risks` entity set has a default limit of 200.</span></span> <span data-ttu-id="b4d03-311">For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects.</span><span class="sxs-lookup"><span data-stu-id="b4d03-311">For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects.</span></span> <span data-ttu-id="b4d03-312">For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="b4d03-312">For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

## <a name="example-code-for-the-helloprojectodata-add-in"></a><span data-ttu-id="b4d03-313">Exemple de code pour le complément HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="b4d03-313">Example code for the HelloProjectOData add-in</span></span>

### <a name="helloprojectodatahtml-file"></a><span data-ttu-id="b4d03-314">Fichier HelloProjectOData.html</span><span class="sxs-lookup"><span data-stu-id="b4d03-314">HelloProjectOData.html file</span></span>

<span data-ttu-id="b4d03-315">Le code suivant se trouve dans le fichier `Pages\HelloProjectOData.html` du projet **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-315">The following code is in the `Pages\HelloProjectOData.html` file of the **HelloProjectODataWeb** project.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files -->
        <script src="../Scripts/HelloProjectOData.js"></script>
        <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br />
            <br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
            </tr>
        </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
</html>
```

### <a name="helloprojectodatajs-file"></a><span data-ttu-id="b4d03-316">Fichier HelloProjectOData.js</span><span class="sxs-lookup"><span data-stu-id="b4d03-316">HelloProjectOData.js file</span></span>

<span data-ttu-id="b4d03-317">Le code suivant se trouve dans le fichier `Scripts\Office\HelloProjectOData.js` du projet **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-317">The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project.</span></span>

```js
/* File: HelloProjectOData.js
* JavaScript functions for the HelloProjectOData example task pane app.
* October 2, 2012
*/

var PROJDATA = "/_api/ProjectData";
var PROJQUERY = "/Projects?";
var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
var _pwa;           // URL of Project Web App.
var _projectUid;    // GUID of the active project.
var _docUrl;        // Path of the project document.
var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
    });
}

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project is not connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                // If you debug with Visual Studio on a local Project Server computer,
                // uncomment the following lines to use the localhost URL.
                //var localhost = location.host.split(":", 1);
                //var pwaStartPosition = _pwa.lastIndexOf("/");
                //var pwaLength = _pwa.length - pwaStartPosition;
                //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                //_pwa = location.protocol + "//" + localhost + pwaName;

                if (_pwa.substring(0, 4) == "http") {
                    _odataUrl = _pwa + PROJDATA;
                    $("#compareProjects").removeAttr("disabled");
                    getProjectGuid();
                }
                else {
                    _odataUrl = "No connection!";
                    throwError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                $("#projectDataEndPoint").text(_odataUrl);
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project web app, which is in the form <>\ProjectName .
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

//  Functions to get and parse the Project Server reporting data./

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project on the web.
    $.support.cors = true;

    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json",
        data: "",      // Empty string for the optional data.
        //headers: { "Accept": accept },
        beforeSend: function (xhr) {
            xhr.setRequestHeader("ACCEPT", accept);
        },
        complete: function (xhr, textStatus) {
            // Create a message to display in the text box.
            var message = "\r\ntextStatus: " + textStatus +
                "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                "\r\nStatus: " + xhr.status +
                "\r\nResponseText:\r\n" + xhr.responseText;

            // xhr.responseText is the result from an XmlHttpRequest, which 
            // contains the JSON response from the OData service.
            parseODataResult(xhr.responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            $("#odataText").text(_docUrl);
            $("#odataText").append("\r\nREST query:\r\n" + restUrl);
            $("#odataText").append(message);

            if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
            }
        },
        error: getProjectDataErrorHandler
    });
}

function getProjectDataErrorHandler(data, errorCode, errorMessage) {
    $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
    throwError(errorCode, errorMessage);
}

// Calculate the average values of actual cost, cost, work, and percent complete
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
    var len = res.d.results.length;
    var projActualCost = 0;
    var projCost = 0;
    var projWork = 0;
    var projPercentCompleted = 0;
    var myProjectIndex = -1;

    for (i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,  
        // then store the project index.
        if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);

    }
    var avgProjCost = projCost / len;
    var avgProjWork = projWork / len;
    var avgProjActualCost = projActualCost / len;
    var avgProjPercentCompleted = projPercentCompleted / len;

    // Round off cost to two decimal places, and round off other values to one decimal place.
    avgProjCost = avgProjCost.toFixed(2);
    avgProjWork = avgProjWork.toFixed(1);
    avgProjActualCost = avgProjActualCost.toFixed(2);
    avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

    // Display averages in the table, with the correct units. 
    document.getElementById("AverageProjectCost").innerHTML = "$"
        + avgProjCost;
    document.getElementById("AverageProjectActualCost").innerHTML
        = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").innerHTML
        = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").innerHTML
        = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex != -1) {

        var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
        var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
        var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
        var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

        myProjCost = myProjCost.toFixed(2);
        myProjWork = myProjWork.toFixed(1);
        myProjActualCost = myProjActualCost.toFixed(2);
        myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

        document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

        if (Number(myProjCost) <= Number(avgProjCost)) {
            document.getElementById("CurrentProjectCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectCost").style.color = "red"
        }

        document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

        if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
            document.getElementById("CurrentProjectActualCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectActualCost").style.color = "red"
        }

        document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

        if (Number(myProjWork) <= Number(avgProjWork)) {
            document.getElementById("CurrentProjectWork").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectWork").style.color = "green"
        }

        document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

        if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
            document.getElementById("CurrentProjectPercentComplete").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectPercentComplete").style.color = "green"
        }
    }
    else {    // The current project is not published.
        document.getElementById("CurrentProjectCost").innerHTML = "NA";
        document.getElementById("CurrentProjectCost").style.color = "blue"

        document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
        document.getElementById("CurrentProjectActualCost").style.color = "blue"

        document.getElementById("CurrentProjectWork").innerHTML = "NA";
        document.getElementById("CurrentProjectWork").style.color = "blue"

        document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
        document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
    }
}
```

### <a name="appcss-file"></a><span data-ttu-id="b4d03-318">Fichier App.css</span><span class="sxs-lookup"><span data-stu-id="b4d03-318">App.css file</span></span>

<span data-ttu-id="b4d03-319">Le code suivant se trouve dans le fichier `Content\App.css` du projet **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-319">The following code is in the `Content\App.css` file of the **HelloProjectODataWeb** project.</span></span>

```css
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/

body
{
    font-size: 11pt;
}
h1
{
    font-size: 22pt;
}
h2
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium;
    font-weight: bold;
}
.row_leftCol
{
    width: 20px;
    font-size: small;
    font-weight: bold;
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

### <a name="surfaceerrorsjs-file"></a><span data-ttu-id="b4d03-320">Fichier SurfaceErrors.js</span><span class="sxs-lookup"><span data-stu-id="b4d03-320">SurfaceErrors.js file</span></span>

<span data-ttu-id="b4d03-321">Vous pouvez copier le code du fichier SurfaceErrors.js présenté dans la section _Programmation fiable_ de la page [Créer votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span><span class="sxs-lookup"><span data-stu-id="b4d03-321">You can copy code for the SurfaceErrors.js file from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="b4d03-322">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="b4d03-322">Next steps</span></span>

<span data-ttu-id="b4d03-323">Si **HelloProjectOData** était un complément de production à vendre dans AppSource ou distribué dans un catalogue d’applications SharePoint, il serait conçu différemment.</span><span class="sxs-lookup"><span data-stu-id="b4d03-323">If **HelloProjectOData** were a production add-in to be sold in AppSource or distributed in a SharePoint app catalog, it would be designed differently.</span></span> <span data-ttu-id="b4d03-324">Par exemple, il n’y aurait pas de sortie de débogage dans une zone de texte et probablement pas de bouton permettant d’obtenir le point de terminaison de **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="b4d03-324">For example, there would be no debug output in a text box, and probably no button to get the **ProjectData** endpoint.</span></span> <span data-ttu-id="b4d03-325">Vous devez également réécrire la fonction pour `retireveOData` gérer les instances Project Web App qui contiennent plus de 100 projets.</span><span class="sxs-lookup"><span data-stu-id="b4d03-325">You would also have to rewrite the `retireveOData` function to handle Project Web App instances that have more than 100 projects.</span></span>

<span data-ttu-id="b4d03-p165">Le complément devrait contenir des contrôles d’erreurs supplémentaires, ainsi qu’une logique permettant d’identifier et d’expliquer ou d’illustrer les cas extrêmes. Par exemple, si une instance de Project Web App a 1 000 projets d’une durée moyenne de cinq jours et d’un coût moyen de 2 400 €, et que le projet actif est le seul dont la durée est supérieure à 20 jours, la comparaison des coûts et du travail est faussée. Cela pourrait être illustré avec un graphique de fréquences. Vous pouvez ajouter des options pour afficher la durée, comparer les projets de durée similaire ou comparer les projets de services identiques ou distincts. Sinon, vous pouvez également permettre à l’utilisateur d’effectuer des choix parmi une liste de champs affichés.</span><span class="sxs-lookup"><span data-stu-id="b4d03-p165">The add-in should contain additional error checks, plus logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1000 projects with an average duration of five days and average cost of $2400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. That could be shown with a frequency graph. You could add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.</span></span>

<span data-ttu-id="b4d03-331">Pour les autres requêtes du service **ProjectData** , il existe des limites à la longueur de la chaîne de requête, ce qui affecte le nombre d’étapes qu’une requête peut effectuer d’une collection parent vers un objet d’une collection enfant.</span><span class="sxs-lookup"><span data-stu-id="b4d03-331">For other queries of the **ProjectData** service, there are limits to the length of the query string, which affects the number of steps that a query can take from a parent collection to an object in a child collection.</span></span> <span data-ttu-id="b4d03-332">Par exemple, une requête en deux étapes telle que  **Projects** vers **Tasks**, puis vers un élément de tâche fonctionne, mais une requête en trois étapes telle que  **Projects** vers **Tasks** vers **Assignments**, puis vers l’élément d’affectation risque de dépasser la longueur maximale par défaut de l’URL.</span><span class="sxs-lookup"><span data-stu-id="b4d03-332">For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item may exceed the default maximum URL length.</span></span> <span data-ttu-id="b4d03-333">Pour plus d’informations, voir [Interrogation des flux OData pour les données de création de rapports Project](/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="b4d03-333">For more information, see [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

<span data-ttu-id="b4d03-334">Si vous modifiez le complément **HelloProjectOData** pour une utilisation en production, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="b4d03-334">If you modify the **HelloProjectOData** add-in for production use, do the following steps:</span></span>

- <span data-ttu-id="b4d03-335">Dans le fichier HelloProjectOData.html, pour de meilleures performances, remplacez la référence du projet local à office.js par la référence au réseau de distribution de contenu :</span><span class="sxs-lookup"><span data-stu-id="b4d03-335">In the HelloProjectOData.html file, for better performance, change the office.js reference from the local project to the CDN reference:</span></span>

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- <span data-ttu-id="b4d03-336">Réécrivez la `retrieveOData` fonction pour activer les requêtes de plus de 100 projets.</span><span class="sxs-lookup"><span data-stu-id="b4d03-336">Rewrite the `retrieveOData` function to enable queries of more than 100 projects.</span></span> <span data-ttu-id="b4d03-337">Par exemple, vous pouvez obtenir le nombre de projets avec une requête `~/ProjectData/Projects()/$count`, puis utiliser l’opérateur _$skip_ et l’opérateur _$top_ de la requête REST pour les données de projet.</span><span class="sxs-lookup"><span data-stu-id="b4d03-337">For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query, and use the _$skip_ operator and _$top_ operator in the REST query for project data.</span></span> <span data-ttu-id="b4d03-338">Exécutez plusieurs requêtes dans une boucle, puis établissez la moyenne des données de chaque requête.</span><span class="sxs-lookup"><span data-stu-id="b4d03-338">Run multiple queries in a loop, and then average the data from each query.</span></span> <span data-ttu-id="b4d03-339">Chaque requête de données de projet est de la forme :</span><span class="sxs-lookup"><span data-stu-id="b4d03-339">Each query for project data would be of the form:</span></span> 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  <span data-ttu-id="b4d03-p168">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="b4d03-p168">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span></span>

- <span data-ttu-id="b4d03-343">Pour déployer le complément, voir [Publier votre complément Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="b4d03-343">To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b4d03-344">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b4d03-344">See also</span></span>

- [<span data-ttu-id="b4d03-345">Compléments du volet Office pour Project</span><span class="sxs-lookup"><span data-stu-id="b4d03-345">Task pane add-ins for Project</span></span>](project-add-ins.md)
- [<span data-ttu-id="b4d03-346">Créer votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte</span><span class="sxs-lookup"><span data-stu-id="b4d03-346">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- <span data-ttu-id="b4d03-347">[ProjectData – Référence de service Project OData](/previous-versions/office/project-odata/jj163015(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="b4d03-347">[ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15))</span></span>
- [<span data-ttu-id="b4d03-348">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b4d03-348">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="b4d03-349">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="b4d03-349">Publish your Office Add-in</span></span>](../publish/publish.md)
