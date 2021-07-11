---
title: Ouvrez Excel à partir de votre page web et incorporez votre Office de recherche
description: Ouvrez Excel à partir de votre page web et incorporez votre Office de recherche.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 18f40b0030f4132a413a879e8b3419af49984b45
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349377"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="ad464-103">Ouvrez Excel à partir de votre page web et incorporez votre Office de recherche</span><span class="sxs-lookup"><span data-stu-id="ad464-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Image du Excel sur votre page web ouvrant un nouveau document Excel avec votre application incorporée et à l’ouverture automatique.":::

<span data-ttu-id="ad464-105">Étendez votre application web SaaS afin que vos clients peuvent ouvrir leurs données à partir d’une page web directement Microsoft Excel.</span><span class="sxs-lookup"><span data-stu-id="ad464-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="ad464-106">Un scénario courant est que les clients vont travailler avec des données dans votre application web.</span><span class="sxs-lookup"><span data-stu-id="ad464-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="ad464-107">Ensuite, ils souhaiteront copier les données dans un Excel document.</span><span class="sxs-lookup"><span data-stu-id="ad464-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="ad464-108">Par exemple, ils peuvent effectuer des analyses supplémentaires à l’aide de Excel.</span><span class="sxs-lookup"><span data-stu-id="ad464-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="ad464-109">En règle générale, le client doit exporter les données dans un fichier, tel qu’un fichier .csv, puis importer ces données dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ad464-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="ad464-110">Ils doivent également ajouter manuellement votre Office au document.</span><span class="sxs-lookup"><span data-stu-id="ad464-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="ad464-111">Réduisez le nombre d’étapes en un seul clic sur votre page web qui génère et ouvre Excel document.</span><span class="sxs-lookup"><span data-stu-id="ad464-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="ad464-112">Vous pouvez également incorporer votre Office dans le document et l’afficher à l’ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="ad464-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="ad464-113">Cela garantit que le client a toujours accès aux fonctionnalités de votre application.</span><span class="sxs-lookup"><span data-stu-id="ad464-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="ad464-114">Lorsque le document s’ouvre, les données que le client a sélectionnées et votre Office est déjà disponible pour qu’il continue de fonctionner.</span><span class="sxs-lookup"><span data-stu-id="ad464-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="ad464-115">Cet article vous présente le code et les techniques permettant d’implémenter ce scénario dans votre propre application web SaaS.</span><span class="sxs-lookup"><span data-stu-id="ad464-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="ad464-116">Créer un document Excel et incorporer un Office de document</span><span class="sxs-lookup"><span data-stu-id="ad464-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="ad464-117">Tout d’abord, nous allons apprendre à créer un document Excel à partir d’une page web et à incorporer un add-in dans le document.</span><span class="sxs-lookup"><span data-stu-id="ad464-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="ad464-118">L Office exemple de code de l’incorporation de code [](https://appsource.microsoft.com/product/office/wa104380862) [ooXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) dans un Script Lab dans un nouveau document Office document.</span><span class="sxs-lookup"><span data-stu-id="ad464-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="ad464-119">Bien que l’exemple fonctionne avec Office document, nous nous concentrerons simplement sur Excel feuilles de calcul dans cet article.</span><span class="sxs-lookup"><span data-stu-id="ad464-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="ad464-120">Utilisez les étapes suivantes pour créer et exécuter l’exemple.</span><span class="sxs-lookup"><span data-stu-id="ad464-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="ad464-121">Extrayez l’exemple de code  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip dans un dossier de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="ad464-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="ad464-122">Pour créer et exécuter l’exemple, suivez les étapes de la section **Pour utiliser le** projet du lisez-moi.</span><span class="sxs-lookup"><span data-stu-id="ad464-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="ad464-123">Lorsque vous exécutez l’exemple, il affiche une page web semblable à la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="ad464-123">When you run the sample it will display a web page similar to the following screenshot.</span></span> <span data-ttu-id="ad464-124">Utilisez la page web pour créer un document Excel qui contient les Script Lab lors de son ouverture.</span><span class="sxs-lookup"><span data-stu-id="ad464-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Capture d’écran de la page web affichée par l’exemple de laboratoire de script pour la sélection d’un fichier Excel et son incorporation.":::

### <a name="how-the-sample-works"></a><span data-ttu-id="ad464-126">Fonctionnement de l’exemple</span><span class="sxs-lookup"><span data-stu-id="ad464-126">How the sample works</span></span>

<span data-ttu-id="ad464-127">L’exemple de code utilise le SDK OOXML pour incorporer le Script Lab dans le document Excel que vous choisissez.</span><span class="sxs-lookup"><span data-stu-id="ad464-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="ad464-128">Les informations suivantes sont issues de la section [ **à propos du code**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) dans le fichier Lisez-moi.</span><span class="sxs-lookup"><span data-stu-id="ad464-128">The following information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="ad464-129">Le fichier **Home.aspx.cs**:</span><span class="sxs-lookup"><span data-stu-id="ad464-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="ad464-130">Fournit les handlers d’événements de bouton et la manipulation de base de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ad464-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="ad464-131">Utilise des techniques ASP.NET standard pour télécharger le fichier.</span><span class="sxs-lookup"><span data-stu-id="ad464-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="ad464-132">Utilise l’extension de nom de fichier du fichier téléchargé (xlsx, docx ou pptx) pour déterminer le type de fichier.</span><span class="sxs-lookup"><span data-stu-id="ad464-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="ad464-133">Cette étape doit être effectuée au départ, car le SDK Open XML possède généralement des API distinctes pour chaque type de fichier.</span><span class="sxs-lookup"><span data-stu-id="ad464-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="ad464-134">Appels dans **OOXMLHelper** pour valider le fichier et appels dans **le AddInEmbedder** pour incorporer des Script Lab dans le fichier et définir pour s’ouvrir automatiquement.</span><span class="sxs-lookup"><span data-stu-id="ad464-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="ad464-135">Le fichier **AddInEmbedder.cs**:</span><span class="sxs-lookup"><span data-stu-id="ad464-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="ad464-136">Fournit la logique métier principale, qui dans cet exemple est une méthode qui incorpore Script Lab.</span><span class="sxs-lookup"><span data-stu-id="ad464-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="ad464-137">Effectue des appels dans l’aide OOXML en fonction du type de fichier.</span><span class="sxs-lookup"><span data-stu-id="ad464-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="ad464-138">Le fichier **OOXMLHelper.cs**:</span><span class="sxs-lookup"><span data-stu-id="ad464-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="ad464-139">Fournit toutes les manipulations OOXML détaillées.</span><span class="sxs-lookup"><span data-stu-id="ad464-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="ad464-140">Utilise une technique standard pour valider le fichier Office, qui consiste simplement à appeler la **méthode Document.Open** dessus.</span><span class="sxs-lookup"><span data-stu-id="ad464-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="ad464-141">Si le fichier n’est pas valide, la méthode envoie une exception.</span><span class="sxs-lookup"><span data-stu-id="ad464-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="ad464-142">Contient principalement du code généré par les outils de productivité du SDK Open XML 2.5 qui sont disponibles sur le lien du [SDK Open XML 2.5.](/office/open-xml/open-xml-sdk)</span><span class="sxs-lookup"><span data-stu-id="ad464-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="ad464-143">La **méthode GenerateWebExtensionPart1Content** dans le fichier **OOXMLHelper.cs** définit la référence à l’ID de Script Lab dans Microsoft AppSource :</span><span class="sxs-lookup"><span data-stu-id="ad464-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="ad464-144">La **valeur StoreType** est « OMEX », alias de Microsoft AppSource.</span><span class="sxs-lookup"><span data-stu-id="ad464-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="ad464-145">La **valeur** du Store est « en-US » dans la section culture Microsoft AppSource Script Lab.</span><span class="sxs-lookup"><span data-stu-id="ad464-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="ad464-146">La **valeur d’ID** est l’ID d’actif Microsoft AppSource Script Lab.</span><span class="sxs-lookup"><span data-stu-id="ad464-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="ad464-147">Si vous souhaitez ouvrir automatiquement un add-in à partir d’un catalogue de partages de fichiers, vous utiliserez différentes valeurs :</span><span class="sxs-lookup"><span data-stu-id="ad464-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="ad464-148">La **valeur StoreType** est « FileSystem ».</span><span class="sxs-lookup"><span data-stu-id="ad464-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="ad464-149">La **valeur du Store** est l’URL du partage réseau . par exemple, « \\ \\ MyComputer \\ MySharedFolder ».</span><span class="sxs-lookup"><span data-stu-id="ad464-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="ad464-150">Il doit s’agit de l’URL exacte qui apparaît en tant qu’adresse de catalogue approuvé du partage dans Office de confiance.</span><span class="sxs-lookup"><span data-stu-id="ad464-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="ad464-151">La **valeur de l’ID** est l’ID de l’application dans le manifeste des applications.</span><span class="sxs-lookup"><span data-stu-id="ad464-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="ad464-152">Pour plus d’informations sur les autres valeurs de ces attributs, voir Ouvrir automatiquement un volet Des tâches [avec un document.](../develop/automatically-open-a-task-pane-with-a-document.md)</span><span class="sxs-lookup"><span data-stu-id="ad464-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="ad464-153">Utiliser l’interface Fluent utilisateur</span><span class="sxs-lookup"><span data-stu-id="ad464-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent Icônes d’interface utilisateur pour Word, Excel et PowerPoint.":::

<span data-ttu-id="ad464-155">Une meilleure pratique consiste à utiliser l’interface utilisateur Fluent pour aider vos utilisateurs à passer d’un produit Microsoft à un autre.</span><span class="sxs-lookup"><span data-stu-id="ad464-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="ad464-156">Vous devez toujours utiliser une icône Office pour indiquer quelle application Office sera lancée à partir de votre page web.</span><span class="sxs-lookup"><span data-stu-id="ad464-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="ad464-157">Nous allons modifier l’exemple de code pour utiliser l’icône Excel pour indiquer qu’il lance l’application Excel application.</span><span class="sxs-lookup"><span data-stu-id="ad464-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="ad464-158">Ouvrez l’exemple dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="ad464-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="ad464-159">Ouvrez la page **Home.aspx.**</span><span class="sxs-lookup"><span data-stu-id="ad464-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="ad464-160">Recherchez le code suivant qui est le bouton de téléchargement sur le formulaire.</span><span class="sxs-lookup"><span data-stu-id="ad464-160">Find following code that is the download button on the form.</span></span>

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. <span data-ttu-id="ad464-161">Remplacez le code du bouton par la balise d’image suivante.</span><span class="sxs-lookup"><span data-stu-id="ad464-161">Replace the button code with the following image tag.</span></span>

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. <span data-ttu-id="ad464-162">Appuyez **sur F5** (ou **déboguer > démarrer le débogage).**</span><span class="sxs-lookup"><span data-stu-id="ad464-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="ad464-163">L’icône s’affiche lors du chargement de la page d’accueil.</span><span class="sxs-lookup"><span data-stu-id="ad464-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="ad464-164">Pour plus d’informations, [voir Office Icônes](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) de marque sur le portail Fluent de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ad464-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="ad464-165">Télécharger le Excel document à Microsoft OneDrive</span><span class="sxs-lookup"><span data-stu-id="ad464-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="ad464-166">Nous vous recommandons de télécharger de nouveaux documents vers OneDrive si votre client utilise OneDrive.</span><span class="sxs-lookup"><span data-stu-id="ad464-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="ad464-167">Cela leur permet de trouver et d’utiliser plus facilement les documents.</span><span class="sxs-lookup"><span data-stu-id="ad464-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="ad464-168">Nous allons créer un exemple de code et voir comment vous pouvez utiliser le SDK Microsoft Graph pour télécharger un nouveau document Excel sur OneDrive.</span><span class="sxs-lookup"><span data-stu-id="ad464-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="ad464-169">Utiliser un démarrage rapide pour créer une application web Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="ad464-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="ad464-170">Suivez les étapes de création et d’ouverture d’un exemple de code de démarrage rapide qui [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) interagit avec Office services.</span><span class="sxs-lookup"><span data-stu-id="ad464-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office services.</span></span>
1. <span data-ttu-id="ad464-171">À **l’étape 1 : choisissez la langue ou la** plateforme, choisissez ASP.NET **MVC.**</span><span class="sxs-lookup"><span data-stu-id="ad464-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="ad464-172">Bien que les étapes de cette procédure utilisent l’option ASP.NET MVC, elles suivent un modèle qui s’applique à n’importe quelle langue ou plateforme.</span><span class="sxs-lookup"><span data-stu-id="ad464-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="ad464-173">À **l’étape 2 : Obtenez un ID d’application** et une secret, choisissez Obtenir **un ID d’application et un secret**.</span><span class="sxs-lookup"><span data-stu-id="ad464-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="ad464-174">Connectez-vous à Microsoft 365 compte.</span><span class="sxs-lookup"><span data-stu-id="ad464-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="ad464-175">Dans la page **Web Veuillez enregistrer votre** secret d’application, enregistrez-la dans un emplacement de fichier où vous pourrez l’extraire et l’utiliser ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="ad464-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="ad464-176">Choose **Got it, take me back to the quick start**.</span><span class="sxs-lookup"><span data-stu-id="ad464-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="ad464-177">À **l’étape 2 : l’inscription a réussi !**</span><span class="sxs-lookup"><span data-stu-id="ad464-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="ad464-178">Entrez la secret de l’application générée.</span><span class="sxs-lookup"><span data-stu-id="ad464-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="ad464-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span><span class="sxs-lookup"><span data-stu-id="ad464-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="ad464-180">Extrayez le dossier zip de téléchargement dans un dossier local.</span><span class="sxs-lookup"><span data-stu-id="ad464-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="ad464-181">Ouvrez le fichier graph-tutorial.sln dans Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="ad464-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="ad464-182">Créez et exécutez la solution et confirmez qu’elle fonctionne correctement.</span><span class="sxs-lookup"><span data-stu-id="ad464-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="ad464-183">Vous devriez être en mesure d’utiliser la page web de calendrier pour afficher votre Microsoft 365 calendrier.</span><span class="sxs-lookup"><span data-stu-id="ad464-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="ad464-184">Télécharger fichier à OneDrive</span><span class="sxs-lookup"><span data-stu-id="ad464-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="ad464-185">Ouvrez la solution **graph-tutorial.sln** Visual Studio 2019 et ouvrez **PrivateSettings.config** fichier.</span><span class="sxs-lookup"><span data-stu-id="ad464-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="ad464-186">Ajoutez une nouvelle **étendue Files.ReadWrite** à la clé   **ida:AppScopes** afin qu’elle ressemble au code suivant.</span><span class="sxs-lookup"><span data-stu-id="ad464-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code.</span></span>

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. <span data-ttu-id="ad464-187">Ouvrez **le fichier Index.cshtml.**</span><span class="sxs-lookup"><span data-stu-id="ad464-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="ad464-188">Insérez le code ActionLink suivant pour créer un bouton pour télécharger un fichier vers OneDrive.</span><span class="sxs-lookup"><span data-stu-id="ad464-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. <span data-ttu-id="ad464-189">Ouvrez **le fichier HomeController.cs.**</span><span class="sxs-lookup"><span data-stu-id="ad464-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="ad464-190">Insérez le code suivant pour gérer la demande à partir du lien d’action.</span><span class="sxs-lookup"><span data-stu-id="ad464-190">Insert the following code to handle the request from the action link.</span></span>

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. <span data-ttu-id="ad464-191">Ouvrez **le fichier GraphHelper.cs.**</span><span class="sxs-lookup"><span data-stu-id="ad464-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="ad464-192">Insérez le code suivant pour appeler l’API microsoft Graph pour créer un fichier sur OneDrive.</span><span class="sxs-lookup"><span data-stu-id="ad464-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>

    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```

1. <span data-ttu-id="ad464-193">Appuyez **sur F5** (ou **déboguer > démarrer le débogage).**</span><span class="sxs-lookup"><span data-stu-id="ad464-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="ad464-194">L’application web démarre.</span><span class="sxs-lookup"><span data-stu-id="ad464-194">The web application will start.</span></span>
1. <span data-ttu-id="ad464-195">Cliquez **ici pour vous connectez** et connectez-vous.</span><span class="sxs-lookup"><span data-stu-id="ad464-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="ad464-196">Cliquez **ici pour créer un fichier sur OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="ad464-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="ad464-197">Ouvrez un nouvel onglet de navigateur et connectez-vous à OneDrive compte.</span><span class="sxs-lookup"><span data-stu-id="ad464-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="ad464-198">Vous verrez le fichier test.txt dans le dossier racine.</span><span class="sxs-lookup"><span data-stu-id="ad464-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="ad464-199">Maintenant que vous avez appris à télécharger un fichier vers OneDrive, vous pouvez réutiliser ce code pour télécharger n’importe quel document Excel que vous créez.</span><span class="sxs-lookup"><span data-stu-id="ad464-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="ad464-200">Considérations supplémentaires pour votre solution</span><span class="sxs-lookup"><span data-stu-id="ad464-200">Additional considerations for your solution</span></span>

<span data-ttu-id="ad464-201">La solution de tout le monde est différente en termes de technologies et d’approches.</span><span class="sxs-lookup"><span data-stu-id="ad464-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="ad464-202">Les considérations suivantes vous aideront à planifier la modification de votre solution pour ouvrir des documents et incorporer votre Office de données.</span><span class="sxs-lookup"><span data-stu-id="ad464-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="ad464-203">Créer une feuille Excel feuille de calcul à partir de la page web</span><span class="sxs-lookup"><span data-stu-id="ad464-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="ad464-204">L’exemple modifie un document Excel existant.</span><span class="sxs-lookup"><span data-stu-id="ad464-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="ad464-205">Un scénario plus courant consiste à créer une feuille de calcul Excel feuille de calcul à partir de votre page web.</span><span class="sxs-lookup"><span data-stu-id="ad464-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="ad464-206">Vous trouverez des détails supplémentaires sur la création d’une feuille de calcul dans Créer un **document** de feuilles de calcul en fournissant un nom de fichier.</span><span class="sxs-lookup"><span data-stu-id="ad464-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="ad464-207">Cet article montre comment créer le fichier localement, mais vous pouvez également créer le fichier dans un flux à l’aide d’une surcharge sur la méthode SpreadsheetDocument.Create.</span><span class="sxs-lookup"><span data-stu-id="ad464-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="ad464-208">Lire les propriétés personnalisées au démarrage de votre add-in</span><span class="sxs-lookup"><span data-stu-id="ad464-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="ad464-209">L’exemple de code stocke un ID d’extrait de code dans le nouveau document Excel l’aide du SDK OOXML.</span><span class="sxs-lookup"><span data-stu-id="ad464-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="ad464-210">Script Lab lit l’ID d’extrait de code du document Excel puis affiche ce code d’extrait de code lorsqu’il s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="ad464-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="ad464-211">Vous devrez peut-être envoyer des propriétés personnalisées à votre propre add-in (par exemple, une chaîne de requête ou un jeton d’authentification temporaire).) Pour **plus d’informations** sur la lecture des propriétés personnalisées au démarrage de votre compl?ment, voir l’état et les paramètres persistants du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="ad464-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="ad464-212">Initialiser le document Excel avec des données</span><span class="sxs-lookup"><span data-stu-id="ad464-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="ad464-213">En règle générale, lorsque le client ouvre un document Excel partir de votre site web, il s’attend à ce que le document contienne des données du site web.</span><span class="sxs-lookup"><span data-stu-id="ad464-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="ad464-214">Il existe deux façons d’écrire des données dans le document.</span><span class="sxs-lookup"><span data-stu-id="ad464-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="ad464-215">**Utilisez le SDK OOXML pour écrire les données.**</span><span class="sxs-lookup"><span data-stu-id="ad464-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="ad464-216">Vous pouvez utiliser le SDK pour écrire directement des données dans le document.</span><span class="sxs-lookup"><span data-stu-id="ad464-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="ad464-217">Cette approche est utile si vous souhaitez que les données soient disponibles dès que le document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="ad464-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="ad464-218">**Passez une propriété de requête personnalisée à votre Office.**</span><span class="sxs-lookup"><span data-stu-id="ad464-218">**Pass a custom query property to your Office Add-in**.</span></span> <span data-ttu-id="ad464-219">Lorsque vous générez le document, vous incorporez une propriété personnalisée pour le Office qui contient une chaîne de requête qui récupère toutes les données requises.</span><span class="sxs-lookup"><span data-stu-id="ad464-219">When you generate the document, you embed a custom property for the Office Add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="ad464-220">Lorsque votre application s’ouvre, elle récupère la requête, l’exécute et utilise l’API JS Office pour insérer le résultat de la requête dans le document.</span><span class="sxs-lookup"><span data-stu-id="ad464-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="ad464-221">Travailler avec le SDK OOXML</span><span class="sxs-lookup"><span data-stu-id="ad464-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="ad464-222">Le SDK OOXML est basé sur .NET.</span><span class="sxs-lookup"><span data-stu-id="ad464-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="ad464-223">Si votre application web n’est pas .NET, vous devez rechercher une autre façon de travailler avec OOXML.</span><span class="sxs-lookup"><span data-stu-id="ad464-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="ad464-224">Il existe une version JavaScript du SDK OOXML disponible dans le [SDK Open XML pour JavaScript.](https://archive.codeplex.com/?p=openxmlsdkjs)</span><span class="sxs-lookup"><span data-stu-id="ad464-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="ad464-225">Vous pouvez placer le code OOXML dans une fonction Azure pour séparer le code .NET du reste de votre application web.</span><span class="sxs-lookup"><span data-stu-id="ad464-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="ad464-226">Appelez ensuite la fonction Azure (pour générer le Excel document) à partir de votre application Web.</span><span class="sxs-lookup"><span data-stu-id="ad464-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="ad464-227">Pour plus d’informations sur les fonctions Azure, voir [une présentation des fonctions Azure.](/azure/azure-functions/functions-overview)</span><span class="sxs-lookup"><span data-stu-id="ad464-227">For more information on Azure functions, see [An introduction to Azure Functions](/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="ad464-228">Utiliser l' sign-on unique</span><span class="sxs-lookup"><span data-stu-id="ad464-228">Use single sign-on</span></span>

<span data-ttu-id="ad464-229">Pour simplifier l’authentification, nous recommandons que votre application implémente l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="ad464-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="ad464-230">Pour plus d’informations, [voir Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ad464-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="ad464-231">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ad464-231">See also</span></span>

- [<span data-ttu-id="ad464-232">Bienvenue dans le SDK Open XML 2.5 pour Office</span><span class="sxs-lookup"><span data-stu-id="ad464-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="ad464-233">Ouvrir automatiquement un volet de tâches avec un document</span><span class="sxs-lookup"><span data-stu-id="ad464-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="ad464-234">Conservation de l’état et des paramètres des compléments</span><span class="sxs-lookup"><span data-stu-id="ad464-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="ad464-235">Créer un document de feuilles de calcul en fournissant un nom de fichier</span><span class="sxs-lookup"><span data-stu-id="ad464-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)