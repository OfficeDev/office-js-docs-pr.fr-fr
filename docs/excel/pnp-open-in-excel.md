---
title: Ouvrez Excel à partir de votre page web et incorporez votre Office de recherche
description: Ouvrez Excel à partir de votre page web et incorporez votre Office de recherche.
ms.date: 02/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0ac644de03c1f3a4c382dbe151c3224afffdbc81
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149102"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Ouvrez Excel à partir de votre page web et incorporez votre Office de recherche

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Image du Excel sur votre page web ouvrant un nouveau document Excel avec votre application incorporée et à l’ouverture automatique.":::

Étendez votre application web SaaS afin que vos clients peuvent ouvrir leurs données directement à partir d’une page web Microsoft Excel. Un scénario courant est que les clients vont travailler avec des données dans votre application web. Ensuite, ils souhaiteront copier les données dans un document Excel document. Par exemple, ils peuvent effectuer des analyses supplémentaires à l’aide de Excel. En règle générale, le client doit exporter les données dans un fichier, tel qu’un fichier .csv, puis importer ces données dans Excel. Ils doivent également ajouter manuellement votre Office au document.

Réduisez le nombre d’étapes en un seul clic sur votre page web qui génère et ouvre Excel document. Vous pouvez également incorporer votre Office dans le document et l’afficher à l’ouverture du document. Cela garantit que le client a toujours accès aux fonctionnalités de votre application. Lorsque le document s’ouvre, les données que le client a sélectionnées et votre Office est déjà disponible pour qu’il continue de fonctionner.

Cet article vous présente le code et les techniques permettant d’implémenter ce scénario dans votre propre application web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Créer un document Excel et incorporer un Office de document

Tout d’abord, nous allons apprendre à créer un document Excel à partir d’une page web et à incorporer un add-in dans le document. L Office exemple de code de l’incorporation de code [](https://appsource.microsoft.com/product/office/wa104380862) [ooXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) dans un Script Lab dans un nouveau document Office document. Bien que l’exemple fonctionne avec Office document, nous nous concentrerons simplement sur Excel feuilles de calcul dans cet article. Utilisez les étapes suivantes pour créer et exécuter l’exemple.

1. Extrayez l’exemple de code  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip dans un dossier de votre ordinateur.
2. Pour créer et exécuter l’exemple, suivez les étapes de la section **Pour utiliser le** projet du lisez-moi.
3. Lorsque vous exécutez l’exemple, il affiche une page web semblable à la capture d’écran suivante. Utilisez la page web pour créer un document Excel qui contient les Script Lab lors de son ouverture.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Capture d’écran de la page web affichée par l’exemple de laboratoire de script pour la sélection d’un fichier Excel et son incorporation.":::

### <a name="how-the-sample-works"></a>Fonctionnement de l’exemple

L’exemple de code utilise le SDK OOXML pour incorporer le Script Lab dans le document Excel que vous choisissez. Les informations suivantes sont issues de la section à propos [ **du code**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) dans le fichier Lisez-moi.

Le fichier **Home.aspx.cs**:

- Fournit les handlers d’événements de bouton et la manipulation de base de l’interface utilisateur.
- Utilise des techniques ASP.NET standard pour charger et télécharger le fichier.
- Utilise l’extension de nom de fichier du fichier téléchargé (xlsx, docx ou pptx) pour déterminer le type de fichier. Cette étape doit être effectuée au départ, car le SDK Open XML possède généralement des API distinctes pour chaque type de fichier.
- Appels dans **OOXMLHelper** pour valider le fichier et appels dans **le AddInEmbedder** pour incorporer des Script Lab dans le fichier et définir pour s’ouvrir automatiquement.

Le fichier **AddInEmbedder.cs**:

- Fournit la logique métier principale, qui dans cet exemple est une méthode qui incorpore Script Lab.
- Effectue des appels dans l’aide OOXML en fonction du type de fichier.

Le fichier **OOXMLHelper.cs**:

- Fournit toutes les manipulations OOXML détaillées.
- Utilise une technique standard pour valider le fichier Office, qui consiste simplement à appeler la **méthode Document.Open** dessus. Si le fichier n’est pas valide, la méthode envoie une exception.
- Contient principalement du code généré par les outils de productivité du SDK Open XML 2.5 qui sont disponibles sur le lien du [SDK Open XML 2.5.](/office/open-xml/open-xml-sdk)

La **méthode GenerateWebExtensionPart1Content** dans le fichier **OOXMLHelper.cs** définit la référence à l’ID de Script Lab dans Microsoft AppSource :

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- La **valeur StoreType** est « OMEX », alias de Microsoft AppSource.
- La **valeur** du Store est « en-US » dans la section culture Microsoft AppSource Script Lab.
- La **valeur d’ID** est l’ID d’actif Microsoft AppSource Script Lab.

Si vous souhaitez ouvrir automatiquement un module de partage de fichiers à partir d’un catalogue de partages de fichiers, vous utiliserez différentes valeurs :

La **valeur StoreType** est « FileSystem ».

- La **valeur du Store** est l’URL du partage réseau . par exemple, « \\ \\ MyComputer \\ MySharedFolder ». Il doit s’agit de l’URL exacte qui apparaît en tant qu’adresse de catalogue approuvé du partage dans Office de confiance.
- La **valeur de l’ID** est l’ID de l’application dans le manifeste des applications.
> [!NOTE]
> Pour plus d’informations sur les autres valeurs de ces attributs, voir Ouvrir automatiquement un volet Des tâches [avec un document.](../develop/automatically-open-a-task-pane-with-a-document.md)

## <a name="use-the-fluent-ui"></a>Utiliser l’interface Fluent utilisateur

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent Icônes d’interface utilisateur pour Word, Excel et PowerPoint.":::

Une meilleure pratique consiste à utiliser l’interface utilisateur Fluent pour aider vos utilisateurs à passer d’un produit Microsoft à un autre. Vous devez toujours utiliser une icône Office pour indiquer quelle application Office sera lancée à partir de votre page web. Nous allons modifier l’exemple de code pour utiliser l’icône Excel pour indiquer qu’il lance l’application Excel application.

1. Ouvrez l’exemple dans Visual Studio.
1. Ouvrez la page **Home.aspx.**
1. Recherchez le code suivant qui est le bouton de téléchargement sur le formulaire.

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. Remplacez le code du bouton par la balise d’image suivante.

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. Appuyez **sur F5** (ou **déboguer > démarrer le débogage).** L’icône s’affiche lors du chargement de la page d’accueil.

Pour plus d’informations, [voir Office Icônes](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) de marque sur le portail Fluent de l’interface utilisateur.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Télécharger le Excel document à Microsoft OneDrive

Nous vous recommandons de télécharger de nouveaux documents vers OneDrive si votre client utilise OneDrive. Cela leur permet de trouver et d’utiliser plus facilement les documents. Nous allons créer un exemple de code et voir comment vous pouvez utiliser le SDK Microsoft Graph pour télécharger un nouveau document Excel sur OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Utiliser un démarrage rapide pour créer une application web Microsoft Graph

1. Suivez les étapes de création et d’ouverture d’un exemple de code de démarrage rapide qui [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) interagit avec Office services.
1. À **l’étape 1 : choisissez la langue ou la** plateforme, choisissez ASP.NET **MVC.** Bien que les étapes de cette procédure utilisent l’option ASP.NET MVC, elles suivent un modèle qui s’applique à n’importe quelle langue ou plateforme.
1. À **l’étape 2 : Obtenez un ID d’application** et une secret, choisissez Obtenir un ID d’application et un **secret**.
1. Connectez-vous à Microsoft 365 compte.  
1. Dans la page **Web Veuillez enregistrer votre** secret d’application, enregistrez-la dans un emplacement de fichier où vous pourrez l’extraire et l’utiliser ultérieurement.
1. Choose **Got it, take me back to the quick start**.
1. À **l’étape 2 : l’inscription a réussi !** Entrez la secret de l’application générée.
1. In **step 3: Start coding**, choose **Download the SDK-based code sample**.
1. Extrayez le dossier zip de téléchargement dans un dossier local.  
1. Ouvrez le fichier graph-tutorial.sln dans Visual Studio 2019.
1. Créez et exécutez la solution et confirmez qu’elle fonctionne correctement. Vous devriez être en mesure d’utiliser la page web de calendrier pour afficher votre Microsoft 365 calendrier.

### <a name="upload-a-file-to-onedrive"></a>Télécharger fichier à OneDrive

1. Ouvrez la solution **graph-tutorial.sln** Visual Studio 2019 et ouvrez **le fichierPrivateSettings.config** graphique.

1. Ajoutez une nouvelle **étendue Files.ReadWrite** à la clé   **ida:AppScopes** afin qu’elle ressemble au code suivant.

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. Ouvrez **le fichier Index.cshtml.**
1. Insérez le code ActionLink suivant pour créer un bouton pour télécharger un fichier vers OneDrive.

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. Ouvrez **le fichier HomeController.cs.**
1. Insérez le code suivant pour gérer la demande à partir du lien d’action.

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. Ouvrez **le fichier GraphHelper.cs.**
1. Insérez le code suivant pour appeler l’API microsoft Graph pour créer un fichier sur OneDrive.

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

1. Appuyez **sur F5** (ou **déboguer > démarrer le débogage).** L’application web démarre.
1. Cliquez **ici pour vous connectez** et connectez-vous.
1. Cliquez **ici pour créer un fichier sur OneDrive**.
1. Ouvrez un nouvel onglet de navigateur et connectez-vous à OneDrive compte. Vous verrez le fichier test.txt dans le dossier racine.

Maintenant que vous avez appris à télécharger un fichier vers OneDrive, vous pouvez réutiliser ce code pour télécharger n’importe quel document Excel que vous créez.

## <a name="additional-considerations-for-your-solution"></a>Considérations supplémentaires pour votre solution

La solution de tout le monde est différente en termes de technologies et d’approches. Les considérations suivantes vous aideront à planifier la modification de votre solution pour ouvrir des documents et incorporer votre Office de données.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Créer une feuille Excel feuille de calcul à partir de la page web

L’exemple modifie un document Excel existant. Un scénario plus courant consiste à créer une feuille de calcul Excel feuille de calcul à partir de votre page web. Vous trouverez des détails supplémentaires sur la création d’une feuille de calcul dans Créer un **document** de feuilles de calcul en fournissant un nom de fichier. Cet article montre comment créer le fichier localement, mais vous pouvez également créer le fichier dans un flux à l’aide d’une surcharge sur la méthode SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Lire les propriétés personnalisées au démarrage de votre add-in

L’exemple de code stocke un ID d’extrait de code dans le nouveau document Excel l’aide du SDK OOXML. Script Lab lit l’ID de l’extrait de code à partir du document Excel puis affiche ce code d’extrait de code lorsqu’il s’ouvre. Vous devrez peut-être envoyer des propriétés personnalisées à votre propre add-in (par exemple, une chaîne de requête ou un jeton d’authentification temporaire).) Pour **plus d’informations** sur la lecture des propriétés personnalisées au démarrage de votre compl?ment, voir l’état et les paramètres persistants du compl?ment.

### <a name="initialize-the-excel-document-with-data"></a>Initialiser le document Excel avec des données

En règle générale, lorsque le client ouvre un document Excel partir de votre site web, il s’attend à ce que le document contienne des données du site web. Il existe deux façons d’écrire des données dans le document.

- **Utilisez le SDK OOXML pour écrire les données.** Vous pouvez utiliser le SDK pour écrire directement des données dans le document. Cette approche est utile si vous souhaitez que les données soient disponibles dès que le document est ouvert.
- **Passez une propriété de requête personnalisée à votre Office.** Lorsque vous générez le document, vous incorporez une propriété personnalisée pour le Office qui contient une chaîne de requête qui récupère toutes les données requises. Lorsque votre application s’ouvre, elle récupère la requête, l’exécute et utilise l’API JS Office pour insérer le résultat de la requête dans le document.

### <a name="working-with-the-ooxml-sdk"></a>Travailler avec le SDK OOXML

Le SDK OOXML est basé sur .NET. Si votre application web n’est pas .NET, vous devez rechercher une autre façon de travailler avec OOXML.

Il existe une version JavaScript du SDK OOXML disponible dans le [SDK Open XML pour JavaScript.](https://archive.codeplex.com/?p=openxmlsdkjs)

Vous pouvez placer le code OOXML dans une fonction Azure pour séparer le code .NET du reste de votre application web. Appelez ensuite la fonction Azure (pour générer le Excel document) à partir de votre application Web. Pour plus d’informations sur les fonctions Azure, voir [une présentation des fonctions Azure.](/azure/azure-functions/functions-overview)

### <a name="use-single-sign-on"></a>Utiliser l' sign-on unique

Pour simplifier l’authentification, nous vous recommandons d’implémenter l’authentification unique. Pour plus d’informations, [voir Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>Voir aussi

- [Bienvenue dans le SDK Open XML 2.5 pour Office](/office/open-xml/open-xml-sdk)
- [Ouvrir automatiquement un volet de tâches avec un document](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Conservation de l’état et des paramètres des compléments](../develop/persisting-add-in-state-and-settings.md)
- [Créer un document de feuilles de calcul en fournissant un nom de fichier](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)