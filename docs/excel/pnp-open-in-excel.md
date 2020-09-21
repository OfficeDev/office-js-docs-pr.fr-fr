---
title: Ouvrir Excel à partir de votre page Web et incorporer votre complément Office
description: Ouvrez Excel à partir de votre page Web et incorporez votre complément Office.
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 49df253c714f3ad84d2523b87e7df894b9027355
ms.sourcegitcommit: ea03e4ea2e8537d5f6d52477816209f6c1a6579c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2020
ms.locfileid: "48166921"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Ouvrir Excel à partir de votre page Web et incorporer votre complément Office

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Image du bouton Excel de votre page Web ouverture d’un nouveau document Excel avec votre complément incorporé et en cours d’ouverture automatique.":::

Étendez votre application Web SaaS pour permettre à vos clients d’ouvrir leurs données à partir d’une page Web directement dans Microsoft Excel. Un scénario courant est que les clients travailleront avec des données dans votre application Web. Ils voudront ensuite copier les données dans un document Excel. Par exemple, ils peuvent souhaiter effectuer des analyses supplémentaires à l’aide d’Excel. En règle générale, le client doit exporter les données vers un fichier, par exemple un fichier. csv, puis importer ces données dans Excel. Ils doivent également ajouter manuellement votre complément Office au document.

Réduisez le nombre d’étapes à un seul clic sur votre page Web qui génère et ouvre le document Excel. Vous pouvez également incorporer votre complément Office dans le document et l’afficher à l’ouverture du document. Cela garantit que le client a toujours accès aux fonctionnalités de votre application. Lorsque le document s’ouvre, les données sélectionnées par le client et votre complément Office sont déjà disponibles pour qu’ils continuent de fonctionner.

Cet article décrit le code et les techniques à appliquer pour mettre en œuvre ce scénario dans votre application Web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Créer un nouveau document Excel et incorporer un complément Office

Tout d’abord, nous allons apprendre à créer un document Excel à partir d’une page Web et à incorporer un complément dans le document. L' [exemple de code de complément Office OOXML embed](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) montre comment incorporer le [complément script Lab](https://appsource.microsoft.com/product/office/wa104380862) dans un nouveau document Office. Bien que l’exemple fonctionne avec n’importe quel document Office, nous allons nous concentrer sur les feuilles de calcul Excel dans cet article. Procédez comme suit pour générer et exécuter l’exemple.

1. Extrayez l’exemple de code de  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip dans un dossier sur votre ordinateur.
2. Pour générer et exécuter l’exemple, suivez les étapes de la section **pour utiliser le projet** du fichier Lisez-moi.
3. Lorsque vous exécutez l’exemple, il affiche une page Web semblable à la capture d’écran suivante. Utilisez la page Web pour créer un nouveau document Excel qui contient script Lab lors de son ouverture.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Capture d’écran de la page Web que l’exemple de script embed permet d’afficher pour sélectionner un fichier Excel et incorporer le complément script Lab dans celui-ci.":::

### <a name="how-the-sample-works"></a>Fonctionnement de l’exemple

L’exemple de code utilise le kit de développement logiciel (SDK) OOXML pour incorporer le complément script Lab dans le document Excel que vous choisissez. Les informations suivantes sont extraites de la [section **à propos du code** ](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) du fichier Lisez-moi.

Le fichier **Home.aspx.cs**:

- Fournit les gestionnaires d’événements de bouton et la manipulation d’interface utilisateur de base.
- Utilise des techniques ASP.NET standard pour charger et télécharger le fichier.
- Utilise l’extension de nom de fichier du fichier téléchargé (xlsx, docx ou pptx) pour déterminer le type de fichier. Cela doit être fait dès le départ, car le kit de développement logiciel (SDK) Open XML comporte généralement des API distinctes pour chaque type de fichier.
- Appelle le **OOXMLHelper** pour valider le fichier et appelle l' **AddInEmbedder** pour incorporer le script Lab dans le fichier et pour qu’il s’ouvre automatiquement.

Le fichier **AddInEmbedder.cs**:

- Fournit la logique métier principale, qui dans cet exemple est une méthode qui incorpore le script Lab.
- Effectue des appels dans l’assistance OOXML en fonction du type de fichier.

Le fichier **OOXMLHelper.cs**:

- Fournit toutes les manipulations OOXML détaillées.
- Utilise une technique standard pour valider le fichier Office, qui consiste simplement à appeler la méthode **document. Open** dessus. Si le fichier n’est pas valide, la méthode génère une exception.
- Contient principalement du code qui a été généré par les outils de productivité du kit de développement logiciel (SDK) Open XML 2,5 qui sont disponibles sur le lien du [Kit de développement logiciel (SDK) Open xml 2,5](/office/open-xml/open-xml-sdk).

La méthode **GenerateWebExtensionPart1Content** dans le fichier **OOXMLHelper.cs** définit la référence à l’ID de script Lab dans Microsoft AppSource :

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- La valeur **StoreType** est « OMEX », un alias pour Microsoft AppSource.
- La valeur **Store** est « en-US » dans la section culture de Microsoft AppSource pour script Lab.
- La valeur **ID** est l’ID de la ressource Microsoft AppSource pour script Lab.

Si vous configurez un complément à partir d’un catalogue de partage de fichiers pour une ouverture automatique, vous utiliserez des valeurs différentes :

La valeur **StoreType** est « FileSystem ».

- La valeur **Store** est l’URL du partage réseau ; par exemple, " \\ \\ MyComputer \\ MySharedFolder". Il doit s’agir de l’URL exacte qui apparaît en tant qu’adresse de catalogue approuvé du partage dans le centre de gestion de la confidentialité Office.
- La valeur **ID** est l’ID de l’application dans le manifeste des compléments.
> [!NOTE]
> Pour plus d’informations sur les autres valeurs de ces attributs, consultez la rubrique [ouvrir automatiquement un volet Office avec un document](../develop/automatically-open-a-task-pane-with-a-document.md).

## <a name="use-the-fluent-ui"></a>Utiliser l’interface utilisateur Fluent

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Icônes de l’interface utilisateur Fluent pour Word, Excel et PowerPoint.":::

Il est recommandé d’utiliser l’interface utilisateur Fluent pour aider vos utilisateurs à effectuer une transition entre les produits Microsoft. Vous devez toujours utiliser une icône Office pour indiquer l’application Office qui sera lancée à partir de votre page Web. Nous allons modifier l’exemple de code de façon à utiliser l’icône Excel pour indiquer qu’il lance l’application Excel.

1. Ouvrez l’exemple dans Visual Studio.
1. Ouvrez la page **Home. aspx** .
1. Recherchez le code suivant qui est le bouton Télécharger sur le formulaire :
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. Remplacez le code du bouton par la balise d’image suivante.
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. Appuyez sur **F5** (ou **débogage > démarrer le débogage**). L’icône apparaît lorsque la page d’accueil est chargée.

Pour plus d’informations, consultez la rubrique icônes de la [marque Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) sur le portail du développeur de l’interface utilisateur Fluent.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Télécharger le document Excel dans Microsoft OneDrive

Nous vous recommandons de télécharger de nouveaux documents sur OneDrive si votre client utilise OneDrive. Cela leur permet de rechercher et d’utiliser plus facilement les documents. Nous allons créer un nouvel exemple de code et voir comment vous pouvez utiliser le kit de développement logiciel (SDK) Microsoft Graph pour télécharger un nouveau document Excel dans OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Utiliser un démarrage rapide pour créer une nouvelle application Web Microsoft Graph

1. Accédez à [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) et suivez les étapes pour créer et ouvrir un exemple de code de démarrage rapide qui interagit avec les services Office 365.
1. À l' **étape 1 : choisissez une langue ou une plateforme**, choisissez **ASP.NET MVC**. Bien que les étapes de cette procédure utilisent l’option MVC ASP.NET, les étapes suivent un modèle qui s’applique à n’importe quel langage ou plateforme.
1. À l' **étape 2 : obtenir un ID d’application et une clé secrète**, choisissez **obtenir un ID d’application et une clé secrète**.
1. Connectez-vous à votre compte Microsoft 365.  
1. Sur la page **Veuillez enregistrer votre application secrète** de l’application, enregistrez la clé secrète de l’application à un emplacement de fichier où vous pourrez la récupérer et l’utiliser ultérieurement.
1. Choisissez **obtenu, revenez au démarrage rapide**.
1. À l' **étape 2 : inscription réussie !** Entrez le code secret de l’application générée.
1. À l' **étape 3 : démarrer le codage**, choisissez **Télécharger l’exemple de code basé sur le kit de développement logiciel (SDK)**.
1. Extrayez le dossier zip de téléchargement dans un dossier local.  
1. Ouvrez le fichier Graph-Tutorial. sln dans Visual Studio 2019.
1. Générez et exécutez la solution et assurez-vous qu’elle fonctionne correctement. Vous devriez être en mesure d’utiliser la page Web du calendrier pour afficher votre calendrier Microsoft 365.

### <a name="upload-a-file-to-onedrive"></a>Charger un fichier dans OneDrive

1. Ouvrez la solution **Graph-Tutorial. sln** dans Visual Studio 2019, puis ouvrez le fichier **PrivateSettings.config** .
1. Ajoutez un nouveau fichier d’étendue **. ReadWrite**   à la clé **Ida : AppScopes** afin qu’il se présente comme le code suivant :
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. Ouvrez le fichier **index. cshtml** .
1. Insérez le code ActionLink suivant pour créer un bouton permettant de télécharger un fichier dans OneDrive.
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. Ouvrez le fichier **HomeController.cs** .
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
1. Ouvrez le fichier **GraphHelper.cs** .
1. Insérez le code suivant pour appeler l’API Microsoft Graph afin de créer un nouveau fichier sur OneDrive.
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
1. Appuyez sur **F5** (ou **débogage > démarrer le débogage**). L’application Web démarrera.
1. Choisissez **Cliquer ici pour vous connecter**et vous connecter.
1. Choisissez **Cliquer ici pour créer un fichier sur OneDrive**.
1. Ouvrez un nouvel onglet de navigateur et connectez-vous à votre compte OneDrive. Le fichier test.txt est affiché dans le dossier racine.

Maintenant que vous avez appris à télécharger un fichier dans OneDrive, vous pouvez réutiliser ce code pour télécharger un document Excel que vous créez.

## <a name="additional-considerations-for-your-solution"></a>Considérations supplémentaires pour votre solution

La solution de tout le monde est différente en termes de technologies et d’approches. Les considérations suivantes vous aideront à planifier comment modifier votre solution pour ouvrir des documents et incorporer votre complément Office.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Créer une nouvelle feuille de calcul Excel à partir de la page Web

L’exemple modifie un document Excel existant. Un scénario plus courant consiste à créer une nouvelle feuille de calcul Excel à partir de votre page Web. Vous trouverez des informations supplémentaires sur la création d’une nouvelle feuille de calcul dans **créer un document de feuilles de calcul** en fournissant un nom de fichier. Cet article explique comment créer le fichier localement, mais vous pouvez également créer le fichier dans un flux à l’aide d’une surcharge de la méthode SpreadsheetDocument. Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Lecture des propriétés personnalisées au démarrage de votre complément

L’exemple de code stocke un ID d’extrait de code dans le nouveau document Excel à l’aide du kit de développement logiciel (SDK) OOXML. Script Lab lit l’ID d’extrait de code à partir du document Excel, puis affiche cet extrait de code à l’ouverture. Vous devrez peut-être envoyer des propriétés personnalisées à votre propre complément (par exemple, une chaîne de requête ou un jeton d’authentification temporaire). Pour plus d’informations sur la lecture des propriétés personnalisées au démarrage de votre complément, voir **Persisting Add-in State and Settings** .

### <a name="initialize-the-excel-document-with-data"></a>Initialiser le document Excel avec des données

En règle générale, lorsque le client ouvre un document Excel à partir de votre site Web, il s’attend à ce que le document contienne des données du site Web. Il existe deux façons d’écrire des données dans le document.

- **Utilisez le kit de développement logiciel (SDK) OOXML pour écrire les données**. Vous pouvez utiliser le kit de développement logiciel (SDK) pour écrire directement des données dans le document. Cette approche est utile si vous souhaitez que les données soient disponibles le moment où le document est ouvert.
- **Transmettez une propriété de requête personnalisée à votre complément Office**. Lorsque vous générez le document, vous incorporez une propriété personnalisée pour le complément Office qui contient une chaîne de requête qui récupère toutes les données requises. Lorsque votre complément s’ouvre, il récupère la requête, exécute la requête et utilise l’API Office JS pour insérer le résultat de la requête dans le document.

### <a name="working-with-the-ooxml-sdk"></a>Utilisation du kit de développement logiciel (SDK) OOXML

Le kit de développement logiciel (SDK) OOXML est basé sur .NET. Si votre application Web n’est pas .NET, vous devez rechercher une autre façon de travailler avec OOXML.

Une version JavaScript du kit de développement logiciel (SDK) OOXML est disponible dans le [Kit de développement logiciel (SDK) Open XML pour JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).

Vous pouvez placer le code OOXML dans une fonction Azure pour séparer le code .NET du reste de votre application Web. Appelez ensuite la fonction Azure (pour générer le document Excel) à partir de votre application Web. Pour plus d’informations sur les fonctions Azure, reportez-vous à la rubrique [Présentation des fonctions Azure](https://docs.microsoft.com/azure/azure-functions/functions-overview).

### <a name="simplify-authentication"></a>Simplifier l’authentification

En règle générale, le client est authentifié et connecté lorsqu’il travaille dans votre application Web. Il est recommandé de rester connecté lorsqu’il ouvre le document afin qu’il ne soit pas obligé de se connecter à nouveau pour utiliser votre complément Office. Un moyen efficace de gérer cela consiste à transmettre un jeton d’authentification de courte durée au complément.

1. Utilisez le kit de développement logiciel (SDK) OOXML pour enregistrer le jeton d’authentification en tant que propriété personnalisée dans le document.
1. Lire le jeton dans le document au démarrage du complément.
1. Le complément peut ensuite se connecter à vos services sans nécessiter d’autres étapes d’authentification de la part du client.

> [!WARNING]
> L’incorporation d’un jeton d’authentification dans le document pose un risque de sécurité où un utilisateur non autorisé peut obtenir le jeton. Nous vous recommandons d’utiliser un jeton d’authentification à courte durée de vie. Lorsque le complément utilise le jeton éphémère, il doit immédiatement demander un nouveau jeton d’authentification qui n’est pas enregistré dans le document.

## <a name="see-also"></a>Voir aussi

- [Bienvenue dans le kit de développement logiciel (SDK) Open XML 2,5 pour Office](/office/open-xml/open-xml-sdk)
- [Ouvrir automatiquement un volet de tâches avec un document](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Conservation de l’état et des paramètres des compléments](../develop/persisting-add-in-state-and-settings.md)
- [Créer un document de feuilles de calcul en fournissant un nom de fichier](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
