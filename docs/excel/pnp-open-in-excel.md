---
title: Ouvrir Excel à partir de votre page web et incorporer votre complément Office
description: Ouvrez Excel à partir de votre page web et incorporez votre complément Office.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 835518fb822602d6ca1af633f96d2be1e2697f44
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810343"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Ouvrir Excel à partir de votre page web et incorporer votre complément Office

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Image du bouton Excel sur votre page web ouvrant un nouveau document Excel avec votre complément incorporé et ouvert automatiquement.":::

Étendez votre application web SaaS afin que vos clients puissent ouvrir leurs données à partir d’une page web directement dans Microsoft Excel. Un scénario courant est que les clients devront utiliser des données dans votre application web. Ensuite, ils souhaitent copier les données dans un document Excel. Par exemple, ils peuvent souhaiter effectuer une analyse supplémentaire à l’aide d’Excel. En règle générale, le client doit exporter les données vers un fichier, tel qu’un fichier .csv, puis importer ces données dans Excel. Ils doivent également ajouter manuellement votre complément Office au document.

Réduisez le nombre d’étapes à un seul clic sur votre page web qui génère et ouvre le document Excel. Vous pouvez également incorporer votre complément Office dans le document et l’afficher à l’ouverture du document. Cela garantit que le client a toujours accès aux fonctionnalités de votre application. Lorsque le document s’ouvre, les données sélectionnées par le client et votre complément Office sont déjà disponibles pour qu’il continue à travailler.

Cet article présente le code et les techniques d’implémentation de ce scénario dans votre propre application web SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Créer un document Excel et incorporer un complément Office

Tout d’abord, nous allons apprendre à créer un document Excel à partir d’une page web et à incorporer un complément dans le document. [L’exemple de code de complément incorporé Office OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) montre comment incorporer le [complément Script Lab dans](https://appsource.microsoft.com/product/office/wa104380862) un nouveau document Office. Bien que l’exemple fonctionne avec n’importe quel document Office, nous allons nous concentrer sur les feuilles de calcul Excel dans cet article. Procédez comme suit pour générer et exécuter l’exemple.

1. Extrayez l’exemple de code de  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip dans un dossier sur votre ordinateur.
2. Pour générer et exécuter l’exemple, suivez les étapes décrites dans la section **Pour utiliser le projet** du fichier Lisez-moi.
3. Lorsque vous exécutez l’exemple, il affiche une page web similaire à la capture d’écran suivante. Utilisez la page web pour créer un document Excel qui contient Script Lab lorsqu’il s’ouvre.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Capture d’écran de la page web que l’exemple de labo de script incorporé affiche pour sélectionner un fichier Excel et incorporer le complément de labo de script dans celui-ci.":::

### <a name="how-the-sample-works"></a>Fonctionnement de l’exemple

L’exemple de code utilise le Kit de développement logiciel (SDK) OOXML pour incorporer le complément Script Lab au document Excel que vous choisissez. Les informations suivantes proviennent de la [section **À propos du code**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) dans le fichier Lisez-moi.

Le fichier **Home.aspx.cs** :

- Fournit les gestionnaires d’événements de bouton et la manipulation de base de l’interface utilisateur.
- Utilise des techniques de ASP.NET standard pour charger et télécharger le fichier.
- Utilise l’extension de nom de fichier du fichier chargé (xlsx, docx ou pptx) pour déterminer le type de fichier. Cette opération doit être effectuée au début, car le Kit de développement logiciel (SDK) Open XML a généralement des API distinctes pour chaque type de fichier.
- Appelle **l’outil OOXMLHelper** pour valider le fichier et appelle **l’AddInEmbedder** pour incorporer Script Lab dans le fichier et l’ouvrir automatiquement.

Le fichier **AddInEmbedder.cs** :

- Fournit la logique métier principale, qui dans cet exemple est une méthode qui incorpore Script Lab.
- Effectue des appels dans l’assistance OOXML en fonction du type du fichier.

Le fichier **OOXMLHelper.cs** :

- Fournit toutes les manipulations OOXML détaillées.
- Utilise une technique standard pour valider le fichier Office, qui consiste simplement à appeler la méthode **Document.Open** sur celui-ci. Si le fichier n’est pas valide, la méthode lève une exception.
- Contient principalement du code généré par les outils de productivité du Kit de développement logiciel (SDK) Open XML 2.5, qui sont disponibles sur le lien pour le [Kit de développement logiciel (SDK) Open XML 2.5](/office/open-xml/open-xml-sdk).

La méthode **GenerateWebExtensionPart1Content** dans le fichier **OOXMLHelper.cs** définit la référence à l’ID de Script Lab dans Microsoft AppSource :

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- La valeur **StoreType** est « OMEX », un alias pour Microsoft AppSource.
- La valeur **du Store** est « en-US » dans la section Culture Microsoft AppSource pour Script Lab.
- La valeur **Id** est l’ID de ressource Microsoft AppSource pour Script Lab.

Si vous configurez un complément à partir d’un catalogue de partage de fichiers pour l’ouverture automatique, vous utiliserez différentes valeurs :

La valeur **StoreType** est « FileSystem ».

- La valeur **Store** est l’URL du partage réseau ; par exemple, «\\\\ MyComputer\\MySharedFolder ». Il doit s’agir de l’URL exacte qui apparaît comme adresse de catalogue approuvé du partage dans le Centre de gestion de la confidentialité Office.
- La valeur **Id** est l’ID d’application dans le manifeste des compléments.
> [!NOTE]
> Pour plus d’informations sur les valeurs alternatives pour ces attributs, consultez [Ouvrir automatiquement un volet Office avec un document](../develop/automatically-open-a-task-pane-with-a-document.md).

## <a name="use-the-fluent-ui"></a>Utiliser l’interface utilisateur Fluent

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Icônes de l’interface utilisateur Fluent pour Word, Excel et PowerPoint.":::

Une bonne pratique consiste à utiliser l’interface utilisateur Fluent pour aider vos utilisateurs à passer d’un produit Microsoft à l’autre. Vous devez toujours utiliser une icône Office pour indiquer quelle application Office sera lancée à partir de votre page web. Nous allons modifier l’exemple de code pour utiliser l’icône Excel pour indiquer qu’il lance l’application Excel.

1. Ouvrez l’exemple dans Visual Studio.
1. Ouvrez la page **Home.aspx** .
1. Recherchez le code suivant, qui est le bouton de téléchargement sur le formulaire.

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. Remplacez le code du bouton par la balise image suivante.

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. Appuyez sur **F5** (ou **Déboguer** > **Démarrer le débogage**). L’icône s’affiche lorsque la page d’accueil se charge.

Pour plus d’informations, voir [Icônes de marque Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) sur le portail des développeurs Fluent UI.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Charger le document Excel dans Microsoft OneDrive

Nous vous recommandons de charger de nouveaux documents sur OneDrive si votre client utilise OneDrive. Cela leur permet de trouver et d’utiliser plus facilement les documents. Nous allons créer un exemple de code et voir comment vous pouvez utiliser le Kit de développement logiciel (SDK) Microsoft Graph pour charger un nouveau document Excel sur OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Utiliser un démarrage rapide pour créer une nouvelle application web Microsoft Graph

1. Accédez à [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) et suivez les étapes pour créer et ouvrir un exemple de code de démarrage rapide qui interagit avec les services Office.
1. À **l’étape 1 : Choisissez votre langue ou votre plateforme**, choisissez **ASP.NET MVC**. Bien que les étapes de cette procédure utilisent l’option ASP.NET MVC, les étapes suivent un modèle qui s’applique à n’importe quel langage ou plateforme.
1. À **l’étape 2 : Obtenir un ID d’application et un secret**, choisissez **Obtenir un ID d’application et un secret**.
1. Connectez-vous à votre compte Microsoft 365.  
1. Dans la page web **Veuillez enregistrer votre secret d’application** , enregistrez le secret de l’application dans un emplacement de fichier où vous pourrez le récupérer et l’utiliser ultérieurement.
1. Choisissez **Vous l’avez obtenu, revenez-moi au démarrage rapide**.
1. À **l’étape 2 : Inscription réussie !** Entrez le secret d’application généré.
1. À **l’étape 3 : Démarrer le codage**, choisissez **Télécharger l’exemple de code basé sur le SDK**.
1. Extrayez le dossier zip de téléchargement dans un dossier local.  
1. Ouvrez le fichier graph-tutorial.sln dans Visual Studio 2019.
1. Générez et exécutez la solution et vérifiez qu’elle fonctionne correctement. Vous devez être en mesure d’utiliser la page web du calendrier pour afficher votre calendrier Microsoft 365.

### <a name="upload-a-file-to-onedrive"></a>Charger un fichier sur OneDrive

1. Ouvrez la solution **graph-tutorial.sln** dans Visual Studio 2019 et ouvrez le fichier **PrivateSettings.config** .

1. Ajoutez une nouvelle étendue **Files.ReadWrite** à la clé **ida:AppScopes** afin qu’elle ressemble au code suivant.

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. Ouvrez le fichier **Index.cshtml** .
1. Insérez le code ActionLink suivant pour créer un bouton permettant de charger un fichier sur OneDrive.

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. Ouvrez le fichier **HomeController.cs** .
1. Insérez le code suivant pour gérer la requête à partir du lien d’action.

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
1. Insérez le code suivant pour appeler le API Graph Microsoft afin de créer un fichier sur OneDrive.

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

1. Appuyez sur **F5** (ou **Déboguer** > **Démarrer le débogage**). L’application web démarre.
1. Choisissez **Cliquez ici pour vous connecter**, puis connectez-vous.
1. Choisissez **Cliquez ici pour créer un fichier sur OneDrive**.
1. Ouvrez un nouvel onglet de navigateur et connectez-vous à votre compte OneDrive. Le fichier test.txt s’affiche dans le dossier racine.

Maintenant que vous avez appris à charger un fichier sur OneDrive, vous pouvez réutiliser ce code pour charger n’importe quel document Excel que vous créez.

## <a name="additional-considerations-for-your-solution"></a>Considérations supplémentaires pour votre solution

La solution de chacun est différente en termes de technologies et d’approches. Les considérations suivantes vous aideront à planifier la modification de votre solution pour ouvrir des documents et incorporer votre complément Office.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Créer une feuille de calcul Excel à partir de la page web

L’exemple modifie un document Excel existant. Un scénario plus courant consiste à créer une feuille de calcul Excel à partir de votre page web. Vous trouverez des informations supplémentaires sur la création d’une feuille de calcul dans **Créer un document de feuille de calcul** en fournissant un nom de fichier. Cet article explique comment créer le fichier localement, mais vous pouvez également créer le fichier dans un flux à l’aide d’une surcharge sur la méthode SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Lire les propriétés personnalisées au démarrage de votre complément

L’exemple de code stocke un ID d’extrait de code dans le nouveau document Excel à l’aide du Kit de développement logiciel (SDK) OOXML. Script Lab lit l’ID de l’extrait de code à partir du document Excel, puis affiche ce code d’extrait de code lorsqu’il s’ouvre. Vous devrez peut-être envoyer des propriétés personnalisées à votre propre complément (par exemple, une chaîne de requête ou un jeton d’authentification temporaire). Pour plus d’informations sur la façon de lire les propriétés personnalisées au démarrage de votre complément, consultez **Persistance de l’état et des paramètres** du complément.

### <a name="initialize-the-excel-document-with-data"></a>Initialiser le document Excel avec des données

En règle générale, quand le client ouvre un document Excel à partir de votre site web, il s’attend à ce que le document contienne des données du site web. Il existe deux façons d’écrire des données dans le document.

- **Utilisez le Kit de développement logiciel (SDK) OOXML pour écrire les données**. Vous pouvez utiliser le Kit de développement logiciel (SDK) pour écrire directement des données dans le document. Cette approche est utile si vous souhaitez que les données soient disponibles au moment où le document est ouvert.
- **Transmettez une propriété de requête personnalisée à votre complément Office**. Lorsque vous générez le document, vous incorporez une propriété personnalisée pour le complément Office qui contient une chaîne de requête qui récupère toutes les données requises. Lorsque votre complément s’ouvre, il récupère la requête, exécute la requête et utilise l’API Office JS pour insérer le résultat de la requête dans le document.

### <a name="working-with-the-ooxml-sdk"></a>Utilisation du Kit de développement logiciel (SDK) OOXML

Le Kit de développement logiciel (SDK) OOXML est basé sur .NET. Si votre application web n’est pas .NET, vous devez rechercher un autre moyen d’utiliser OOXML.

Vous pouvez placer le code OOXML dans une fonction Azure pour séparer le code .NET du reste de votre application web. Appelez ensuite la fonction Azure (pour générer le document Excel) à partir de votre application web. Pour plus d’informations sur les fonctions Azure, consultez [Présentation de Azure Functions](/azure/azure-functions/functions-overview).

### <a name="use-single-sign-on"></a>Utiliser l’authentification unique

Pour simplifier l’authentification, nous vous recommandons d’implémenter l’authentification unique dans votre complément. Pour plus d’informations, voir [Activer l’authentification unique pour les compléments Office](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>Voir aussi

- [Bienvenue dans le Kit de développement logiciel (SDK) Open XML 2.5 pour Office](/office/open-xml/open-xml-sdk)
- [Ouvrir automatiquement un volet de tâches avec un document](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Conservation de l’état et des paramètres des compléments](../develop/persisting-add-in-state-and-settings.md)
- [Créer un document de feuilles de calcul en fournissant un nom de fichier](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)