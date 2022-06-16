---
title: Création de votre premier complément Outlook
description: Découvrez comment créer un complément de volet des tâches Outlook simple à l’aide de l’API JavaScript pour Office.
ms.date: 06/10/2022
ms.prod: outlook
ms.localizationpriority: high
ms.openlocfilehash: 56f43e157db9875165689af59ade50b0752fe8dc
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091096"
---
# <a name="build-your-first-outlook-add-in"></a>Création de votre premier complément Outlook

Dans cet article, vous découvrirez comment créer un complément du volet Office Outlook qui affiche au moins une propriété d’un message sélectionné.

## <a name="create-the-add-in"></a>Créer le complément

Vous pouvez créer un complément Office à l’aide du [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) ou de Visual Studio. Le générateur Yeoman crée un projet Node.js qui peut être géré avec du Visual Studio Code ou n’importe quel autre éditeur, alors que Visual Studio crée une solution Visual Studio. Sélectionnez l’onglet correspondant à votre choix, puis suivez les instructions de création de votre complément et testez-le localement.

# <a name="yeoman-generator"></a>[Générateur Yeoman](#tab/yeomangenerator)

### <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Visual Studio Code (VS Code)](https://code.visualstudio.com/) ou votre éditeur de code préféré

- Outlook 2016 ou plus récent sur Windows (connecté à un compte Microsoft 365) ou Outlook sur le web

### <a name="create-the-add-in-project"></a>Création du projet de complément

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Sélectionnez un type de projet** - `Office Add-in Task Pane project`

    - **Sélectionnez un type de script** - `JavaScript`

    - **Comment souhaitez-vous nommer votre complément ?** - `My Office Add-in`

    - **Quelle application client Office voulez-vous prendre en charge ?** - `Outlook`

    ![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande.](../images/yo-office-outlook-1.png)

    Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Accédez au dossier racine du projet de l’application web.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.
- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/taskpane.js** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Outlook.

### <a name="update-the-code"></a>Mettre à jour le code

1. Ouvrez votre projet dans VS Code ou votre éditeur de code préféré.
   [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

1. Ouvrez le fichier **./src/taskpane/taskpane.html** et remplacez tout **\<main\>** l'élément (dans **\<body\>** l'élément) par le balisage suivant. Ce nouveau balisage ajoute une étiquette à l’emplacement où le script dans **./src/taskpane/taskpane.js** écrira des données.

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. Dans votre éditeur de code, ouvrez le fichier **./src/taskpane/taskpane.js** et ajoutez le code suivant à la fonction **run**. Ce code utilise l'API Office JavaScript pour obtenir une référence au message actuel et écrire la valeur de sa propriété **objet** dans le volet des tâches.

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a>Essayez

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre et votre complément est chargé en mode [sideload](../outlook/sideload-outlook-add-ins-for-testing.md).

    ```command&nbsp;line
    npm start
    ```

1. Sur Outlook, affichez un message dans le [volet de lecture](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0)ou ouvrez le message dans sa propre fenêtre.

1. Sélectionnez l’onglet **Accueil** (ou l’onglet **Message** si vous avez ouvert le message dans une nouvelle fenêtre), puis sélectionnez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran illustrant la fenêtre d’un message dans Outlook avec le bouton du ruban du complément mis en évidence.](../images/quick-start-button-1.png)

    > [!NOTE]
    > Si le message d’erreur « Désolé... nous ne pouvons pas ouvrir ce complément à partir de localhost » s’affiche dans le volet Office, suivez les étapes décrites dans l’[article résolution des problèmes](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

1. Lorsque la boîte de dialogue **WebView Stop On Load** apparaît, sélectionnez **OK**.

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

1. Faites défiler vers le bas du volet Office et sélectionnez le lien **Exécuter** pour écrire l’objet du message dans le volet Office.

    ![Capture d’écran illustrant le volet Office du complément avec le lien d’exécution mis en évidence.](../images/quick-start-task-pane-2.png)

    ![Capture d’écran du volet Office du complément, affichant le sujet du message.](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez réussi à créer votre premier complément de volet de tâches Outlook ! Ensuite, apprenez-en davantage sur les capacités d'un complément Outlook et créez un complément plus complexe en suivant le [tutoriel sur les compléments Outlook](../tutorials/outlook-tutorial.md).

# <a name="visual-studio"></a>[Visual Studio](#tab/visualstudio)

### <a name="prerequisites"></a>Conditions préalables

- [Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée

    > [!NOTE]
    > Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.

- Microsoft 365

    > [!NOTE]
    > Si vous n’avez pas d’abonnement Microsoft 365, vous pouvez en obtenir un gratuitement en vous inscrivant au [programme développeur Microsoft 365](https://developer.microsoft.com/office/dev-program).

### <a name="create-the-add-in-project"></a>Création du projet de complément

1. Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.

1. Dans la liste des types de projets sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web Outlook** pour le type de projet.

1. Nommez le projet, puis cliquez sur **OK**.

1. Visual Studio crée une solution et ses deux projets apparaissent dans **Solution Explorer** . Le **fichier MessageRead.html** s'ouvre dans Visual Studio.

### <a name="explore-the-visual-studio-solution"></a>Explorer la solution Visual Studio

Quand vous arrivez au bout de l’Assistant, Visual Studio crée une solution qui contient deux projets.

|**Project**|**Description**|
|:-----|:-----|
|Projet de complément|Contient seulement un fichier de manifeste XML, qui contient tous les paramètres qui décrivent votre complément. Ces paramètres aident l’application Office à déterminer quand votre complément devrait être activé et où il devrait apparaître. Visual Studio génère les contenus de ce fichier pour vous afin que vous puissiez exécuter le projet et utiliser immédiatement votre complément. Vous pouvez modifier ces paramètres à tout moment en modifiant le fichier XML.|
|Projet d’application web|Contient les pages de contenu de votre complément, notamment tous les fichiers et références de fichiers dont vous avez besoin pour développer des pages HTML et JavaScript compatibles avec Office. Pendant que vous développez votre complément, Visual Studio héberge l’application web sur votre serveur IIS local. Lorsque vous êtes prêt à publier le complément, vous devez déployer ce projet d’application web sur un serveur web.|

### <a name="update-the-code"></a>Mise à jour du code

1. **MessageRead.html** spécifie le code HTML qui s’affichera dans le volet Office du complément. Dans **MessageRead.html** , remplacez **\<body\>** l'élément par le balisage suivant et enregistrez le fichier.
 
    ```HTML
    <body class="ms-font-m ms-welcome">
        <div class="ms-Fabric content-main">
            <h1 class="ms-font-xxl">Message properties</h1>
            <table class="ms-Table ms-Table--selectable">
                <thead>
                    <tr>
                        <th>Property</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>Id</strong></td>
                        <td class="prop-val"><code><label id="item-id"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Subject</strong></td>
                        <td class="prop-val"><code><label id="item-subject"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Message Id</strong></td>
                        <td class="prop-val"><code><label id="item-internetMessageId"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>From</strong></td>
                        <td class="prop-val"><code><label id="item-from"></label></code></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
    ```

1. Ouvrez le fichier **MessageRead.js** à la racine du projet d'application Web. Ce fichier spécifie le script de ce complément. Remplacez l'ensemble du contenu par le code suivant et enregistrez le fichier.

    ```js
    'use strict';

    (function () {

        Office.onReady(function () {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                loadItemProps(Office.context.mailbox.item);
            });
        });

        function loadItemProps(item) {
            // Write message property values to the task pane
            $('#item-id').text(item.itemId);
            $('#item-subject').text(item.subject);
            $('#item-internetMessageId').text(item.internetMessageId);
            $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        }
    })();
    ```

1. Ouvrez le fichier **MessageRead.css** à la racine du projet d'application Web. Ce fichier spécifie les styles personnalisés pour le module d'extension. Remplacez l'ensemble du contenu par le code suivant et enregistrez le fichier.

    ```CSS
    html,
    body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    td.prop-val {
        word-break: break-all;
    }

    .content-main {
        margin: 10px;
    }
    ```

### <a name="update-the-manifest"></a>Mise à jour du manifeste

1. Ouvrez le fichier manifeste XML dans le projet de complément. Ce fichier définit les paramètres et les fonctionnalités du complément.

1. L'élément **ProviderName** a une valeur de type placeholder. Remplacez-la par votre nom.

1. L'attribut **DefaultValue** de l'élément **DisplayName** comporte un espace réservé. Remplacez-le par `My Office Add-in`.

1. L'attribut **DefaultValue** de l'élément **Description** contient un caractère générique. Remplacez-le par`My First Outlook add-in`.

1. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a>Try it out

1. À l’aide de Visual Studio, testez le complément Outlook que vous venez de créer en appuyant sur F5 ou en sélectionnant le bouton **Démarrer**. Le complément est hébergé localement sur IIS.

1. Dans la boîte de dialogue **Se connecter à un compte de messagerie Exchange**, entrez l’adresse de messagerie et mot de passe pour votre [compte Microsoft](https://account.microsoft.com/account), puis sélectionnez **Se connecter**. Lorsque la page de connexion Outlook.com s’ouvre dans un navigateur, connectez-vous à votre compte de courrier avec les mêmes informations d’identification que vous avez entrées précédemment.

    > [!NOTE]
    > Si la boîte de dialogue **Se connecter au compte de messagerie Exchange** vous invite à vous connecter à plusieurs reprises ou si vous recevez une erreur indiquant que vous n’êtes pas autorisé, il se peut que l’authentification de base soit désactivée pour les comptes sur votre client Microsoft 365. Pour tester ce complément, réessayez de vous connecter après avoir défini la propriété **Utilisez l’auth multifacteur** sur True dans la boîte de dialogue propriétés du projet de complément web, ou connectez-vous à l’aide d’un [compte Microsoft](https://account.microsoft.com/account) à la place.

1. Dans Outlook sur le web, sélectionnez ou ouvrez un message.

1. Dans le message, recherchez les points de suspension du menu de dépassement de capacité contenant le bouton du complément.

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le web avec les points de suspension mis en surbrillance.](../images/quick-start-button-owa-1.png)

1. Dans le menu de dépassement de capacité, recherchez le bouton du complément.

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le web avec le bouton du complément mis en surbrillance.](../images/quick-start-button-owa-2.png)

1. Cliquez sur le bouton pour ouvrir le volet Office du complément.

    ![Capture d’écran du volet Office du complément dans Outlook sur le web, affichant les propriétés des messages.](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > Si le volet Office n’est pas chargé, essayez de l’ouvrir dans un navigateur sur le même ordinateur.

### <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé votre premier complément de volet de tâches Outlook ! Ensuite, en savoir plus sur la [création de compléments Office avec Visual Studio](../develop/develop-add-ins-visual-studio.md).

---

## <a name="see-also"></a>Voir aussi

- [Utilisation de Visual Studio Code pour publier](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
