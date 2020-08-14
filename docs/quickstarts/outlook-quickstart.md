---
title: Création de votre premier complément Outlook
description: Découvrez comment créer un complément de volet des tâches Outlook simple à l’aide de l’API JavaScript pour Office.
ms.date: 08/11/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 6ed50b52e0f4d5667e835c875851ed14c68bfe49
ms.sourcegitcommit: 65c15a9040279901ea7ff7f522d86c8fddb98e14
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/14/2020
ms.locfileid: "46672714"
---
# <a name="build-your-first-outlook-add-in"></a>Création de votre premier complément Outlook

Dans cet article, vous découvrirez comment créer un complément du volet Office Outlook qui affiche au moins une propriété d’un message sélectionné.

## <a name="create-the-add-in"></a>Créer le complément

Vous pouvez créer un complément Office à l’aide du [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) ou de Visual Studio. Le générateur Yeoman crée un projet Node.js qui peut être géré avec du Visual Studio Code ou n’importe quel autre éditeur, alors que Visual Studio crée une solution Visual Studio.  Sélectionnez l’onglet correspondant à votre choix, puis suivez les instructions de création de votre complément et testez-le localement.

# <a name="yeoman-generator"></a>[Générateur Yeoman](#tab/yeomangenerator)

### <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- [Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))

- La dernière version de[Yeoman](https://github.com/yeoman/yo) et de [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Même si vous avez précédemment installé le générateur Yeoman, nous vous recommandons de mettre à jour votre package vers la dernière version de npm.

### <a name="create-the-add-in-project"></a>Création du projet de complément

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Sélectionnez un type de projet** - `Office Add-in Task Pane project`

    - **Sélectionnez un type de script** - `Javascript`

    - **Comment souhaitez-vous nommer votre complément ?** - `My Office Add-in`

    - **Quelle application client Office voulez-vous prendre en charge ?** - `Outlook`

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-outlook.png)
    
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

1. Dans votre éditeur de code, ouvrez le fichier **./src/taskpane/taskpane.html** et remplacez l’élément `<main>` (dans l’élément `<body>`) par le balisage suivant. Ce nouveau balisage ajoute une étiquette à l’emplacement où le script dans **./src/taskpane/taskpane.js** écrira des données.

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. Ouvrez le fichier **./src/taskpane/taskpane.js** dans l’éditeur de code et ajoutez le code suivant à la fonction `run`. Ce code utilise l’API JavaScript pour Office afin d’obtenir une référence au message en cours et écrire sa valeur de propriété `subject` dans le volet Office.

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a>Essayez !

> [!NOTE]
> Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

    ```command&nbsp;line
    npm start
    ```

1. Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md) pour charger le complément dans Outlook.

1. Dans Outlook, sélectionnez ou ouvrez un message.

1. Sélectionnez l’onglet **Accueil** (ou l’onglet **Message** si vous avez ouvert le message dans une nouvelle fenêtre), puis sélectionnez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran d’une fenêtre de message dans Outlook avec le bouton du complément mis en surbrillance](../images/quick-start-button-1.png)

    > [!NOTE]
    > Si le message d’erreur « Désolé... nous ne pouvons pas ouvrir ce complément à partir de localhost » s’affiche dans le volet Office, suivez les étapes décrites dans l’[article résolution des problèmes](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

1. Faites défiler vers le bas du volet Office et sélectionnez le lien **Exécuter** pour écrire l’objet du message dans le volet Office.

    ![Capture d’écran du volet Office du complément avec le lien d’exécution mis en évidence](../images/quick-start-task-pane-2.png)

    ![Capture d’écran du volet Office du complément, affichant le sujet du message](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé votre premier complément de volet de tâches Outlook ! Ensuite, découvrez les fonctionnalités d’un complément Outlook et créez-en un plus complexe en suivant le [didacticiel pour complément Outlook](../tutorials/outlook-tutorial.md).

# <a name="visual-studio"></a>[Visual Studio](#tab/visualstudio)

### <a name="prerequisites"></a>Conditions préalables

- [Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée

    > [!NOTE]
    > Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.

- Office 365

    > [!NOTE]
    > Si vous n’avez pas d’abonnement Microsoft 365, vous pouvez en obtenir un gratuitement en vous inscrivant au [programme développeur Microsoft 365](https://developer.microsoft.com/office/dev-program).

### <a name="create-the-add-in-project"></a>Création du projet de complément

1. Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.

1. Dans la liste des types de projets sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web Outlook** pour le type de projet.

1. Nommez le projet, puis cliquez sur **OK**.

1. Visual Studio crée une solution et ses deux projets apparaissent dans l’**Explorateur de solutions**. Le fichier **MessageRead.html** s’ouvre dans Visual Studio.

### <a name="explore-the-visual-studio-solution"></a>Explorer la solution Visual Studio

Quand vous arrivez au bout de l’Assistant, Visual Studio crée une solution qui contient deux projets.

|**Project**|**Description**|
|:-----|:-----|
|Projet de complément|Contient uniquement un fichier manifeste XML contenant tous les paramètres qui décrivent votre complément. Ces paramètres aident l’hôte Office à déterminer le moment où votre complément doit être activé et l’emplacement où il doit apparaître. Visual Studio génère le contenu de ce fichier pour vous permettre d’exécuter le projet et d’utiliser votre complément immédiatement. Vous pouvez modifier ces paramètres à tout moment en modifiant le fichier XML.|
|Projet d’application web|Contient les pages de contenu de votre complément, notamment tous les fichiers et références de fichiers dont vous avez besoin pour développer des pages HTML et JavaScript compatibles avec Office. Pendant que vous développez votre complément, Visual Studio héberge l’application web sur votre serveur IIS local. Lorsque vous êtes prêt à publier le complément, vous devez déployer ce projet d’application web sur un serveur web.|

### <a name="update-the-code"></a>Mise à jour du code

1. **MessageRead.html** spécifie le code HTML qui s’affichera dans le volet Office du complément. Dans **MessageRead.html**, remplacez l’élément `<body>` par les marques suivantes et enregistrez le fichier.
 
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

1. Ouvrez le fichier **MessageRead.js** à la racine du projet d’application web. Ce fichier spécifie le script pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

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

1. Ouvrez le fichier **MessageRead.css** à la racine du projet d’application web. Ce fichier spécifie les styles personnalisés pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

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

1. L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.

1. L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé. Remplacez-le par `My Office Add-in`.

1. L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par `My First Outlook add-in`.

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

1. À l’aide de Visual Studio, testez le complément Outlook que vous venez de créer en appuyant sur F5 ou en sélectionnant le bouton **Démarrer**. Le complément est hébergé localement sur IIS.

1. Dans la boîte de dialogue**Se connecter à un compte de messagerie Exchange**, entrez l’adresse de messagerie et mot de passe pour votre [compte Microsoft](https://account.microsoft.com/account), puis sélectionnez**Se connecter**. Lorsque la page de connexion Outlook.com s’ouvre dans un navigateur, connectez-vous à votre compte de courrier avec les mêmes informations d’identification que vous avez entrées précédemment.

    > [!NOTE]
    > Si la boîte de dialogue **Se connecter au compte de messagerie Exchange** vous invite à vous connecter à plusieurs reprises, l’authentification de base est peut-être désactivée pour les comptes sur votre client Microsoft 365. Pour tester ce complément, connectez-vous à l’aide d’un [compte Microsoft](https://account.microsoft.com/account) à la place.

1. Dans Outlook sur le web, sélectionnez ou ouvrez un message.

1. Dans le message, recherchez les points de suspension du menu de dépassement de capacité contenant le bouton du complément.

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le web avec les points de suspension mis en surbrillance](../images/quick-start-button-owa-1.png)

1. Dans le menu de dépassement de capacité, recherchez le bouton du complément.

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le web avec le bouton du complément mis en surbrillance](../images/quick-start-button-owa-2.png)

1. Cliquez sur le bouton pour ouvrir le volet Office du complément.

    ![Capture d’écran du volet Office du complément dans Outlook sur le web, affichant les propriétés des messages](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > Si le volet Office n’est pas chargé, essayez de l’ouvrir dans un navigateur sur le même ordinateur.

### <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé votre premier complément de volet de tâches Outlook ! Ensuite, en savoir plus sur la [création de compléments Office avec Visual Studio](../develop/develop-add-ins-visual-studio.md).

---
