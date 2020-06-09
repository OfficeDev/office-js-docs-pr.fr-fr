---
title: 'Didacticiel : créer un complément de composition de message Outlook'
description: Dans ce didacticiel, vous allez créer un complément Outlook qui insère des informations GitHub dans le corps d'un nouveau message.
ms.date: 12/10/2019
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 66ce5aa01cc8ef82d8a4bd0030c9e94a9433daad
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611056"
---
# <a name="tutorial-build-a-message-compose-outlook-add-in"></a><span data-ttu-id="f6865-103">Didacticiel : créer un complément de composition de message Outlook</span><span class="sxs-lookup"><span data-stu-id="f6865-103">Tutorial: Build a message compose Outlook add-in</span></span>

<span data-ttu-id="f6865-104">Ce didacticiel vous apprend à créer un complément Outlook qui peut être utilisé pour dans le mode composer un message pour insérer du contenu dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6865-104">This tutorial teaches you how to build an Outlook add-in that can be used in message compose mode to insert content into the body of a message.</span></span>

<span data-ttu-id="f6865-105">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="f6865-105">In this tutorial, you will:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="f6865-106">Créer un projet de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="f6865-106">Create an Outlook add-in project</span></span>
> * <span data-ttu-id="f6865-107">Définir des boutons qui s’afficheront dans la fenêtre composer un message</span><span class="sxs-lookup"><span data-stu-id="f6865-107">Define buttons that will render in the compose message window</span></span>
> * <span data-ttu-id="f6865-108">Implémenter une expérience de première exécution qui collecte des informations de l’utilisateur et extrait les données à partir d’un service externe</span><span class="sxs-lookup"><span data-stu-id="f6865-108">Implement a first-run experience that collects information from the user and fetches data from an external service</span></span>
> * <span data-ttu-id="f6865-109">Implémenter un bouton de l’interface utilisateur qui appelle une fonction</span><span class="sxs-lookup"><span data-stu-id="f6865-109">Implement a UI-less button that invokes a function</span></span>
> * <span data-ttu-id="f6865-110">Implémenter un volet des tâches qui insère du contenu dans le corps d’un message</span><span class="sxs-lookup"><span data-stu-id="f6865-110">Implement a task pane that inserts content into the body of a message</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f6865-111">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="f6865-111">Prerequisites</span></span>

- <span data-ttu-id="f6865-112">[Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="f6865-112">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

- <span data-ttu-id="f6865-113">La dernière version de[Yeoman](https://github.com/yeoman/yo) et de [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="f6865-113">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="f6865-114">Même si vous avez précédemment installé le générateur Yeoman, nous vous recommandons de mettre à jour votre package vers la dernière version de npm.</span><span class="sxs-lookup"><span data-stu-id="f6865-114">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

- <span data-ttu-id="f6865-115">Outlook 2016 ou versions ultérieures sur Windows (connecté à un compte Office 365) ou Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="f6865-115">Outlook 2016 or later on Windows (connected to an Office 365 account) or Outlook on the web</span></span>

- <span data-ttu-id="f6865-116">Un compte[GitHub](https://www.github.com) </span><span class="sxs-lookup"><span data-stu-id="f6865-116">A [GitHub](https://www.github.com) account</span></span>

## <a name="setup"></a><span data-ttu-id="f6865-117">Configuration</span><span class="sxs-lookup"><span data-stu-id="f6865-117">Setup</span></span>

<span data-ttu-id="f6865-118">Le complément que vous allez créer dans ce didacticiel lit les[gists](https://gist.github.com) à partir du compte utilisateur GitHub et ajoute le gist sélectionné dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6865-118">The add-in that you'll create in this tutorial will read [gists](https://gist.github.com) from the user's GitHub account and add the selected gist to the body of a message.</span></span> <span data-ttu-id="f6865-119">Procédez comme suit pour créer deux nouveaux gists que vous pouvez utiliser pour tester le complément que vous allez créer.</span><span class="sxs-lookup"><span data-stu-id="f6865-119">Complete the following steps to create two new gists that you can use to test the add-in you're going to build.</span></span>

1. <span data-ttu-id="f6865-120">[Connectez-vous à GitHub](https://github.com/login).</span><span class="sxs-lookup"><span data-stu-id="f6865-120">[Login to GitHub](https://github.com/login).</span></span>

1. <span data-ttu-id="f6865-121">[Créer une nouveau gist](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="f6865-121">[Create a new gist](https://gist.github.com).</span></span>

    - <span data-ttu-id="f6865-122">Dans la zone**Description gist...**, entrez **Hello World Markdown**.</span><span class="sxs-lookup"><span data-stu-id="f6865-122">In the **Gist description...** field, enter **Hello World Markdown**.</span></span>

    - <span data-ttu-id="f6865-123">Dans la zone\*\*Nom de fichier incluant l’extension... \*\*, entrez **test.md**.</span><span class="sxs-lookup"><span data-stu-id="f6865-123">In the **Filename including extension...** field, enter **test.md**.</span></span>

    - <span data-ttu-id="f6865-124">Ajouter la démarque suivante à la zone de texte multiligne :</span><span class="sxs-lookup"><span data-stu-id="f6865-124">Add the following markdown to the multiline textbox:</span></span>

        ```markdown
        # Hello World

        This is content converted from Markdown!

        Here's a JSON sample:

          ```json
          {
            "foo": "bar"
          }
          ```
        ```

    - <span data-ttu-id="f6865-125">Sélectionnez le bouton**créer un gist public**.</span><span class="sxs-lookup"><span data-stu-id="f6865-125">Select the **Create public gist** button.</span></span>

1. <span data-ttu-id="f6865-126">[Créer un nouveau gist](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="f6865-126">[Create another new gist](https://gist.github.com).</span></span>

    - <span data-ttu-id="f6865-127">Dans la zone**Description gist...**, entrez **Hello World Html**.</span><span class="sxs-lookup"><span data-stu-id="f6865-127">In the **Gist description...** field, enter **Hello World Html**.</span></span>

    - <span data-ttu-id="f6865-128">Dans la zone**Nom de fichier incluant l’extension...**, entrez **test.html**.</span><span class="sxs-lookup"><span data-stu-id="f6865-128">In the **Filename including extension...** field, enter **test.html**.</span></span>

    - <span data-ttu-id="f6865-129">Ajouter la démarque suivante à la zone de texte multiligne :</span><span class="sxs-lookup"><span data-stu-id="f6865-129">Add the following markdown to the multiline textbox:</span></span>

        ```HTML
        <html>
          <head>
            <style>
            h1 {
              font-family: Calibri;
            }
            </style>
          </head>
          <body>
            <h1>Hello World!</h1>
            <p>This is a test</p>
          </body>
        </html>
        ```

    - <span data-ttu-id="f6865-130">Sélectionnez le bouton**créer un gist public**.</span><span class="sxs-lookup"><span data-stu-id="f6865-130">Select the **Create public gist** button.</span></span>

## <a name="create-an-outlook-add-in-project"></a><span data-ttu-id="f6865-131">Créer un projet de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="f6865-131">Create an Outlook add-in project</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="f6865-132">**Sélectionnez un type de projet** - `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="f6865-132">**Choose a project type** - `Office Add-in Task Pane project`</span></span>

    - <span data-ttu-id="f6865-133">**Sélectionnez un type de script** - `Javascript`</span><span class="sxs-lookup"><span data-stu-id="f6865-133">**Choose a script type** - `Javascript`</span></span>

    - <span data-ttu-id="f6865-134">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="f6865-134">**What do you want to name your add-in?**</span></span> - `Git the gist`

    - <span data-ttu-id="f6865-135">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="f6865-135">**Which Office client application would you like to support?**</span></span> - `Outlook`

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yeoman-prompts-2.png)
    
    <span data-ttu-id="f6865-137">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="f6865-137">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="f6865-138">Accédez au registre racine du projet.</span><span class="sxs-lookup"><span data-stu-id="f6865-138">Navigate to the root directory of the project.</span></span>

    ```command&nbsp;line
    cd "Git the gist"
    ```

1. <span data-ttu-id="f6865-139">Ce complément utilise les bibliothèques suivantes :</span><span class="sxs-lookup"><span data-stu-id="f6865-139">This add-in will use the following libraries:</span></span>

    - <span data-ttu-id="f6865-140">Bibliothèque [Showdown](https://github.com/showdownjs/showdown) pour convertir Markdown en HTML.</span><span class="sxs-lookup"><span data-stu-id="f6865-140">[Showdown](https://github.com/showdownjs/showdown) library to convert Markdown to HTML</span></span>
    - <span data-ttu-id="f6865-141">Bibliothèque [URI.js](https://github.com/medialize/URI.js) pour créer des URL relatives.</span><span class="sxs-lookup"><span data-stu-id="f6865-141">[URI.js](https://github.com/medialize/URI.js) library to build relative URLs.</span></span>
    - <span data-ttu-id="f6865-142">Bibliothèque [jQuery](https://jquery.com/) pour simplifier les interactions DOM.</span><span class="sxs-lookup"><span data-stu-id="f6865-142">[jquery](https://jquery.com/) library to simplify DOM interactions.</span></span>

     <span data-ttu-id="f6865-143">Pour installer ces outils pour votre projet, exécutez la commande suivante dans le répertoire racine du projet :</span><span class="sxs-lookup"><span data-stu-id="f6865-143">To install these tools for your project, run the following command in the root directory of the project:</span></span>

    ```command&nbsp;line
    npm install showdown urijs jquery --save
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="f6865-144">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="f6865-144">Update the manifest</span></span>

<span data-ttu-id="f6865-145">Le manifeste d’un complément contrôle la manière dont il apparaît dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f6865-145">The manifest for an add-in controls how it appears in Outlook.</span></span> <span data-ttu-id="f6865-146">Il définit la façon dont le complément est affiché dans la liste des compléments, les boutons qui apparaissent sur le ruban, et il configure les URL pour les fichiers HTML et JavaScript utilisés par le complément.</span><span class="sxs-lookup"><span data-stu-id="f6865-146">It defines the way the add-in appears in the add-in list and the buttons that appear on the ribbon, and it sets the URLs for the HTML and JavaScript files used by the add-in.</span></span>

#### <a name="specify-basic-information"></a><span data-ttu-id="f6865-147">Spécifiez les informations de base</span><span class="sxs-lookup"><span data-stu-id="f6865-147">Specify basic information</span></span>

<span data-ttu-id="f6865-148">Apportez les mises à jour suivantes dans le fichier **manifest.xml** pour spécifier les informations de base sur le complément :</span><span class="sxs-lookup"><span data-stu-id="f6865-148">Make the following updates in the **manifest.xml** file to specify some basic information about the add-in:</span></span>

1. <span data-ttu-id="f6865-149">Recherchez l’élément `ProviderName`et remplacez la valeur par défaut par le nom de votre société.</span><span class="sxs-lookup"><span data-stu-id="f6865-149">Locate the `ProviderName` element and replace the default value with your company name.</span></span>

    ```xml
    <ProviderName>Contoso</ProviderName>
    ```
1. <span data-ttu-id="f6865-150">Recherchez l’`Description` élément, remplacez la valeur par défaut avec une description du complément et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f6865-150">Locate the `Description` element, replace the default value with a description of the add-in, and save the file.</span></span>

    ```xml
    <Description DefaultValue="Allows users to access their GitHub gists."/>
    ```

#### <a name="test-the-generated-add-in"></a><span data-ttu-id="f6865-151">Tester le complément généré</span><span class="sxs-lookup"><span data-stu-id="f6865-151">Test the generated add-in</span></span>

<span data-ttu-id="f6865-152">Avant d’aller plus loin, nous allons tester le complément base créé par le générateur pour confirmer que le projet est correctement configuré.</span><span class="sxs-lookup"><span data-stu-id="f6865-152">Before going any further, let's test the basic add-in that the generator created to confirm that the project is set up correctly.</span></span>

> [!NOTE]
> <span data-ttu-id="f6865-153">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="f6865-153">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f6865-154">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f6865-154">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

1. <span data-ttu-id="f6865-155">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f6865-155">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="f6865-156">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="f6865-156">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="f6865-157">Suivez les instructions disponibles dans [Chargement indépendant de compléments Outlook à des fins de test](../outlook/sideload-outlook-add-ins-for-testing.md) pour charger le fichier **manifest.xml** situé dans le répertoire racine du projet.</span><span class="sxs-lookup"><span data-stu-id="f6865-157">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the **manifest.xml** file that's located in the root directory of the project.</span></span>

1. <span data-ttu-id="f6865-158">Dans Outlook, ouvrez un message existant et sélectionnez le bouton **Afficher le volet Office**.</span><span class="sxs-lookup"><span data-stu-id="f6865-158">In Outlook, open an existing message and select the **Show Taskpane** button.</span></span> <span data-ttu-id="f6865-159">Si tout est configuré correctement, le volet des tâches va s’ouvrir et afficher la page d’accueil du complément.</span><span class="sxs-lookup"><span data-stu-id="f6865-159">If everything's been set up correctly, the task pane will open and render the add-in's welcome page.</span></span>

    ![Capture d’écran du bouton et du volet des tâches ajoutés par l’exemple](../images/button-and-pane.png)

## <a name="define-buttons"></a><span data-ttu-id="f6865-161">Définir des boutons</span><span class="sxs-lookup"><span data-stu-id="f6865-161">Define buttons</span></span>

<span data-ttu-id="f6865-162">À présent que vous avez vérifié que le complément base fonctionne, vous pouvez le personnaliser pour ajouter davantage de fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="f6865-162">Now that you've verified the base add-in works, you can customize it to add more functionality.</span></span> <span data-ttu-id="f6865-163">Par défaut, le manifeste définit uniquement les boutons de la fenêtre de lecture de message.</span><span class="sxs-lookup"><span data-stu-id="f6865-163">By default, the manifest only defines buttons for the read message window.</span></span> <span data-ttu-id="f6865-164">Nous allons mettre à jour le manifeste pour supprimer les boutons de la fenêtre de lecture de message et définir deux nouveaux boutons pour la fenêtre composer un message :</span><span class="sxs-lookup"><span data-stu-id="f6865-164">Let's update the manifest to remove the buttons from the read message window and define two new buttons for the compose message window:</span></span>

- <span data-ttu-id="f6865-165">**Insérer un gist**: bouton qui ouvre un le volet des tâches</span><span class="sxs-lookup"><span data-stu-id="f6865-165">**Insert gist**: a button that opens a task pane</span></span>

- <span data-ttu-id="f6865-166">**Insérer gist par défaut**: bouton qui appelle une fonction</span><span class="sxs-lookup"><span data-stu-id="f6865-166">**Insert default gist**: a button that invokes a function</span></span>

### <a name="remove-the-messagereadcommandsurface-extension-point"></a><span data-ttu-id="f6865-167">Supprimer le point d’extension MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="f6865-167">Remove the MessageReadCommandSurface extension point</span></span>

<span data-ttu-id="f6865-168">Ouvrir le fichier **manifest.xml** et rechercher l’`ExtensionPoint` élément avec un type `MessageReadCommandSurface`.</span><span class="sxs-lookup"><span data-stu-id="f6865-168">Open the **manifest.xml** file and locate the `ExtensionPoint` element with type `MessageReadCommandSurface`.</span></span> <span data-ttu-id="f6865-169">Supprimer cet `ExtensionPoint` élément (y compris sa balise de fermeture) pour supprimer les boutons de la fenêtre de lecture de message.</span><span class="sxs-lookup"><span data-stu-id="f6865-169">Delete this `ExtensionPoint` element (including its closing tag) to remove the buttons from the read message window.</span></span>

### <a name="add-the-messagecomposecommandsurface-extension-point"></a><span data-ttu-id="f6865-170">Supprimer le point d’extension MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="f6865-170">Add the MessageComposeCommandSurface extension point</span></span>

<span data-ttu-id="f6865-171">Recherchez la ligne dans le manifeste qui lit `</DesktopFormFactor>`.</span><span class="sxs-lookup"><span data-stu-id="f6865-171">Locate the line in the manifest that reads `</DesktopFormFactor>`.</span></span> <span data-ttu-id="f6865-172">Situé immédiatement avant cette ligne, insérez le balisage XML suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-172">Immediately before this line, insert the following XML markup.</span></span> <span data-ttu-id="f6865-173">Notez les points suivants concernant ce balisage :</span><span class="sxs-lookup"><span data-stu-id="f6865-173">Note the following about this markup:</span></span>

- <span data-ttu-id="f6865-174">L’élément `ExtensionPoint` avec `xsi:type="MessageComposeCommandSurface"` indique que vous définissez des boutons à ajouter à la fenêtre de composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6865-174">The `ExtensionPoint` with `xsi:type="MessageComposeCommandSurface"` indicates that you're defining buttons to add to the compose message window.</span></span>

- <span data-ttu-id="f6865-175">En utilisant un élément `OfficeTab` avec `id="TabDefault"`, vous indiquez que vous voulez ajouter des boutons à l’onglet par défaut dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="f6865-175">By using an `OfficeTab` element with `id="TabDefault"`, you're indicating you want to add the buttons to the default tab on the ribbon.</span></span>

- <span data-ttu-id="f6865-176">L’élément `Group` définit le regroupement de nouveaux boutons, avec une étiquette définie par la ressource `groupLabel`.</span><span class="sxs-lookup"><span data-stu-id="f6865-176">The `Group` element defines the grouping for the new buttons, with a label set by the `groupLabel` resource.</span></span>

- <span data-ttu-id="f6865-177">Le premier élément `Control` contient un élément `Action` avec `xsi:type="ShowTaskPane"`, afin que le bouton ouvre un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="f6865-177">The first `Control` element contains an `Action` element with `xsi:type="ShowTaskPane"`, so this button opens a task pane.</span></span>

- <span data-ttu-id="f6865-178">Le deuxième élément `Control` contient un élément `Action` avec `xsi:type="ExecuteFunction"`, afin que le bouton appelle une fonction JavaScript contenue dans le fichier de fonction.</span><span class="sxs-lookup"><span data-stu-id="f6865-178">The second `Control` element contains an `Action` element with `xsi:type="ExecuteFunction"`, so this button invokes a JavaScript function contained in the function file.</span></span>

```xml
<!-- Message Compose -->
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgComposeCmdGroup">
      <Label resid="GroupLabel"/>
      <Control xsi:type="Button" id="msgComposeInsertGist">
        <Label resid="TaskpaneButton.Label"/>
        <Supertip>
          <Title resid="TaskpaneButton.Title"/>
          <Description resid="TaskpaneButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="Taskpane.Url"/>
        </Action>
      </Control>
      <Control xsi:type="Button" id="msgComposeInsertDefaultGist">
        <Label resid="FunctionButton.Label"/>
        <Supertip>
          <Title resid="FunctionButton.Title"/>
          <Description resid="FunctionButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
          <FunctionName>insertDefaultGist</FunctionName>
        </Action>
      </Control>
    </Group>
  </OfficeTab>
</ExtensionPoint>
```

### <a name="update-resources-in-the-manifest"></a><span data-ttu-id="f6865-179">Ressources de mise à jour dans le fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="f6865-179">Update resources in the manifest</span></span>

<span data-ttu-id="f6865-180">Le code précédent fait référence à des étiquettes, des info-bulles et des URL que vous devez définir avant que le manifeste ne soit valide.</span><span class="sxs-lookup"><span data-stu-id="f6865-180">The previous code references labels, tooltips, and URLs that you need to define before the manifest will be valid.</span></span> <span data-ttu-id="f6865-181">Vous devez spécifier ces informations dans la section `Resources` du manifeste.</span><span class="sxs-lookup"><span data-stu-id="f6865-181">You'll specify this information in the `Resources` section of the manifest.</span></span>

1. <span data-ttu-id="f6865-182">Recherchez l’élément `Resources` dans le fichier manifeste, puis supprimez entièrement l’élément (balise de fermeture comprise).</span><span class="sxs-lookup"><span data-stu-id="f6865-182">Locate the `Resources` element in the manifest file and delete the entire element (including its closing tag).</span></span>

1. <span data-ttu-id="f6865-183">À ce même emplacement, ajoutez le balisage suivant pour remplacer l’élément `Resources` que vous venez de supprimer :</span><span class="sxs-lookup"><span data-stu-id="f6865-183">In that same location, add the following markup to replace the `Resources` element you just removed:</span></span>

    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Git the gist"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Insert gist"/>
        <bt:String id="TaskpaneButton.Title" DefaultValue="Insert gist"/>
        <bt:String id="FunctionButton.Label" DefaultValue="Insert default gist"/>
        <bt:String id="FunctionButton.Title" DefaultValue="Insert default gist"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Displays a list of your gists and allows you to insert their contents into the current message."/>
        <bt:String id="FunctionButton.Tooltip" DefaultValue="Inserts the content of the gist you mark as default into the current message."/>
      </bt:LongStrings>
    </Resources>
    ```

1. <span data-ttu-id="f6865-184">Enregistrez les modifications dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="f6865-184">Save your changes to the manifest.</span></span>

### <a name="reinstall-the-add-in"></a><span data-ttu-id="f6865-185">Réinstallez le complément.</span><span class="sxs-lookup"><span data-stu-id="f6865-185">Reinstall the add-in</span></span>

<span data-ttu-id="f6865-186">Étant donné que vous avez installé le complément à partir d’un fichier, vous devez le réinstaller afin que les modifications soient prises en compte.</span><span class="sxs-lookup"><span data-stu-id="f6865-186">Since you previously installed the add-in from a file, you must reinstall it in order for the manifest changes to take effect.</span></span>

1. <span data-ttu-id="f6865-187">Suivez les instructions de [Charger les compléments Outlook pour les tests](../outlook/sideload-outlook-add-ins-for-testing.md) pour localiser la section **Compléments personnalisés** en bas de la boîte de dialogue**Mes compléments** .</span><span class="sxs-lookup"><span data-stu-id="f6865-187">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to locate the **Custom add-ins** section at the bottom of the **My add-ins** dialog box.</span></span>

1. <span data-ttu-id="f6865-188">Cliquez sur le bouton \*\*... \*\* en regard de l’entrée **Git the Gist**, puis sélectionnez **Supprimer**.</span><span class="sxs-lookup"><span data-stu-id="f6865-188">Select the **...** button next to the **Git the gist** entry and then choose **Remove**.</span></span>

1. <span data-ttu-id="f6865-189">Fermer la fenêtre**Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="f6865-189">Close the **My add-ins** window.</span></span>

1. <span data-ttu-id="f6865-190">Le bouton personnalisé doit disparaître du ruban temporairement.</span><span class="sxs-lookup"><span data-stu-id="f6865-190">The custom button should disappear from the ribbon momentarily.</span></span>

1. <span data-ttu-id="f6865-191">Suivez les instructions de [Charger compléments Outlook pour les tests](../outlook/sideload-outlook-add-ins-for-testing.md) pour réinstaller le complément à l’aide du fichier mis à jour **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="f6865-191">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to reinstall the add-in using the updated **manifest.xml** file.</span></span>

<span data-ttu-id="f6865-192">Une fois le complément réinstallé, vous pouvez vérifier qu’il a été correctement installé en consultant les commandes **Insérer gist** et **Insérer gist par défaut** dans le fenêtre de composition du message.</span><span class="sxs-lookup"><span data-stu-id="f6865-192">After you've reinstalled the add-in, you can verify that it installed successfully by checking for the commands **Insert gist** and **Insert default gist** in a compose message window.</span></span> <span data-ttu-id="f6865-193">Notez que rien ne se produit si vous sélectionnez un des ces éléments, car vous n’avez pas encore terminé de générer ce complément.</span><span class="sxs-lookup"><span data-stu-id="f6865-193">Note that nothing will happen if you select either of these items, because you haven't yet finished building this add-in.</span></span>

- <span data-ttu-id="f6865-194">Si vous exécutez ce complément dans Outlook 2016 ou versions ultérieures sur Windows, vous devriez voir deux nouveaux boutons dans le ruban de la fenêtre de composition d’un message : **Insérer gist** et **Insérer gist par défaut**.</span><span class="sxs-lookup"><span data-stu-id="f6865-194">If you're running this add-in in Outlook 2016 or later on Windows, you should see two new buttons in the ribbon of the compose message window: **Insert gist** and **Insert default gist**.</span></span>

    ![Capture d’écran du ruban dans Outlook sur Windows avec boutons du complément mis en évidence](../images/add-in-buttons-in-windows.png)

- <span data-ttu-id="f6865-196">Si vous exécutez ce complément dans Outlook sur le web, vous devriez voir apparaître un nouveau bouton en bas de la fenêtre de composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6865-196">If you're running this add-in in Outlook on the web, you should see a new button at the bottom of the compose message window.</span></span> <span data-ttu-id="f6865-197">Sélectionnez ce bouton pour afficher les options **Insérer gist** et **Insérer gist par défaut**.</span><span class="sxs-lookup"><span data-stu-id="f6865-197">Select that button to see the options **Insert gist** and **Insert default gist**.</span></span>

    ![Capture d’écran du formulaire composer message dans Outlook sur le web avec le bouton complément et menu contextuel mis en évidence](../images/add-in-buttons-in-owa.png)

## <a name="implement-a-first-run-experience"></a><span data-ttu-id="f6865-199">Mettre en œuvre une expérience de première exécution</span><span class="sxs-lookup"><span data-stu-id="f6865-199">Implement a first-run experience</span></span>

<span data-ttu-id="f6865-200">Ce complément doit être en mesure de lire les gists du compte d’utilisateur GitHub et d’identifier lequel l’utilisateur a choisi en tant que gist par défaut.</span><span class="sxs-lookup"><span data-stu-id="f6865-200">This add-in needs to be able to read gists from the user's GitHub account and identify which one the user has chosen as the default gist.</span></span> <span data-ttu-id="f6865-201">Pour atteindre ces objectifs, le complément doit inviter l’utilisateur à fournir son nom d’utilisateur GitHub et choisir un gist par défaut parmi leur collection de gists existants.</span><span class="sxs-lookup"><span data-stu-id="f6865-201">In order to achieve these goals, the add-in must prompt the user to provide their GitHub username and choose a default gist from their collection of existing gists.</span></span> <span data-ttu-id="f6865-202">Suivez les étapes décrites dans cette section pour implémenter une expérience de première exécution qui affiche une boîte de dialogue pour collecter ces informations à partir de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f6865-202">Complete the steps in this section to implement a first-run experience that will display a dialog to collect this information from the user.</span></span>

### <a name="collect-data-from-the-user"></a><span data-ttu-id="f6865-203">Collecter les données d’un utilisateur</span><span class="sxs-lookup"><span data-stu-id="f6865-203">Collect data from the user</span></span>

<span data-ttu-id="f6865-204">Commençons par créer l’interface utilisateur pour la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="f6865-204">Let's start by creating the UI for the dialog itself.</span></span> <span data-ttu-id="f6865-205">Dans le dossier **./src**, créez un sous-dossier nommé **settings**.</span><span class="sxs-lookup"><span data-stu-id="f6865-205">Within the **./src** folder, create a new subfolder named **settings**.</span></span> <span data-ttu-id="f6865-206">Dans le dossier **./src/settings**, créez un fichier nommé **dialog.html** et ajoutez le balisage suivant pour définir un formulaire très simple avec une entrée de texte pour un nom d’utilisateur GitHub et une liste vide pour gists qui sera renseignée via JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f6865-206">In the **./src/settings** folder, create a file named **dialog.html**, and add the following markup to define a very basic form with a text input for a GitHub username and an empty list for gists that'll be populated via JavaScript.</span></span>

```html
<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <title>Settings</title>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

  <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

  <!-- Template styles -->
  <link href="dialog.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l">
  <main>
    <section class="ms-font-m ms-fontColor-neutralPrimary">
      <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--Info"></i>
          </div>
          <div class="ms-MessageBar-text">
            Oops! It looks like you haven't configured <strong>Git the gist</strong> yet.
            <br/>
            Please configure your GitHub username and select a default gist, then try that action again!
          </div>
        </div>
      </div>
      <div class="ms-font-xxl">Settings</div>
      <div class="ms-Grid">
        <div class="ms-Grid-row">
          <div class="ms-TextField">
            <label class="ms-Label">GitHub Username</label>
            <input class="ms-TextField-field" id="github-user" type="text" value="" placeholder="Please enter your GitHub username">
          </div>
        </div>
        <div class="error-display ms-Grid-row">
          <div class="ms-font-l ms-fontWeight-semibold">An error occurred:</div>
          <pre><code id="error-text"></code></pre>
        </div>
        <div class="gist-list-container ms-Grid-row">
          <div class="list-title ms-font-xl ms-fontWeight-regular">Choose Default Gist</div>
          <form>
            <div id="gist-list">
            </div>
          </form>
        </div>
      </div>
      <div class="ms-Dialog-actions">
        <div class="ms-Dialog-actionsRight">
          <button class="ms-Dialog-action ms-Button ms-Button--primary" id="settings-done" disabled>
            <span class="ms-Button-label">Done</span>
          </button>
        </div>
      </div>
    </section>
  </main>
  <script type="text/javascript" src="../../node_modules/core-js/client/core.js"></script>
  <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../helpers/gist-api.js"></script>
  <script type="text/javascript" src="dialog.js"></script>
</body>

</html>
```

<span data-ttu-id="f6865-207">Ensuite, créez un fichier dans le dossier **./src/settings** nommé **dialog.css** et ajoutez le code suivant pour spécifier les styles utilisés par **dialog.html**.</span><span class="sxs-lookup"><span data-stu-id="f6865-207">Next, create a file in the **./src/settings** folder named **dialog.css**, and add the following code to specify the styles that are used by **dialog.html**.</span></span>

```CSS
section {
  margin: 10px 20px;
}

.not-configured-warning {
  display: none;
}

.error-display {
  display: none;
}

.gist-list-container {
  margin: 10px -8px;
  display: none;
}

.list-title {
  border-bottom: 1px solid #a6a6a6;
  padding-bottom: 5px;
}

ul {
  margin-top: 10px;
}

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}
```

<span data-ttu-id="f6865-208">Maintenant que vous avez défini la boîte de dialogue interface utilisateur, vous pouvez écrire du code pour l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="f6865-208">Now that you've defined the dialog UI, you can write the code that makes it actually do something.</span></span> <span data-ttu-id="f6865-209">Créez un fichier dans le dossier **./src/settings** nommé **dialog.js** et ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-209">Create a file in the **./src/settings** folder named **dialog.js** and add the following code.</span></span> <span data-ttu-id="f6865-210">Notez que ce code utilise jQuery pour enregistrer des événements et la fonction `messageParent` pour renvoyer les choix de l’utilisateur à l’appelant.</span><span class="sxs-lookup"><span data-stu-id="f6865-210">Note that this code uses jQuery to register events and uses the `messageParent` function to send the user's choices back to the caller.</span></span>

```js
(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      if (window.location.search) {
        // Check if warning should be displayed.
        var warn = getParameterByName('warn');
        if (warn) {
          $('.not-configured-warning').show();
        } else {
          // See if the config values were passed.
          // If so, pre-populate the values.
          var user = getParameterByName('gitHubUserName');
          var gistId = getParameterByName('defaultGistId');

          $('#github-user').val(user);
          loadGists(user, function(success){
            if (success) {
              $('.ms-ListItem').removeClass('is-selected');
              $('input').filter(function() {
                return this.value === gistId;
              }).addClass('is-selected').attr('checked', 'checked');
              $('#settings-done').removeAttr('disabled');
            }
          });
        }
      }

      // When the GitHub username changes,
      // try to load gists.
      $('#github-user').on('change', function(){
        $('#gist-list').empty();
        var ghUser = $('#github-user').val();
        if (ghUser.length > 0) {
          loadGists(ghUser);
        }
      });

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', function() {
        var settings = {};

        settings.gitHubUserName = $('#github-user').val();

        var selectedGist = $('.ms-ListItem.is-selected');
        if (selectedGist) {
          settings.defaultGistId = selectedGist.val();

          sendMessage(JSON.stringify(settings));
        }
      });
    });
  };

  // Load gists for the user using the GitHub API
  // and build the list.
  function loadGists(user, callback) {
    getUserGists(user, function(gists, error){
      if (error) {
        $('.gist-list-container').hide();
        $('#error-text').text(JSON.stringify(error, null, 2));
        $('.error-display').show();
        if (callback) callback(false);
      } else {
        $('.error-display').hide();
        buildGistList($('#gist-list'), gists, onGistSelected);
        $('.gist-list-container').show();
        if (callback) callback(true);
      }
    });
  }

  function onGistSelected() {
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
    $('.not-configured-warning').hide();
    $('#settings-done').removeAttr('disabled');
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
```

#### <a name="update-webpack-config-settings"></a><span data-ttu-id="f6865-211">Mettre à jour les paramètres de configuration webapck</span><span class="sxs-lookup"><span data-stu-id="f6865-211">Update webpack config settings</span></span>

<span data-ttu-id="f6865-212">Enfin, ouvrez le fichier **webpack.config.js** situé dans le répertoire racine du projet et procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="f6865-212">Finally, open the file **webpack.config.js** file in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="f6865-213">Recherchez l’objet `entry` dans l’objet `config` et ajoutez une nouvelle entrée pour `dialog`.</span><span class="sxs-lookup"><span data-stu-id="f6865-213">Locate the `entry` object within the `config` object and add a new entry for `dialog`.</span></span>

    ```js
    dialog: "./src/settings/dialog.js"
    ```

    <span data-ttu-id="f6865-214">Lorsque c’est chose faite, le nouvel objet `entry` se présente comme suit :</span><span class="sxs-lookup"><span data-stu-id="f6865-214">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      dialog: "./src/settings/dialog.js"
    },
    ```
  
2. <span data-ttu-id="f6865-215">Recherchez la matrice `plugins` dans l’objet `config` et ajoutez ces deux nouveaux objets à la fin de cette matrice.</span><span class="sxs-lookup"><span data-stu-id="f6865-215">Locate the `plugins` array within the `config` object and add these two new objects to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "dialog.html",
      template: "./src/settings/dialog.html",
      chunks: ["polyfill", "dialog"]
    }),
    new CopyWebpackPlugin([
      {
        to: "dialog.css",
        from: "./src/settings/dialog.css"
      }
    ])
    ```

    <span data-ttu-id="f6865-216">Lorsque c’est chose faite, la nouvelle matrice `plugins` se présente comme suit :</span><span class="sxs-lookup"><span data-stu-id="f6865-216">After you've done this, the new `plugins` array will look like this:</span></span>

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
      {
        to: "taskpane.css",
        from: "./src/taskpane/taskpane.css"
      }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/settings/dialog.html",
        chunks: ['polyfill', 'dialog']
      }),
      new CopyWebpackPlugin([
      {
        to: "dialog.css",
        from: "./src/settings/dialog.css"
      }
      ])
    ],
    ```

3. <span data-ttu-id="f6865-217">Si le serveur web est en cours d’exécution, fermez la fenêtre de commande de nœud.</span><span class="sxs-lookup"><span data-stu-id="f6865-217">If the web server is running, close the node command window.</span></span>

4. <span data-ttu-id="f6865-218">Exécutez la commande suivante pour regénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="f6865-218">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

5. <span data-ttu-id="f6865-219">Entrez la commande suivante pour démarrer le serveur web.</span><span class="sxs-lookup"><span data-stu-id="f6865-219">Run the following command to start the web server.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

### <a name="fetch-data-from-github"></a><span data-ttu-id="f6865-220">Récupérer des données à partir de GitHub</span><span class="sxs-lookup"><span data-stu-id="f6865-220">Fetch data from GitHub</span></span>

<span data-ttu-id="f6865-221">Le fichier**dialog.js** que vous venez de créer spécifie que le complément doit charger les gists lorsque l’`change` événement se déclenche pour le champ nom d’utilisateur GitHub.</span><span class="sxs-lookup"><span data-stu-id="f6865-221">The **dialog.js** file you just created specifies that the add-in should load gists when the `change` event fires for the GitHub username field.</span></span> <span data-ttu-id="f6865-222">Pour récupérer les gists de l’utilisateur à partir de GitHub, vous utiliserez le [API GitHub Gists](https://developer.github.com/v3/gists/).</span><span class="sxs-lookup"><span data-stu-id="f6865-222">To retrieve the user's gists from GitHub, you'll use the [GitHub Gists API](https://developer.github.com/v3/gists/).</span></span>

<span data-ttu-id="f6865-223">Dans le dossier **./src**, créez un nouveau sous-dossier nommé **helpers**.</span><span class="sxs-lookup"><span data-stu-id="f6865-223">Within the **./src** folder, create a new subfolder named **helpers**.</span></span> <span data-ttu-id="f6865-224">Dans le dossier **./src/helpers**, créez un fichier nommé **gist-api.js** et ajoutez le code suivant pour récupérer les gists de l’utilisateur à partir de GitHub et créer la liste des gists.</span><span class="sxs-lookup"><span data-stu-id="f6865-224">In the **./src/helpers** folder, create a file named **gist-api.js**, and add the following code to retrieve the user's gists from GitHub and build the list of gists.</span></span>

```js
function getUserGists(user, callback) {
  var requestUrl = 'https://api.github.com/users/' + user + '/gists';

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gists){
    callback(gists);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function(gist) {

    var listItem = $('<div/>')
      .appendTo(parent);

    var radioItem = $('<input>')
      .addClass('ms-ListItem')
      .addClass('is-selectable')
      .attr('type', 'radio')
      .attr('name', 'gists')
      .attr('tabindex', 0)
      .val(gist.id)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-primaryText')
      .text(gist.description)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-secondaryText')
      .text(' - ' + buildFileList(gist.files))
      .appendTo(listItem);

    var updated = new Date(gist.updated_at);

    var desc = $('<span/>')
      .addClass('ms-ListItem-tertiaryText')
      .text(' - Last updated ' + updated.toLocaleString())
      .appendTo(listItem);

    listItem.on('click', clickFunc);
  });  
}

function buildFileList(files) {

  var fileList = '';

  for (var file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ', ';
      }

      fileList = fileList + files[file].filename + ' (' + files[file].language + ')';
    }
  }

  return fileList;
}
```

> [!NOTE]
> <span data-ttu-id="f6865-225">Vous avez sans doute remarqué qu’il n’existe pas de bouton pour appeler la boîte de dialogue Paramètres.</span><span class="sxs-lookup"><span data-stu-id="f6865-225">You may have noticed that there's no button to invoke the settings dialog.</span></span> <span data-ttu-id="f6865-226">Au lieu de cela, le complément vérifie si cela a été configuré lorsque l’utilisateur sélectionne le bouton **Insérer gist par défaut** ou le bouton **Insérer gist**.</span><span class="sxs-lookup"><span data-stu-id="f6865-226">Instead, the add-in will check whether it has been configured when the user selects either the **Insert default gist** button or the **Insert gist** button.</span></span> <span data-ttu-id="f6865-227">Si le complément n'a pas encore été configuré, la boîte de dialogue Paramètres invite l’utilisateur à configurer avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="f6865-227">If the add-in has not yet been configured, the settings dialog will prompt the user to configure before proceeding.</span></span>

## <a name="implement-a-ui-less-button"></a><span data-ttu-id="f6865-228">Implémentation d’un bouton sans interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="f6865-228">Implement a UI-less button</span></span>

<span data-ttu-id="f6865-229">Le bouton **Insérer gist par défaut** de ce complément est un bouton sans interface utilisateur qui appelera une fonction JavaScript, plutôt que d’ouvrir un volet des tâches comme de nombreux boutons de complément le font.</span><span class="sxs-lookup"><span data-stu-id="f6865-229">This add-in's **Insert default gist** button is a UI-less button that will invoke a JavaScript function, rather than open a task pane like many add-in buttons do.</span></span> <span data-ttu-id="f6865-230">Lorsque l’utilisateur sélectionne le bouton **Insérer gist par défaut**, la fonction JavaScript correspondante vérifie si le complément a été configuré.</span><span class="sxs-lookup"><span data-stu-id="f6865-230">When the user selects the **Insert default gist** button, the corresponding JavaScript function will check whether the add-in has been configured.</span></span>

- <span data-ttu-id="f6865-231">Si le complément a déjà été configuré, la fonction chargera le contenu du gist que l’utilisateur a sélectionné par défaut et l’insérera dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f6865-231">If the add-in has already been configured, the function will load the content of the gist that the user has selected as the default and insert it into the body of the message.</span></span>

- <span data-ttu-id="f6865-232">Si le complément n'a pas encore été configuré, la boîte de dialogue Paramètres invitera l’utilisateur à fournir les informations nécessaires. </span><span class="sxs-lookup"><span data-stu-id="f6865-232">If the add-in hasn't yet been configured, then the settings dialog will prompt the user to provide the required information.</span></span>

### <a name="update-the-function-file-html"></a><span data-ttu-id="f6865-233">Mettre à jour le fichier de fonction (HTML)</span><span class="sxs-lookup"><span data-stu-id="f6865-233">Update the function file (HTML)</span></span>

<span data-ttu-id="f6865-234">Une fonction appelée par un bouton sans interface utilisateur doit être définie dans le fichier de fonction spécifié par l’élément `FunctionFile` dans le manifeste pour le facteur de formulaire correspondant.</span><span class="sxs-lookup"><span data-stu-id="f6865-234">A function that's invoked by a UI-less button must be defined in the file that's specified by the `FunctionFile` element in the manifest for the corresponding form factor.</span></span> <span data-ttu-id="f6865-235">Le manifeste de ce complément spécifie `https://localhost:3000/commands.html` comme fichier de fonction.</span><span class="sxs-lookup"><span data-stu-id="f6865-235">This add-in's manifest specifies `https://localhost:3000/commands.html` as the function file.</span></span>

<span data-ttu-id="f6865-236">Ouvrez le fichier **./src/commands/commands.html** et remplacez tout le contenu par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-236">Open the file **./src/commands/commands.html** and replace the entire contents with the following markup.</span></span>

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <script type="text/javascript" src="../node_modules/jquery/dist/jquery.js"></script>
    <script type="text/javascript" src="../node_modules/showdown/dist/showdown.min.js"></script>
    <script type="text/javascript" src="../node_modules/urijs/src/URI.min.js"></script>
    <script type="text/javascript" src="../src/helpers/addin-config.js"></script>
    <script type="text/javascript" src="../src/helpers/gist-api.js"></script>
</head>

<body>
  <!-- NOTE: The body is empty on purpose. Since functions in commands.js are
       invoked via a button, there is no UI to render. -->
</body>

</html>
```

### <a name="update-the-function-file-javascript"></a><span data-ttu-id="f6865-237">Mettre à jour le fichier de fonction (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f6865-237">Update the function file (JavaScript)</span></span>

<span data-ttu-id="f6865-238">Ouvrez le fichier **./src/commands/commands.js** et remplacez tout le contenu par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-238">Open the file **./src/commands/commands.js** and replace the entire contents with the following code.</span></span> <span data-ttu-id="f6865-239">Notez que si la `insertDefaultGist` fonction détermine que le complément n'a pas encore été configuré, elle ajoute le `?warn=1` paramètre à l’URL de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="f6865-239">Note that if the `insertDefaultGist` function determines the add-in has not yet been configured, it adds the `?warn=1` parameter to the dialog URL.</span></span> <span data-ttu-id="f6865-240">Cette opération permet à la boîte de dialogue Paramètres de restituer la barre des messages définie dans **./settings/dialog.html**, pour transmettre à l’utilisateur pourquoi il voit la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="f6865-240">Doing so makes the settings dialog render the message bar that's defined in **./settings/dialog.html**, to tell the user why they're seeing the dialog.</span></span>

```js
var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded.
Office.initialize = function (reason) {
};

// Add any UI-less function here.
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
    type: 'errorMessage',
    message: error
  }, function(result){
  });
}

var settingsDialog;

function insertDefaultGist(event) {

  config = getConfig();

  // Check if the add-in has been configured.
  if (config && config.defaultGistId) {
    // Get the default gist content and insert.
    try {
      getGist(config.defaultGistId, function(gist, error) {
        if (gist) {
          buildBodyContent(gist, function (content, error) {
            if (content) {
              Office.context.mailbox.item.body.setSelectedDataAsync(content,
                {coercionType: Office.CoercionType.Html}, function(result) {
                  event.completed();
              });
            } else {
              showError(error);
              event.completed();
            }
          });
        } else {
          showError(error);
          event.completed();
        }
      });
    } catch (err) {
      showError(err);
      event.completed();
    }

  } else {
    // Save the event object so we can finish up later.
    btnEvent = event;
    // Not configured yet, display settings dialog with
    // warn=1 to display warning.
    var url = new URI('../src/settings/dialog.html?warn=1').absoluteTo(window.location).toString();
    var dialogOptions = { width: 20, height: 40, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
      settingsDialog = result.value;
      settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
      settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
    });
  }
}

function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function(result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}

function getGlobal() {
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window :
    (typeof global !== "undefined") ? global :
    undefined;
}

var g = getGlobal();

// The add-in command functions need to be available in global scope.
g.insertDefaultGist = insertDefaultGist;
```

### <a name="create-a-file-to-manage-configuration-settings"></a><span data-ttu-id="f6865-241">Créer un fichier pour gérer les paramètres de configuration</span><span class="sxs-lookup"><span data-stu-id="f6865-241">Create a file to manage configuration settings</span></span>

<span data-ttu-id="f6865-242">Le fichier fonction HTML fait référence à un fichier nommé **addin-config.js**, qui n’existe pas encore.</span><span class="sxs-lookup"><span data-stu-id="f6865-242">The HTML function file references a file named **addin-config.js**, which doesn't yet exist.</span></span> <span data-ttu-id="f6865-243">Créez un fichier nommé **addin-config.js** dans le dossier **./src/helpers** et ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-243">Create a file named **addin-config.js** in the **./src/helpers** folder and add the following code.</span></span> <span data-ttu-id="f6865-244">Ce code utilise l’[objet RoamingSettings](/javascript/api/outlook/office.RoamingSettings) pour obtenir et définir les valeurs de configuration.</span><span class="sxs-lookup"><span data-stu-id="f6865-244">This code uses the [RoamingSettings object](/javascript/api/outlook/office.RoamingSettings) to get and set configuration values.</span></span>

```js
function getConfig() {
  var config = {};

  config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
  config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');

  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
  Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);

  Office.context.roamingSettings.saveAsync(callback);
}
```

### <a name="create-new-functions-to-process-gists"></a><span data-ttu-id="f6865-245">Créer de nouvelles fonctions pour traiter les gists</span><span class="sxs-lookup"><span data-stu-id="f6865-245">Create new functions to process gists</span></span>

<span data-ttu-id="f6865-246">Ensuite, ouvrez le fichier **./src/helpers/gist-api.js** et ajoutez les fonctions suivantes.</span><span class="sxs-lookup"><span data-stu-id="f6865-246">Next, open the **./src/helpers/gist-api.js** file and add the following functions.</span></span> <span data-ttu-id="f6865-247">Veuillez prendre en compte les éléments suivants:</span><span class="sxs-lookup"><span data-stu-id="f6865-247">Note the following:</span></span>

- <span data-ttu-id="f6865-248">Si le gist contient du HTML, le complément insère le code HTML tel quel dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f6865-248">If the gist contains HTML, the add-in will insert the HTML as-is into the body of the message.</span></span>

- <span data-ttu-id="f6865-249">Si le gist contient Markdown, le complément utilisera la bibliothèque[Showdown](https://github.com/showdownjs/showdown) pour convertir le Markdown en HTML, puis insérera le code HTML qui en résulte dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f6865-249">If the gist contains Markdown, the add-in will use the [Showdown](https://github.com/showdownjs/showdown) library to convert the Markdown to HTML, and will then insert the resulting HTML into the body of the message.</span></span>

- <span data-ttu-id="f6865-250">Si le gist contient autre chose que du HTML ou Markdown, le complément l’insère dans le corps du message comme un extrait de code.</span><span class="sxs-lookup"><span data-stu-id="f6865-250">If the gist contains anything other than HTML or Markdown, the add-in will insert it into the body of the message as a code snippet.</span></span>

```js
function getGist(gistId, callback) {
  var requestUrl = 'https://api.github.com/gists/' + gistId;

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gist){
    callback(gist);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (var filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      var file = gist.files[filename];
      if (!file.truncated) {
        // We have a winner.
        switch (file.language) {
          case 'HTML':
            // Insert as-is.
            callback(file.content);
            break;
          case 'Markdown':
            // Convert Markdown to HTML.
            var converter = new showdown.Converter();
            var html = converter.makeHtml(file.content);
            callback(html);
            break;
          default:
            // Insert contents as a <code> block.
            var codeBlock = '<pre><code>';
            codeBlock = codeBlock + file.content;
            codeBlock = codeBlock + '</code></pre>';
            callback(codeBlock);
        }
        return;
      }
    }
  }
  callback(null, 'No suitable file found in the gist');
}
```

### <a name="test-the-button"></a><span data-ttu-id="f6865-251">Tester le bouton</span><span class="sxs-lookup"><span data-stu-id="f6865-251">Test the button</span></span>

<span data-ttu-id="f6865-252">Enregistrez toutes vos modifications et exécutez `npm run dev-server` depuis l’invite de commandes, si le serveur n’est pas déjà en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="f6865-252">Save all of your changes and run `npm run dev-server` from the command prompt, if the server isn't already running.</span></span> <span data-ttu-id="f6865-253">Puis procédez comme suit pour tester le bouton **Insérer gist par défaut** bouton.</span><span class="sxs-lookup"><span data-stu-id="f6865-253">Then complete the following steps to test the **Insert default gist** button.</span></span>

1. <span data-ttu-id="f6865-254">Ouvrez Outlook et rédigez un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="f6865-254">Open Outlook and compose a new message.</span></span>

1. <span data-ttu-id="f6865-255">Dans la fenêtre composer un message, sélectionnez le bouton**Insérer gist par défaut**.</span><span class="sxs-lookup"><span data-stu-id="f6865-255">In the compose message window, select the **Insert default gist** button.</span></span> <span data-ttu-id="f6865-256">Vous devriez être invité à configurer le complément.</span><span class="sxs-lookup"><span data-stu-id="f6865-256">You should be prompted to configure the add-in.</span></span>

    ![Capture d’écran de l’invite à configurer le complément ](../images/addin-prompt-configure.png)

1. <span data-ttu-id="f6865-258">Dans la boîte de dialogue Paramètres, entrez votre nom d’utilisateur GitHub, puis soit **onglet** soit cliquez ailleurs dans la boîte de dialogue pour appeler l’`change` l’événement, qui devrait charger votre liste de gists.</span><span class="sxs-lookup"><span data-stu-id="f6865-258">In the settings dialog, enter your GitHub username and then either **Tab** or click elsewhere in the dialog to invoke the `change` event, which should load your list of gists.</span></span> <span data-ttu-id="f6865-259">Sélectionnez un gist par défaut, puis cliquez sur**Terminer**.</span><span class="sxs-lookup"><span data-stu-id="f6865-259">Select a gist to be the default, and select **Done**.</span></span>

    ![Capture d’écran de la boîte de dialogue des paramètres du complément](../images/addin-settings.png)

1. <span data-ttu-id="f6865-261">Cliquez de nouveau sur le bouton **Insérer un gist par défaut**.</span><span class="sxs-lookup"><span data-stu-id="f6865-261">Select the **Insert default gist** button again.</span></span> <span data-ttu-id="f6865-262">Cette fois, le contenu du gist est inséré dans le corps du courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="f6865-262">This time, you should see the contents of the gist inserted into the body of the email.</span></span>

   > [!NOTE]
   > <span data-ttu-id="f6865-263">Outlook sur Windows : pour récupérer les paramètres les plus récents, vous devrez peut-être fermer et rouvrir la fenêtre de composition d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6865-263">Outlook on Windows: To pick up the latest settings, you may need to close and reopen the compose message window.</span></span>

## <a name="implement-a-task-pane"></a><span data-ttu-id="f6865-264">Implémentation d’un volet de tâches</span><span class="sxs-lookup"><span data-stu-id="f6865-264">Implement a task pane</span></span>

<span data-ttu-id="f6865-265">Le bouton de ce complément **Insérer gist** ouvre un volet de tâches et affiche les gists de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f6865-265">This add-in's **Insert gist** button will open a task pane and display the user's gists.</span></span> <span data-ttu-id="f6865-266">L’utilisateur peut sélectionner un des gists à insérer dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f6865-266">The user can then select one of the gists to insert into the body of the message.</span></span> <span data-ttu-id="f6865-267">Si l’utilisateur n’a pas encore configuré le complément, il sera invité à le faire.</span><span class="sxs-lookup"><span data-stu-id="f6865-267">If the user has not yet configured the add-in, they will be prompted to do so.</span></span>

### <a name="specify-the-html-for-the-task-pane"></a><span data-ttu-id="f6865-268">Spécifier le code HTML pour le volet de tâches</span><span class="sxs-lookup"><span data-stu-id="f6865-268">Specify the HTML for the task pane</span></span>

<span data-ttu-id="f6865-269">Dans le projet que vous avez créé, le code HTML du volet de tâches est spécifié dans le fichier **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f6865-269">In the project that you've created, the task pane HTML is specified in the file **./src/taskpane/taskpane.html**.</span></span> <span data-ttu-id="f6865-270">Ouvrez ce fichier et remplacez l’intégralité de son contenu par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-270">Open that file and replace the entire contents with the following markup.</span></span>

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Contoso Task Pane Add-in</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l ms-landing-page">
  <main class="ms-landing-page__main">
    <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
      <div id="not-configured" style="display: none;">
        <div class="centered ms-font-xxl ms-u-textAlignCenter">Welcome!</div>
        <div class="ms-font-xl" id="settings-prompt">Please choose the <strong>Settings</strong> icon at the bottom of this window to configure this add-in.</div>
      </div>
      <div id="gist-list-container" style="display: none;">
        <form>
          <div id="gist-list">
          </div>
        </form>
      </div>
      <div id="error-display" style="display: none;" class="ms-u-borderBase ms-fontColor-error ms-font-m ms-bgColor-error ms-borderColor-error">
      </div>
    </section>
    <button class="ms-Button ms-Button--primary" id="insert-button" tabindex=0 disabled>
      <span class="ms-Button-label">Insert</span>
    </button>
  </main>
  <footer class="ms-landing-page__footer ms-bgColor-themePrimary">
    <div class="ms-landing-page__footer--left">
      <img src="../../assets/logo-filled.png" />
      <h1 class="ms-font-xl ms-fontWeight-semilight ms-fontColor-white">Git the gist</h1>
    </div>
    <div id="settings-icon" class="ms-landing-page__footer--right" aria-label="Settings" tabindex=0>
      <i class="ms-Icon enlarge ms-Icon--Settings ms-fontColor-white"></i>
    </div>
  </footer>
  <script type="text/javascript" src="../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../node_modules/showdown/dist/showdown.min.js"></script>
  <script type="text/javascript" src="../node_modules/urijs/src/URI.min.js"></script>
  <script type="text/javascript" src="../src/helpers/addin-config.js"></script>
  <script type="text/javascript" src="../src/helpers/gist-api.js"></script>
  <script type="text/javascript" src="taskpane.js"></script>
</body>

</html>
```

### <a name="specify-the-css-for-the-task-pane"></a><span data-ttu-id="f6865-271">Spécifier le style CSS pour le volet de tâches</span><span class="sxs-lookup"><span data-stu-id="f6865-271">Specify the CSS for the task pane</span></span>

<span data-ttu-id="f6865-272">Dans le projet que vous avez créé, le style CSS du volet de tâches est spécifié dans le fichier **./src/taskpane/taskpane.css**.</span><span class="sxs-lookup"><span data-stu-id="f6865-272">In the project that you've created, the task pane CSS is specified in the file **./src/taskpane/taskpane.css**.</span></span> <span data-ttu-id="f6865-273">Ouvrez ce fichier et remplacez l’intégralité de son contenu par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-273">Open that file and replace the entire contents with the following code.</span></span>

```css
/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */
html, body {
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0;
  overflow: auto; }

body {
  position: relative;
  font-size: 16px; }

main {
  height: 100%;
  overflow-y: auto; }

footer {
  width: 100%;
  position: relative;
  bottom: 0;
  margin-top: 10px;}

p, h1, h2, h3, h4, h5, h6 {
  margin: 0;
  padding: 0; }

ul {
  padding: 0; }

#settings-prompt {
  margin: 10px 0;
}

#error-display {
  padding: 10px;
}

#insert-button {
  margin: 0 10px;
}

.clearfix {
  display: block;
  clear: both;
  height: 0; }

.pointerCursor {
  cursor: pointer; }

.invisible {
  visibility: hidden; }

.undisplayed {
  display: none; }

.ms-Icon.enlarge {
  position: relative;
  font-size: 20px;
  top: 4px; }

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}

.ms-landing-page {
  display: -webkit-flex;
  display: flex;
  -webkit-flex-direction: column;
          flex-direction: column;
  -webkit-flex-wrap: nowrap;
          flex-wrap: nowrap;
  height: 100%; }
  .ms-landing-page__main {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    height: 100%; }

  .ms-landing-page__content {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    height: 100%;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    padding: 20px; }
    .ms-landing-page__content h2 {
      margin-bottom: 20px; }
  .ms-landing-page__footer {
    display: -webkit-inline-flex;
    display: inline-flex;
    -webkit-justify-content: center;
            justify-content: center;
    -webkit-align-items: center;
            align-items: center; }
    .ms-landing-page__footer--left {
      transition: background ease 0.1s, color ease 0.1s;
      display: -webkit-inline-flex;
      display: inline-flex;
      -webkit-justify-content: flex-start;
              justify-content: flex-start;
      -webkit-align-items: center;
              align-items: center;
      -webkit-flex: 1 0 0px;
              flex: 1 0 0px;
      padding: 20px; }
      .ms-landing-page__footer--left:active, .ms-landing-page__footer--left:hover {
        background: #005ca4;
        cursor: pointer; }
      .ms-landing-page__footer--left:active {
        background: #005ca4; }
      .ms-landing-page__footer--left--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--left--disabled:active, .ms-landing-page__footer--left--disabled:hover {
          background: transparent; }
      .ms-landing-page__footer--left img {
        width: 40px;
        height: 40px; }
      .ms-landing-page__footer--left h1 {
        -webkit-flex: 1 0 0px;
                flex: 1 0 0px;
        margin-left: 15px;
        text-align: left;
        width: auto;
        max-width: auto;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis; }
    .ms-landing-page__footer--right {
      transition: background ease 0.1s, color ease 0.1s;
      padding: 29px 20px; }
      .ms-landing-page__footer--right:active, .ms-landing-page__footer--right:hover {
        background: #005ca4;
        cursor: pointer; }
      .ms-landing-page__footer--right:active {
        background: #005ca4; }
      .ms-landing-page__footer--right--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--right--disabled:active, .ms-landing-page__footer--right--disabled:hover {
          background: transparent; }
```

### <a name="specify-the-javascript-for-the-task-pane"></a><span data-ttu-id="f6865-274">Spécifier le code JavaScript pour le volet de tâches</span><span class="sxs-lookup"><span data-stu-id="f6865-274">Specify the JavaScript for the task pane</span></span>

<span data-ttu-id="f6865-275">Dans le projet que vous avez créé, le code JavaScript du volet de tâches est spécifié dans le fichier **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f6865-275">In the project that you've created, the task pane JavaScript is specified in the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="f6865-276">Ouvrez ce fichier et remplacez l’intégralité de son contenu par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="f6865-276">Open that file and replace the entire contents with the following code.</span></span>

```js
(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        $('#not-configured').show();
      }

      // When insert button is selected, build the content
      // and insert into the body.
      $('#insert-button').on('click', function(){
        var gistId = $('.ms-ListItem.is-selected').val();
        getGist(gistId, function(gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                  {coercionType: Office.CoercionType.Html}, function(result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError('Could not insert gist: ' + result.error.message);
                    }
                });
              } else {
                showError('Could not create insertable content: ' + error);
              }
            });
          } else {
            showError('Could not retrieve gist: ' + error);
          }
        });
      });

      // When the settings icon is selected, open the settings dialog.
      $('#settings-icon').on('click', function(){
        // Display settings dialog.
        var url = new URI('../src/settings/dialog.html').absoluteTo(window.location).toString();
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
        }

        var dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
        });
      })
    });
  };

  function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function(gists, error) {
      if (error) {

      } else {
        $('#gist-list').empty();
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
    $('#insert-button').removeAttr('disabled');
  }

  function showError(error) {
    $('#not-configured').hide();
    $('#gist-list-container').hide();
    $('#error-display').text(error);
    $('#error-display').show();
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
```

### <a name="test-the-button"></a><span data-ttu-id="f6865-277">Tester le bouton</span><span class="sxs-lookup"><span data-stu-id="f6865-277">Test the button</span></span>

<span data-ttu-id="f6865-278">Enregistrez toutes vos modifications et exécutez `npm run dev-server` depuis l’invite de commandes, si le serveur n’est pas déjà en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="f6865-278">Save all of your changes and run `npm run dev-server` from the command prompt, if the server isn't already running.</span></span> <span data-ttu-id="f6865-279">Puis procédez comme suit pour tester le bouton **Insérer gist**.</span><span class="sxs-lookup"><span data-stu-id="f6865-279">Then complete the following steps to test the **Insert gist** button.</span></span>

1. <span data-ttu-id="f6865-280">Ouvrez Outlook et rédigez un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="f6865-280">Open Outlook and compose a new message.</span></span>

1. <span data-ttu-id="f6865-281">Dans la fenêtre composer un message, sélectionnez le bouton**Insérer gist**.</span><span class="sxs-lookup"><span data-stu-id="f6865-281">In the compose message window, select the **Insert gist** button.</span></span> <span data-ttu-id="f6865-282">Vous devriez voir un volet des tâches qui s’ouvre à droite du formulaire Composer.</span><span class="sxs-lookup"><span data-stu-id="f6865-282">You should see a task pane open to the right of the compose form.</span></span>

1. <span data-ttu-id="f6865-283">Dans le volet des tâches, sélectionnez le gist**Hello World Html**, puis sélectionnez **insérer** pour insérer ce gist dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f6865-283">In the task pane, select the **Hello World Html** gist and select **Insert** to insert that gist into the body of the message.</span></span>

![Capture d’écran du volet de tâhces du complément](../images/addin-taskpane.png)

## <a name="next-steps"></a><span data-ttu-id="f6865-285">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f6865-285">Next steps</span></span>

<span data-ttu-id="f6865-286">Ce didacticiel vous a appris à créer un complément Outlook qui peut être utilisé pour dans le mode composer un message pour insérer du contenu dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f6865-286">In this tutorial, you've created an Outlook add-in that can be used in message compose mode to insert content into the body of a message.</span></span> <span data-ttu-id="f6865-287">Pour en savoir plus sur le développement des complément Outlook, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="f6865-287">To learn more about developing Outlook add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f6865-288">API de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="f6865-288">Outlook add-in APIs</span></span>](../outlook/apis.md)
