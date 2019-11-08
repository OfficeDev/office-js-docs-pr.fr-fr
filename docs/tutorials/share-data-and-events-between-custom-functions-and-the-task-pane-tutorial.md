---
ms.date: 11/04/2019
title: 'Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office (préversion)'
ms.prod: excel
description: Dans Excel, partagez des données et des événements entre des fonctions personnalisées et le volet Office.
localization_priority: Priority
ms.openlocfilehash: dcd4bced7e1419a57256f4ec54e3ff72c0edf9ef
ms.sourcegitcommit: 42bcf9059327a8d71a7ab223805aea68be9ed6b5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/04/2019
ms.locfileid: "37962102"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="30d8d-103">Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="30d8d-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

<span data-ttu-id="30d8d-104">Les fonctions personnalisées Excel et le volet Office partagent des données globales et peuvent effectuer des appels de fonction entre elles.</span><span class="sxs-lookup"><span data-stu-id="30d8d-104">Excel custom functions and the task pane share global data, and can make function calls into each other.</span></span> <span data-ttu-id="30d8d-105">Pour configurer votre projet de sorte que les fonctions personnalisées puissent fonctionner avec le volet Office, suivez les instructions décrites dans cet article.</span><span class="sxs-lookup"><span data-stu-id="30d8d-105">To configure your project so that custom functions can work with the task pane, follow the instructions in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="30d8d-106">Les fonctionnalités décrites dans cet article sont actuellement en préversion et peuvent faire l’objet de modifications.</span><span class="sxs-lookup"><span data-stu-id="30d8d-106">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="30d8d-107">Elles ne sont pas prises en charge dans les environnements de production pour l’instant.</span><span class="sxs-lookup"><span data-stu-id="30d8d-107">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="30d8d-108">Les fonctionnalités en préversion de cet article sont uniquement disponibles dans Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="30d8d-108">The preview features in this article are only available on Excel on Windows.</span></span> <span data-ttu-id="30d8d-109">Pour essayer les fonctionnalités en préversion, vous devez [rejoindre Office Insider](https://insider.office.com/fr-FR/join).</span><span class="sxs-lookup"><span data-stu-id="30d8d-109">To try the preview features, you will need to [join Office Insider](https://insider.office.com/fr-FR/join).</span></span>  <span data-ttu-id="30d8d-110">Un bon moyen de tester les fonctionnalités en préversion consiste à utiliser un abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="30d8d-110">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="30d8d-111">Si vous n’avez pas d’abonnement Office 365, vous pouvez en obtenir un en rejoignant le [programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="30d8d-111">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="30d8d-112">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="30d8d-112">Create the add-in project</span></span>

<span data-ttu-id="30d8d-113">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="30d8d-113">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="30d8d-114">Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :</span><span class="sxs-lookup"><span data-stu-id="30d8d-114">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="30d8d-115">Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**</span><span class="sxs-lookup"><span data-stu-id="30d8d-115">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="30d8d-116">Choose a script type (Choisissez un type de script) :  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="30d8d-116">Choose a script type: </span></span>
- <span data-ttu-id="30d8d-117">What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**</span><span class="sxs-lookup"><span data-stu-id="30d8d-117">What do you want to name your add-in? **My Office Add-in**</span></span>

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

<span data-ttu-id="30d8d-119">Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="30d8d-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="30d8d-120">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="30d8d-120">Configure the add-in manifest</span></span>

1. <span data-ttu-id="30d8d-121">Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-121">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="30d8d-122">Ouvrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-122">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="30d8d-123">Changez la section `<Requirements>` afin d’utiliser **CustomFunctionsRuntime** version **12**, comme illustré dans le code suivant.</span><span class="sxs-lookup"><span data-stu-id="30d8d-123">Change the `<Requirements>` section to use **CustomFunctionsRuntime** version **1.2** as shown in the following code.</span></span>
    
    ```xml
    <Requirements> 
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. <span data-ttu-id="30d8d-124">Sous l’élément `<Host>` du classeur, ajoutez la section `<Runtimes>` suivante.</span><span class="sxs-lookup"><span data-stu-id="30d8d-124">Under the `<Host>` element for the workbook, add the following `<Runtimes>` section.</span></span> <span data-ttu-id="30d8d-125">La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé.</span><span class="sxs-lookup"><span data-stu-id="30d8d-125">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>
    
    ```xml
    <Hosts>
    <Host xsi:type="Workbook">
    <Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
    </Runtimes>
    ```
    
5. <span data-ttu-id="30d8d-126">Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. <span data-ttu-id="30d8d-127">Dans la section `<DesktopFormFactor>`, changez la valeur **Commands.Url** de **FunctionFile** pour utiliser **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-127">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. <span data-ttu-id="30d8d-128">Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-128">In the `<Action>` section, change the source location from **Taskpane.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. <span data-ttu-id="30d8d-129">Ajoutez un nouvel **ID d’URL** pour **TaskPaneAndCustomFunction.Url** qui pointe vers **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-129">Add a new **Url id** for **TaskPaneAndCustomFunction.Url** that points to **taskpane.html**.</span></span>
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. <span data-ttu-id="30d8d-130">Enregistrez vos changements et regénérez le projet.</span><span class="sxs-lookup"><span data-stu-id="30d8d-130">Save your changes and rebuild the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="30d8d-131">Partager l’état entre une fonction personnalisée et du code du volet Office</span><span class="sxs-lookup"><span data-stu-id="30d8d-131">Share state between custom function and task pane code</span></span> 

<span data-ttu-id="30d8d-132">À présent que les fonctions personnalisées s’exécutent dans le même contexte que votre code du volet Office, elles peuvent partager l’état directement, sans utiliser l’objet **Storage**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-132">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="30d8d-133">Les instructions suivantes montrent comment partager une variable globale entre une fonction personnalisée et du code du volet Office.</span><span class="sxs-lookup"><span data-stu-id="30d8d-133">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="30d8d-134">Créer des fonctions personnalisées pour obtenir ou stocker l’état partagé</span><span class="sxs-lookup"><span data-stu-id="30d8d-134">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="30d8d-135">Dans Visual Studio Code, ouvrez le fichier **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-135">In Visual Studio Code, open the file **config\config.json**.</span></span>
2. <span data-ttu-id="30d8d-136">Sur la ligne 1, tout en haut, insérez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="30d8d-136">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="30d8d-137">Cette opération initialise une variable globale nommée **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-137">This will initialize a global variable named **sharedState**.</span></span>
    
    ```js
    window.sharedState = "empty";
    ```
    
3. <span data-ttu-id="30d8d-138">Ajoutez le code suivant pour créer une fonction personnalisée qui stocke des valeurs dans la variable **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-138">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>
    
    ```js
    /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
    function storeValue(sharedValue) {
    window.sharedState = sharedValue;
    return "value stored";
    }
    ```
    
4. <span data-ttu-id="30d8d-139">Ajoutez le code suivant pour créer une fonction personnalisée qui obtient la valeur actuelle de la variable **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-139">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

    ```js
    /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
    function getValue() {
    return window.sharedState;
    }
    ```
    
5. <span data-ttu-id="30d8d-140">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="30d8d-140">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="30d8d-141">Créer des contrôles du volet Office pour utiliser des données globales</span><span class="sxs-lookup"><span data-stu-id="30d8d-141">Create task pane controls to work with global data</span></span> 

1. <span data-ttu-id="30d8d-142">Ouvrez le fichier **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="30d8d-142">Open the file**src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="30d8d-143">Après l’élément de fermeture `</main>`, ajoutez le code HTML suivant.</span><span class="sxs-lookup"><span data-stu-id="30d8d-143">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="30d8d-144">Le code HTML crée deux zones de texte et des boutons permettant d’obtenir ou de stocker des données globales.</span><span class="sxs-lookup"><span data-stu-id="30d8d-144">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
    <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>
    <p>Store new value to shared state</p>
    <div>
    <input type="text" id="storeBox" />
    <button onclick="storeSharedValue()">Store</button>
    </div>
     
    <p>Get shared state value</p>
    <div>
    <input type="text" id="getBox" />
    <button onclick="getSharedValue()">Get</button>
    </div>
    ```
    
3. <span data-ttu-id="30d8d-145">Avant l’élément `<body>`, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="30d8d-145">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="30d8d-146">Ce code gère les événements de clic de bouton quand l’utilisateur souhaite stocker ou obtenir des données globales.</span><span class="sxs-lookup"><span data-stu-id="30d8d-146">This code will handle the button click events when the user wants to store or get global data.</span></span>
    
    ```js
    <script>
    function storeSharedValue() {
    let sharedValue = document.getElementById('storeBox').value;
    window.sharedState = sharedValue;
    }
    
    function getSharedValue() {
    document.getElementById('getBox').value = window.sharedState;
    }</script>
    ```
    
4. <span data-ttu-id="30d8d-147">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="30d8d-147">Save the file.</span></span>
5. <span data-ttu-id="30d8d-148">Générez le projet.</span><span class="sxs-lookup"><span data-stu-id="30d8d-148">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="30d8d-149">Essayer de partager des données entre les fonctions personnalisées et le volet Office</span><span class="sxs-lookup"><span data-stu-id="30d8d-149">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="30d8d-150">Démarrez le projet à l’aide de la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="30d8d-150">Start the migration by using the following command.</span></span>

    ```command&nbsp;line
    npm run start
    ```

<span data-ttu-id="30d8d-151">Une fois Excel démarré, vous pouvez utiliser les boutons du volet Office pour stocker ou obtenir des données partagées.</span><span class="sxs-lookup"><span data-stu-id="30d8d-151">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="30d8d-152">Entrez `=CONTOSO.GETVALUE()` dans une cellule pour que la fonction personnalisée extraie les mêmes données partagées.</span><span class="sxs-lookup"><span data-stu-id="30d8d-152">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="30d8d-153">Vous pouvez également utiliser `=CONTOSO.STOREVALUE(“new value”)` pour remplacer les données partagées par une nouvelle valeur.</span><span class="sxs-lookup"><span data-stu-id="30d8d-153">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="30d8d-154">La configuration de votre projet comme illustré dans cet article permet de partager le contexte entre des fonctions personnalisées et le volet Office.</span><span class="sxs-lookup"><span data-stu-id="30d8d-154">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="30d8d-155">L’appel d’API Office à partir de fonctions personnalisées n’est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="30d8d-155">Calling Office APIs from custom functions is not supported.</span></span> <span data-ttu-id="30d8d-156">Si vous avez besoin d’interagir avec le document, implémentez des appels aux API Office dans l’[événement onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details).</span><span class="sxs-lookup"><span data-stu-id="30d8d-156">If you need to interact with the document, implement calls to the Office APIs in the [onCalculated event](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details).</span></span>

