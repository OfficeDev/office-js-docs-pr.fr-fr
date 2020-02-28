---
ms.date: 02/20/2020
title: 'Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office (préversion)'
ms.prod: excel
description: Dans Excel, partagez des données et des événements entre des fonctions personnalisées et le volet Office.
localization_priority: Priority
ms.openlocfilehash: 13ef4c199f7cb1de84e58f0ada554c851aee0cad
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283890"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="c6ec1-103">Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="c6ec1-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="c6ec1-104">Vous pouvez configurer votre complément Excel pour utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="c6ec1-105">Vous pouvez ainsi partager des données globales ou envoyer des événements entre le volet des tâches et les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-105">This will make it possible to shared global data, or send events between the task pane and custom functions.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="c6ec1-106">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="c6ec1-106">Create the add-in project</span></span>

<span data-ttu-id="c6ec1-107">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-107">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="c6ec1-108">Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :</span><span class="sxs-lookup"><span data-stu-id="c6ec1-108">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="c6ec1-109">Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**</span><span class="sxs-lookup"><span data-stu-id="c6ec1-109">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="c6ec1-110">Choose a script type (Choisissez un type de script) :  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="c6ec1-110">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="c6ec1-111">What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**</span><span class="sxs-lookup"><span data-stu-id="c6ec1-111">What do you want to name your add-in? **My Office Add-in**</span></span>

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

<span data-ttu-id="c6ec1-113">Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-113">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="c6ec1-114">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="c6ec1-114">Configure the manifest</span></span>

1. <span data-ttu-id="c6ec1-115">Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-115">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="c6ec1-116">Ouvrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-116">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="c6ec1-117">Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-117">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="c6ec1-118">La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-118">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="c6ec1-119">Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-119">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="c6ec1-120">Dans la section `<DesktopFormFactor>`, changez la valeur **Commands.Url** de **FunctionFile** pour utiliser **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-120">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="c6ec1-121">Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-121">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="c6ec1-122">Ajoutez un nouvel **ID d’URL** pour **ContosoAddin.Url** qui pointe vers **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-122">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="c6ec1-123">Enregistrez vos changements et regénérez le projet.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-123">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="c6ec1-124">Partager l’état entre une fonction personnalisée et du code du volet Office</span><span class="sxs-lookup"><span data-stu-id="c6ec1-124">Share state between custom function and task pane code</span></span>

<span data-ttu-id="c6ec1-125">À présent que les fonctions personnalisées s’exécutent dans le même contexte que votre code du volet Office, elles peuvent partager l’état directement, sans utiliser l’objet **Storage**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-125">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="c6ec1-126">Les instructions suivantes montrent comment partager une variable globale entre une fonction personnalisée et du code du volet Office.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-126">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="c6ec1-127">Créer des fonctions personnalisées pour obtenir ou stocker l’état partagé</span><span class="sxs-lookup"><span data-stu-id="c6ec1-127">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="c6ec1-128">Dans Visual Studio Code, ouvrez le fichier **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-128">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="c6ec1-129">Sur la ligne 1, tout en haut, insérez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-129">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="c6ec1-130">Cette opération initialise une variable globale nommée **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-130">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="c6ec1-131">Ajoutez le code suivant pour créer une fonction personnalisée qui stocke des valeurs dans la variable **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-131">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="c6ec1-132">Ajoutez le code suivant pour créer une fonction personnalisée qui obtient la valeur actuelle de la variable **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-132">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="c6ec1-133">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-133">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="c6ec1-134">Créer des contrôles du volet Office pour utiliser des données globales</span><span class="sxs-lookup"><span data-stu-id="c6ec1-134">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="c6ec1-135">Ouvrez le fichier **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-135">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="c6ec1-136">Ajoutez l'élément de script suivant juste avant l’élément `</head>`.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-136">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="c6ec1-137">Après l’élément de fermeture `</main>`, ajoutez le code HTML suivant.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-137">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="c6ec1-138">Le code HTML crée deux zones de texte et des boutons permettant d’obtenir ou de stocker des données globales.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-138">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
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

4. <span data-ttu-id="c6ec1-139">Avant l’élément `<body>`, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-139">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="c6ec1-140">Ce code gère les événements de clic de bouton quand l’utilisateur souhaite stocker ou obtenir des données globales.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-140">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="c6ec1-141">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-141">Save the file.</span></span>
6. <span data-ttu-id="c6ec1-142">Générez le projet.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-142">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="c6ec1-143">Essayer de partager des données entre les fonctions personnalisées et le volet Office</span><span class="sxs-lookup"><span data-stu-id="c6ec1-143">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="c6ec1-144">Démarrez le projet à l’aide de la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-144">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="c6ec1-145">Une fois Excel démarré, vous pouvez utiliser les boutons du volet Office pour stocker ou obtenir des données partagées.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-145">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="c6ec1-146">Entrez `=CONTOSO.GETVALUE()` dans une cellule pour que la fonction personnalisée extraie les mêmes données partagées.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-146">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="c6ec1-147">Vous pouvez également utiliser `=CONTOSO.STOREVALUE(“new value”)` pour remplacer les données partagées par une nouvelle valeur.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-147">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="c6ec1-148">La configuration de votre projet comme illustré dans cet article permet de partager le contexte entre des fonctions personnalisées et le volet Office.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-148">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="c6ec1-149">L’appel d’API Office à partir de fonctions personnalisées n’est pas pris en charge dans la préversion.</span><span class="sxs-lookup"><span data-stu-id="c6ec1-149">Calling Office APIs from custom functions is not supported in the preview.</span></span>
