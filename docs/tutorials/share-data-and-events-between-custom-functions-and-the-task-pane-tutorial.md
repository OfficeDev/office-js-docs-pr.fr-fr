---
title: 'Didacticiel : partager des données et des événements entre des fonctions personnalisées Excel et le volet Office'
description: Dans Excel, partagez des données et des événements entre des fonctions personnalisées et le volet Office.
ms.date: 05/17/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6d96b10d6cd6e9bb7909b9d6d64b9a65fcac5b3a
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275600"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a><span data-ttu-id="425a1-103">Didacticiel : partager des données et des événements entre des fonctions personnalisées Excel et le volet Office</span><span class="sxs-lookup"><span data-stu-id="425a1-103">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>

<span data-ttu-id="425a1-104">Vous pouvez configurer votre complément Excel pour utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="425a1-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="425a1-105">Cela permet de partager des données globales ou d’envoyer des événements entre le volet Office et les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="425a1-105">This makes it possible to shared global data, or send events between the task pane and custom functions.</span></span>

<span data-ttu-id="425a1-106">Pour la plupart des scénarios de fonctions personnalisées, nous vous recommandons d’utiliser un runtime partagé, sauf si vous avez une raison particulière d’utiliser une fonction personnalisée de volet non-tâche (sans interface utilisateur).</span><span class="sxs-lookup"><span data-stu-id="425a1-106">For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function.</span></span>

<span data-ttu-id="425a1-107">Ce didacticiel part du principe que vous êtes familiarisé avec l’utilisation du générateur Yo Office pour créer des projets de complément.</span><span class="sxs-lookup"><span data-stu-id="425a1-107">This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects.</span></span> <span data-ttu-id="425a1-108">Envisagez de compléter le [Didacticiel des fonctions personnalisées Excel](./excel-tutorial-create-custom-functions.md), si vous ne l’avez pas encore fait.</span><span class="sxs-lookup"><span data-stu-id="425a1-108">Consider completing the [Excel custom functions tutorial](./excel-tutorial-create-custom-functions.md), if you haven't already.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="425a1-109">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="425a1-109">Create the add-in project</span></span>

<span data-ttu-id="425a1-110">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="425a1-110">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="425a1-111">Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :</span><span class="sxs-lookup"><span data-stu-id="425a1-111">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="425a1-112">Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**</span><span class="sxs-lookup"><span data-stu-id="425a1-112">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="425a1-113">Choose a script type (Choisissez un type de script) :  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="425a1-113">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="425a1-114">What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**</span><span class="sxs-lookup"><span data-stu-id="425a1-114">What do you want to name your add-in? **My Office Add-in**</span></span>

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

<span data-ttu-id="425a1-116">Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="425a1-116">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="425a1-117">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="425a1-117">Configure the manifest</span></span>

1. <span data-ttu-id="425a1-118">Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="425a1-118">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="425a1-119">Ouvrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="425a1-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="425a1-120">Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="425a1-120">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="425a1-121">La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé.</span><span class="sxs-lookup"><span data-stu-id="425a1-121">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="425a1-122">Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="425a1-122">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="425a1-123">Dans la section `<DesktopFormFactor>`, changez la valeur **Commands.Url** de **FunctionFile** pour utiliser **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="425a1-123">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="425a1-124">Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="425a1-124">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="425a1-125">Ajoutez un nouvel **ID d’URL** pour **ContosoAddin.Url** qui pointe vers **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="425a1-125">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="425a1-126">Enregistrez vos changements et regénérez le projet.</span><span class="sxs-lookup"><span data-stu-id="425a1-126">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="425a1-127">Partager l’état entre une fonction personnalisée et du code du volet Office</span><span class="sxs-lookup"><span data-stu-id="425a1-127">Share state between custom function and task pane code</span></span>

<span data-ttu-id="425a1-128">À présent que les fonctions personnalisées s’exécutent dans le même contexte que votre code du volet Office, elles peuvent partager l’état directement, sans utiliser l’objet **Storage**.</span><span class="sxs-lookup"><span data-stu-id="425a1-128">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="425a1-129">Les instructions suivantes montrent comment partager une variable globale entre une fonction personnalisée et du code du volet Office.</span><span class="sxs-lookup"><span data-stu-id="425a1-129">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="425a1-130">Créer des fonctions personnalisées pour obtenir ou stocker l’état partagé</span><span class="sxs-lookup"><span data-stu-id="425a1-130">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="425a1-131">Dans Visual Studio Code, ouvrez le fichier **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="425a1-131">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="425a1-132">Sur la ligne 1, tout en haut, insérez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="425a1-132">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="425a1-133">Cette opération initialise une variable globale nommée **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="425a1-133">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="425a1-134">Ajoutez le code suivant pour créer une fonction personnalisée qui stocke des valeurs dans la variable **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="425a1-134">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="425a1-135">Ajoutez le code suivant pour créer une fonction personnalisée qui obtient la valeur actuelle de la variable **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="425a1-135">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="425a1-136">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="425a1-136">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="425a1-137">Créer des contrôles du volet Office pour utiliser des données globales</span><span class="sxs-lookup"><span data-stu-id="425a1-137">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="425a1-138">Ouvrez le fichier **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="425a1-138">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="425a1-139">Ajoutez l'élément de script suivant juste avant l’élément `</head>`.</span><span class="sxs-lookup"><span data-stu-id="425a1-139">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="425a1-140">Après l’élément de fermeture `</main>`, ajoutez le code HTML suivant.</span><span class="sxs-lookup"><span data-stu-id="425a1-140">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="425a1-141">Le code HTML crée deux zones de texte et des boutons permettant d’obtenir ou de stocker des données globales.</span><span class="sxs-lookup"><span data-stu-id="425a1-141">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

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

4. <span data-ttu-id="425a1-142">Avant l’élément `<body>`, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="425a1-142">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="425a1-143">Ce code gère les événements de clic de bouton quand l’utilisateur souhaite stocker ou obtenir des données globales.</span><span class="sxs-lookup"><span data-stu-id="425a1-143">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="425a1-144">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="425a1-144">Save the file.</span></span>
6. <span data-ttu-id="425a1-145">Générez le projet.</span><span class="sxs-lookup"><span data-stu-id="425a1-145">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="425a1-146">Essayer de partager des données entre les fonctions personnalisées et le volet Office</span><span class="sxs-lookup"><span data-stu-id="425a1-146">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="425a1-147">Démarrez le projet à l’aide de la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="425a1-147">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="425a1-148">Une fois Excel démarré, vous pouvez utiliser les boutons du volet Office pour stocker ou obtenir des données partagées.</span><span class="sxs-lookup"><span data-stu-id="425a1-148">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="425a1-149">Entrez `=CONTOSO.GETVALUE()` dans une cellule pour que la fonction personnalisée extraie les mêmes données partagées.</span><span class="sxs-lookup"><span data-stu-id="425a1-149">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="425a1-150">Vous pouvez également utiliser `=CONTOSO.STOREVALUE("new value")` pour remplacer les données partagées par une nouvelle valeur.</span><span class="sxs-lookup"><span data-stu-id="425a1-150">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="425a1-151">La configuration de votre projet comme illustré dans cet article permet de partager le contexte entre des fonctions personnalisées et le volet Office.</span><span class="sxs-lookup"><span data-stu-id="425a1-151">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="425a1-152">Il est possible d’appeler des API Office à partir de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="425a1-152">Calling some Office APIs from custom functions is possible.</span></span> <span data-ttu-id="425a1-153">Pour plus d’informations, [reportez-vous à la rubrique Call Microsoft Excel API from a Custom Function](../excel/call-excel-apis-from-custom-function.md) .</span><span class="sxs-lookup"><span data-stu-id="425a1-153">[See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.</span></span>
