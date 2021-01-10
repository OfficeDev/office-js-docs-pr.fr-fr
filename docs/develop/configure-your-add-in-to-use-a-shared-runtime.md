---
ms.date: 12/28/2020
title: Configurez votre complément Office pour utiliser un runtime JavaScript partagé
ms.prod: non-product-specific
description: Configurez votre complément Office afin d’utiliser un runtime JavaScript partagé pour prendre en charge un ruban supplémentaire, un volet des tâches et des fonctionnalités personnalisées.
localization_priority: Priority
ms.openlocfilehash: e1248ce28a45ad63ac9b02093a39810ee042bb80
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789235"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="33131-103">Configurez votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="33131-103">Configure your Office Add-in to use a shared JavaScript runtime</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="33131-104">Vous pouvez configurer votre complément Office pour exécuter la totalité de son code dans un seul runtime JavaScript partagé (également connu sous le nom de runtime partagé).</span><span class="sxs-lookup"><span data-stu-id="33131-104">You can configure your Office Add-in to run all of its code in a single shared JavaScript runtime (also known as a shared runtime).</span></span> <span data-ttu-id="33131-105">Vous pouvez ainsi améliorer la coordination dans votre complément et accéder aux DOM et CORS à partir de toutes les parties de votre complément.</span><span class="sxs-lookup"><span data-stu-id="33131-105">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="33131-106">Il active également des fonctionnalités supplémentaires telles que l’exécution d’un code lors de l’ouverture d’un document, ou l’activation et la désactivation des boutons du ruban.</span><span class="sxs-lookup"><span data-stu-id="33131-106">It also enables additional features such as running code when the document opens, or enabling or disabling ribbon buttons.</span></span> <span data-ttu-id="33131-107">Si vous voulez configurer votre complément pour utiliser un runtime partagé JavaScript, suivez les instructions contenues dans cet article.</span><span class="sxs-lookup"><span data-stu-id="33131-107">To configure your add-in to use a shared JavaScript runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="33131-108">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="33131-108">Create the add-in project</span></span>

<span data-ttu-id="33131-109">Si vous démarrez un nouveau projet, suivez ces étapes pour utiliser le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office) et créer le projet de complément Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="33131-109">If you are starting a new project, follow these steps to use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create an Excel or PowerPoint add-in project.</span></span>

<span data-ttu-id="33131-110">Effectuez l'une des opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="33131-110">Do one of the following:</span></span>

- <span data-ttu-id="33131-111">Pour créer un complément Excel avec fonctions personnalisées, exécutez la commande `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js`.</span><span class="sxs-lookup"><span data-stu-id="33131-111">To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js`.</span></span>
    
    <span data-ttu-id="33131-112">ou</span><span class="sxs-lookup"><span data-stu-id="33131-112">or</span></span>
    
- <span data-ttu-id="33131-113">Pour créer un complément PowerPoint, exécutez la commande `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js`.</span><span class="sxs-lookup"><span data-stu-id="33131-113">To generate a PowerPoint add-in, run the command `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js`.</span></span>

<span data-ttu-id="33131-114">Le générateur crée le projet et installe les composants de nœud de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="33131-114">The generator will create the project and install supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="33131-115">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="33131-115">Configure the manifest</span></span>

<span data-ttu-id="33131-116">Procédez comme suit pour configurer un projet nouveau ou existant de manière à utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="33131-116">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span> <span data-ttu-id="33131-117">Ces étapes supposent que vous avez créé votre projet à l’aide du [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="33131-117">These steps assume you have generated your project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="33131-118">Démarrez Visual Studio Code, puis ouvrez le projet de complément Excel ou PowerPoint créé.</span><span class="sxs-lookup"><span data-stu-id="33131-118">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
2. <span data-ttu-id="33131-119">Ouvrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="33131-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="33131-120">Si vous avez créé un complément Excel, mettez à jour la section des conditions préalables pour utiliser un runtime partagé au lieu du runtime de fonction partagé.</span><span class="sxs-lookup"><span data-stu-id="33131-120">If you generated an Excel add-in, update the requirements section to use the shared runtime instead of the custom function runtime.</span></span> <span data-ttu-id="33131-121">Le XML s’affiche comme suit.</span><span class="sxs-lookup"><span data-stu-id="33131-121">The XML should appear as follows.</span></span>
    
    ```xml
    <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
    </Requirements>
    ```
        
4. <span data-ttu-id="33131-122">Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`, juste dans la balise `<Host ...>`.</span><span class="sxs-lookup"><span data-stu-id="33131-122">Find the `<VersionOverrides>` section and add the following `<Runtimes>` section just inside the `<Host ...>` tag.</span></span> <span data-ttu-id="33131-123">La durée de vie doit être **longue** afin que votre code de complément puisse s’exécuter même quand le volet Office est fermé.</span><span class="sxs-lookup"><span data-stu-id="33131-123">The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed.</span></span> <span data-ttu-id="33131-124">La valeur `resid` est **Taskpane.Url** qui se réfère à l’emplacement du fichier **taskpane.html** spécifiée dans la section ` <bt:Urls>` près du bas du fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="33131-124">The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.</span></span>

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       ...
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

5. <span data-ttu-id="33131-125">Si vous avez créé un complément Excel avec des fonctions personnalisées, recherchez l’élément `<Page>`.</span><span class="sxs-lookup"><span data-stu-id="33131-125">If you generated an Excel add-in with custom functions, find the `<Page>` element.</span></span> <span data-ttu-id="33131-126">Puis remplacez l’emplacement de la source **Functions.Page.Url** par **TaskPane.Url**.</span><span class="sxs-lookup"><span data-stu-id="33131-126">Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

6. <span data-ttu-id="33131-127">Recherchez la balise `<FunctionFile ...>` et remplacez le `resid` de **Commands.Url** par **Taskpane.Url**.</span><span class="sxs-lookup"><span data-stu-id="33131-127">Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**.</span></span> <span data-ttu-id="33131-128">Veuillez noter que si vous n'avez pas de commandes d'action, vous ne disposerez pas d'entrée **FunctionFile**. Vous pouvez par conséquent ignorer cette étape.</span><span class="sxs-lookup"><span data-stu-id="33131-128">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

7. <span data-ttu-id="33131-129">Enregistrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="33131-129">Save the **manifest.xml** file.</span></span>

## <a name="configure-the-webpackconfigjs-file"></a><span data-ttu-id="33131-130">Configurer le fichier webpack.config.js</span><span class="sxs-lookup"><span data-stu-id="33131-130">Configure the webpack.config.js file</span></span>

<span data-ttu-id="33131-131">Le fichier **webpack.config.js** générera plusieurs chargeurs runtime.</span><span class="sxs-lookup"><span data-stu-id="33131-131">The **webpack.config.js** will build multiple runtime loaders.</span></span> <span data-ttu-id="33131-132">Vous devez le modifier pour charger uniquement le runtime JavaScript partagé via le fichier **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="33131-132">You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.</span></span>

1. <span data-ttu-id="33131-133">Démarrez Visual Studio Code, puis ouvrez le projet de complément Excel ou PowerPoint créé.</span><span class="sxs-lookup"><span data-stu-id="33131-133">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
2. <span data-ttu-id="33131-134">Ouvrez le fichier **webpack.config.js**.</span><span class="sxs-lookup"><span data-stu-id="33131-134">Open the **webpack.config.js** file.</span></span>
3. <span data-ttu-id="33131-135">Si votre fichier **webpack.config.js** a le code plug-in **functions.html**, supprimez-le.</span><span class="sxs-lookup"><span data-stu-id="33131-135">If your **webpack.config.js** file has the following **functions.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

4. <span data-ttu-id="33131-136">Si votre fichier **webpack.config.js** a le code plug-in **commands.html**, supprimez-le.</span><span class="sxs-lookup"><span data-stu-id="33131-136">If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

5. <span data-ttu-id="33131-137">Si votre projet utilisait les blocs **fonctions** ou **commandes**, ajoutez-les à la liste des blocs comme illustré par la suite (le code suivant sert si votre projet utilisait les deux blocs).</span><span class="sxs-lookup"><span data-stu-id="33131-137">If your project used either the **functions** or **commands** chunks, add them to the chunks list as shown next (the following code is for if your project used both chunks).</span></span>

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

6. <span data-ttu-id="33131-138">Enregistrez vos changements et reconstruisez le projet.</span><span class="sxs-lookup"><span data-stu-id="33131-138">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

> [!NOTE]
> <span data-ttu-id="33131-139">Si votre projet a le fichier **functions.html** ou le fichier **commands.html**, vous pouvez les supprimer.</span><span class="sxs-lookup"><span data-stu-id="33131-139">If your project has a **functions.html** file or **commands.html** file, they can be removed.</span></span> <span data-ttu-id="33131-140">Le fichier **taskpane.html** chargera le code **functions.js** et **commands.js** dans le runtime JavaScript partagé via les mises à jour webpack que vous venez d’effectuer.</span><span class="sxs-lookup"><span data-stu-id="33131-140">The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.</span></span>

## <a name="test-your-office-add-in-changes"></a><span data-ttu-id="33131-141">Tester les modifications apportées à votre complément Office</span><span class="sxs-lookup"><span data-stu-id="33131-141">Test your Office Add-in changes</span></span>

<span data-ttu-id="33131-142">Vous pouvez confirmer que vous utilisez correctement le runtime JavaScript partagé en utilisant les instructions suivantes.</span><span class="sxs-lookup"><span data-stu-id="33131-142">You can confirm that you are using the shared JavaScript runtime correctly by using the following instructions.</span></span>

1. <span data-ttu-id="33131-143">Ouvrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="33131-143">Open the **manifest.xml** file.</span></span>
2. <span data-ttu-id="33131-144">Recherchez la section `<Control xsi:type="Button" id="TaskpaneButton">`, puis modifiez le XML `<Action ...>` suivant.</span><span class="sxs-lookup"><span data-stu-id="33131-144">Find the `<Control xsi:type="Button" id="TaskpaneButton">` section and change the following `<Action ...>` XML.</span></span>
    
    <span data-ttu-id="33131-145">de :</span><span class="sxs-lookup"><span data-stu-id="33131-145">from:</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```
    
    <span data-ttu-id="33131-146">à :</span><span class="sxs-lookup"><span data-stu-id="33131-146">to:</span></span>
    
    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```
3. <span data-ttu-id="33131-147">Ouvrez le fichier **./src/commands/commands.js**.</span><span class="sxs-lookup"><span data-stu-id="33131-147">Open the **./src/commands/commands.js** file.</span></span>
4. <span data-ttu-id="33131-148">Remplacez la fonction **action** existante par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="33131-148">Replace the **action** function with the code below.</span></span> <span data-ttu-id="33131-149">Cette action mettra à jour la fonction pour ouvrir et modifier le bouton de volet des tâches pour incrémenter un compteur.</span><span class="sxs-lookup"><span data-stu-id="33131-149">This will update the function to open and modify the task pane button to increment a counter.</span></span> <span data-ttu-id="33131-150">L’ouverture et l’accès au volet des tâches DOM à partir d’une commande ne fonctionne qu’avec le runtime JavaScript partagé.</span><span class="sxs-lookup"><span data-stu-id="33131-150">Opening and accessing the task pane DOM from a command only works with the shared JavaScript runtime.</span></span>
    
    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

5. <span data-ttu-id="33131-151">Enregistrez vos changements et exécutez le projet.</span><span class="sxs-lookup"><span data-stu-id="33131-151">Save your changes and run the project.</span></span>

   ```command line
   npm start
   ```

<span data-ttu-id="33131-152">Chaque fois que vous sélectionnez le bouton de complément, il changera le texte du bouton **exécuter** par **accéder** et incrémentera un compteur après celui-ci.</span><span class="sxs-lookup"><span data-stu-id="33131-152">Each time you select the add-ins button, it will change the **run** button text to **go** and increment a counter after it.</span></span>

## <a name="runtime-lifetime"></a><span data-ttu-id="33131-153">Durée de vie de l’exécution</span><span class="sxs-lookup"><span data-stu-id="33131-153">Runtime lifetime</span></span>

<span data-ttu-id="33131-154">Lorsque vous ajoutez l’élément `Runtime`, vous spécifiez également une durée de vie ayant une valeur de `long` ou de `short`.</span><span class="sxs-lookup"><span data-stu-id="33131-154">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="33131-155">Configurez cette valeur sur `long` pour tirer parti de fonctionnalités telles que le démarrage de votre complément lorsque le document s’ouvre, continuer à exécuter un code après la fermeture du volet des tâches, ou utiliser CORS et DOM à partir de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="33131-155">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

>[!NOTE]
> <span data-ttu-id="33131-156">La valeur de la durée de vie par défaut est `short`, mais nous vous recommandons d’utiliser `long` dans les compléments Excel. Si vous avez défini votre runtime sur `short` dans cet exemple, votre complément Excel démarre lorsque vous appuyez sur l’un de vos boutons du ruban, mais il se peut qu’il se ferme une fois l’exécution de votre gestionnaire de ruban terminée.</span><span class="sxs-lookup"><span data-stu-id="33131-156">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="33131-157">De la même façon, le complément démarre lorsque le volet des tâches est ouvert, mais il se peut se fermer à la fermeture du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="33131-157">Similarly, your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

>[!NOTE]
> <span data-ttu-id="33131-158">Si votre complément inclut l’élément `Runtimes` dans le manifeste (nécessaire pour une exécution partagée), il utilise Internet Explorer 11 quelle que soit la version de Windows ou de Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="33131-158">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="33131-159">Pour plus d’informations, voir [Services d’exécution](../reference/manifest/runtimes.md).</span><span class="sxs-lookup"><span data-stu-id="33131-159">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="about-the-shared-javascript-runtime"></a><span data-ttu-id="33131-160">À propos du runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="33131-160">About the shared JavaScript runtime</span></span>

<span data-ttu-id="33131-161">Sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts.</span><span class="sxs-lookup"><span data-stu-id="33131-161">On Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="33131-162">Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="33131-162">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="33131-163">Vous pouvez toutefois configurer votre complément Office pour partager un code dans le même runtime JavaScript (également appelé runtime partagé).</span><span class="sxs-lookup"><span data-stu-id="33131-163">However, you can configure your Office Add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="33131-164">Vous pouvez ainsi améliorer la coordination dans votre complément et accéder au volet des tâches DOM et CORS à partir de toutes les parties de votre complément.</span><span class="sxs-lookup"><span data-stu-id="33131-164">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="33131-165">La configuration d’un runtime partagé permet les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="33131-165">Configuring a shared runtime enables the following scenarios.</span></span>

- <span data-ttu-id="33131-166">Votre complément Office peut utiliser des fonctionnalités d’interface utilisateur supplémentaires :</span><span class="sxs-lookup"><span data-stu-id="33131-166">Your Office Add-in can use additional UI features:</span></span>
    - [<span data-ttu-id="33131-167">Ajouter des raccourcis clavier personnalisés à votre complément Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="33131-167">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
    - [<span data-ttu-id="33131-168">Créer des onglets contextuels personnalisés dans des compléments Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="33131-168">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
    - [<span data-ttu-id="33131-169">Activer et désactiver des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="33131-169">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
    - [<span data-ttu-id="33131-170">Exécuter un cote dans votre complément Office lors de l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="33131-170">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
    - [<span data-ttu-id="33131-171">Afficher ou masquer le volet des tâches de votre complément Office</span><span class="sxs-lookup"><span data-stu-id="33131-171">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- <span data-ttu-id="33131-172">Pour les compléments Excel :</span><span class="sxs-lookup"><span data-stu-id="33131-172">For Excel add-ins:</span></span>
    - <span data-ttu-id="33131-173">Les fonctions personnalisées bénéficieront d'une prise en charge complète de CORS.</span><span class="sxs-lookup"><span data-stu-id="33131-173">Custom functions will have full CORS support.</span></span>
    - <span data-ttu-id="33131-174">Les fonctions personnalisées peuvent appeler les API Office.js pour lire les données d’un document feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="33131-174">Custom functions can call Office.js APIs to read spreadsheet document data.</span></span>

<span data-ttu-id="33131-175">Pour Office sur Windows, le runtime partagé requiert une instance de navigateur Microsoft Internet Explorer 11, comme expliqué dans [navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). De plus, les boutons affichés par votre complément sur le ruban s’exécutent dans le même runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="33131-175">For Office on Windows, the shared runtime requires a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="33131-176">L’image ci-après présente l'exécution des fonctions personnalisées, de interface utilisateur du ruban et du code du volet des tâches dans le même runtime JavaScript.</span><span class="sxs-lookup"><span data-stu-id="33131-176">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Diagramme d’une fonction personnalisée, d’un volet des tâches et des boutons de ruban s’exécutant tous dans un runtime partagé de navigateur Internet Explorer dans Excel](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a><span data-ttu-id="33131-178">Débogage</span><span class="sxs-lookup"><span data-stu-id="33131-178">Debugging</span></span>

<span data-ttu-id="33131-179">Lors de l’utilisation d’un runtime partagé, vous ne pouvez pas utiliser Visual Studio Code pour déboguer des fonctions personnalisées dans Excel sur Windows à cette date.</span><span class="sxs-lookup"><span data-stu-id="33131-179">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="33131-180">Vous devez utiliser les outils de développement à la place.</span><span class="sxs-lookup"><span data-stu-id="33131-180">You'll need to use developer tools instead.</span></span> <span data-ttu-id="33131-181">Pour plus d'informations, voir le [Débogage des compléments avec les outils de développement sur Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span><span class="sxs-lookup"><span data-stu-id="33131-181">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

### <a name="multiple-task-panes"></a><span data-ttu-id="33131-182">Multiples volets des tâches</span><span class="sxs-lookup"><span data-stu-id="33131-182">Multiple task panes</span></span>

<span data-ttu-id="33131-183">Ne concevez pas votre complément pour utiliser plusieurs volets des tâches si vous envisagez d’utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="33131-183">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="33131-184">Un runtime partagé prend uniquement en charge l’utilisation d’un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="33131-184">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="33131-185">Notez que tout volet des tâches sans `<TaskpaneID>` est considéré comme un volet des tâches différent.</span><span class="sxs-lookup"><span data-stu-id="33131-185">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="33131-186">Faites-nous part de vos commentaires</span><span class="sxs-lookup"><span data-stu-id="33131-186">Give us feedback</span></span>

<span data-ttu-id="33131-187">Nous aimerions connaître votre avis concernant cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="33131-187">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="33131-188">Si vous trouvez des bogues, des problèmes ou si vous avez des questions relatives à cette fonctionnalité, faites-le nous savoir en créant un problème GitHub dans le [référentiel Office-js](https://github.com/OfficeDev/office-js).</span><span class="sxs-lookup"><span data-stu-id="33131-188">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="33131-189">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="33131-189">See also</span></span>

- [<span data-ttu-id="33131-190">Appeler des API Excel à partir d'une fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="33131-190">Call Excel APIs from a custom function</span></span>](../excel/call-excel-apis-from-custom-function.md)
- [<span data-ttu-id="33131-191">Ajouter des raccourcis clavier personnalisés à votre complément Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="33131-191">Add custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
- [<span data-ttu-id="33131-192">Créer des onglets contextuels personnalisés dans des compléments Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="33131-192">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
- [<span data-ttu-id="33131-193">Activer et désactiver des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="33131-193">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
- [<span data-ttu-id="33131-194">Exécuter un cote dans votre complément Office lors de l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="33131-194">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
- [<span data-ttu-id="33131-195">Afficher ou masquer le volet des tâches de votre complément Office</span><span class="sxs-lookup"><span data-stu-id="33131-195">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- [<span data-ttu-id="33131-196">Tutoriel : Partager des données et des événements entre des fonctions personnalisées Excel et le volet Office</span><span class="sxs-lookup"><span data-stu-id="33131-196">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
