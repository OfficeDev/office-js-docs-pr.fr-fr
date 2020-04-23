---
ms.date: 04/16/2020
title: Configurez votre complément Excel pour partager le runtime du navigateur (préversion)
ms.prod: excel
description: Configurez votre complément Excel pour partager le runtime du navigateur et exécuter le ruban, le volet des tâches et le code de fonction personnalisée dans le même runtime.
localization_priority: Priority
ms.openlocfilehash: f84de010787921eeb13205935b733ec36b2c3d37
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547247"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime-preview"></a><span data-ttu-id="13965-103">Configurez votre complément Excel pour utiliser un runtime JavaScript partagé (préversion).</span><span class="sxs-lookup"><span data-stu-id="13965-103">Configure your Excel add-in to use a shared JavaScript runtime (preview)</span></span>

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="13965-104">Lors de l’exécution d’Excel sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts.</span><span class="sxs-lookup"><span data-stu-id="13965-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="13965-105">Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="13965-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="13965-106">Vous pouvez toutefois configurer votre complément Excel pour partager un code dans un runtime JavaScript partagé.</span><span class="sxs-lookup"><span data-stu-id="13965-106">However, you can configure your Excel add-in to share code in a shared JavaScript runtime.</span></span> <span data-ttu-id="13965-107">Vous pouvez ainsi améliorer la coordination dans votre complément et accéder aux DOM et CORS à partir de toutes les parties de votre complément.</span><span class="sxs-lookup"><span data-stu-id="13965-107">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="13965-108">Il vous permet également d’exécuter un code lorsque le document s’ouvre ou pendant la fermeture du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="13965-108">It also enables you to run code when the document opens, or to run code while the task pane is closed.</span></span> <span data-ttu-id="13965-109">Si vous voulez configurer votre complément pour utiliser un runtime partagé, suivez les instructions contenues dans cet article.</span><span class="sxs-lookup"><span data-stu-id="13965-109">To configure your add-in to use a shared runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="13965-110">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="13965-110">Create the add-in project</span></span>

<span data-ttu-id="13965-111">Si vous démarrez un nouveau projet, suivez ces étapes pour utiliser le générateur Yeoman et créer le projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="13965-111">If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="13965-112">Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :</span><span class="sxs-lookup"><span data-stu-id="13965-112">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="13965-113">Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**</span><span class="sxs-lookup"><span data-stu-id="13965-113">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="13965-114">Choose a script type (Choisissez un type de script) :  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="13965-114">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="13965-115">What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**</span><span class="sxs-lookup"><span data-stu-id="13965-115">What do you want to name your add-in? **My Office Add-in**</span></span>

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

<span data-ttu-id="13965-117">Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="13965-117">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="13965-118">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="13965-118">Configure the manifest</span></span>

<span data-ttu-id="13965-119">Procédez comme suit pour configurer un projet nouveau ou existant de manière à utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="13965-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span>

1. <span data-ttu-id="13965-120">Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="13965-120">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="13965-121">Ouvrez le fichier **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="13965-121">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="13965-122">Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="13965-122">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="13965-123">La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé.</span><span class="sxs-lookup"><span data-stu-id="13965-123">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span> <span data-ttu-id="13965-124">L'ID de ressources est `ContosoAddin.Url`, faisant par la suite référence à une chaîne dans la section des ressources.</span><span class="sxs-lookup"><span data-stu-id="13965-124">The resid is `ContosoAddin.Url` which references a string in the resources section later.</span></span> <span data-ttu-id="13965-125">Vous pouvez utiliser n’importe quelle valeur d'ID de ressources souhaitée. Elle doit cependant correspondre à l'ID de ressources des autres éléments contenus dans les parties de votre complément.</span><span class="sxs-lookup"><span data-stu-id="13965-125">You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="13965-126">Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="13965-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="13965-127">Cet ID de ressources correspond à l'ID de ressources de `<Runtime>`.</span><span class="sxs-lookup"><span data-stu-id="13965-127">This resid matches the `<Runtime>` resid element.</span></span> <span data-ttu-id="13965-128">Veuillez noter que si vous ne disposez pas de fonctions personnalisées, vous n’aurez pas d'entrée de **Page**. Vous pouvez par conséquent ignorer cette étape.</span><span class="sxs-lookup"><span data-stu-id="13965-128">Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="13965-129">Dans la section `<DesktopFormFactor>`, changez la valeur **FunctionFile** de **Command.Url** pour utiliser **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="13965-129">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span> <span data-ttu-id="13965-130">Veuillez noter que si vous n'avez pas de commandes d'action, vous ne disposerez pas d'entrée **FunctionFile**. Vous pouvez par conséquent ignorer cette étape.</span><span class="sxs-lookup"><span data-stu-id="13965-130">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="13965-131">Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="13965-131">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="13965-132">Veuillez noter que si vous n'avez pas de volet des tâches, vous ne disposerez pas de l'action **ShowTaskPane**. Vous pouvez par conséquent ignorer cette étape.</span><span class="sxs-lookup"><span data-stu-id="13965-132">Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="13965-133">Ajoutez un nouvel **ID d’URL** pour **ContosoAddin.Url** pointant vers **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="13965-133">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="13965-134">Enregistrez vos changements et reconstruisez le projet.</span><span class="sxs-lookup"><span data-stu-id="13965-134">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a><span data-ttu-id="13965-135">Durée de vie de l’exécution</span><span class="sxs-lookup"><span data-stu-id="13965-135">Runtime lifetime</span></span>

<span data-ttu-id="13965-136">Lorsque vous ajoutez l’élément `Runtime`, vous spécifiez également une durée de vie ayant une valeur de `long` ou de `short`.</span><span class="sxs-lookup"><span data-stu-id="13965-136">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="13965-137">Configurez cette valeur sur `long` pour tirer parti de fonctionnalités telles que le démarrage de votre complément lorsque le document s’ouvre, continuer à exécuter un code après la fermeture du volet des tâches, ou utiliser CORS et DOM à partir de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="13965-137">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

<span data-ttu-id="13965-138">Si vous configurez cette valeur sur `short`, le complément se comportera comme le comportement par défaut.</span><span class="sxs-lookup"><span data-stu-id="13965-138">If you set this value to `short`, the add-in will behave similar to the default behavior.</span></span> <span data-ttu-id="13965-139">Le complément démarre lorsque l’un des boutons de votre ruban est pressé, mais il peut se fermer lorsque l’exécution de votre gestionnaire de ruban se termine.</span><span class="sxs-lookup"><span data-stu-id="13965-139">The add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="13965-140">De la même façon, le complément démarre lorsque le volet des tâches est ouvert, mais il se peut se fermer à la fermeture du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="13965-140">Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a><span data-ttu-id="13965-141">Multiples volets des tâches</span><span class="sxs-lookup"><span data-stu-id="13965-141">Multiple task panes</span></span>

<span data-ttu-id="13965-142">Ne concevez pas votre complément pour utiliser plusieurs volets des tâches si vous envisagez d’utiliser le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="13965-142">Don't design your add-in to use multiple task panes if you are planning to use shared runtime.</span></span> <span data-ttu-id="13965-143">Le runtime partagé prend uniquement en charge l’utilisation d’un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="13965-143">Shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="13965-144">Notez que tout volet des tâches sans `<TaskpaneID>` est considéré comme un volet des tâches différent.</span><span class="sxs-lookup"><span data-stu-id="13965-144">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="13965-145">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="13965-145">Next steps</span></span>

- <span data-ttu-id="13965-146">Lisez l’article [Appeler des API Microsoft Excel à partir d’une fonction personnalisée](call-excel-apis-from-custom-function.md) pour plus d’informations sur l’utilisation des API JavaScript Excel et des fonctions Excel personnalisées dans un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="13965-146">Read the [Call Excel APIs from a custom function](call-excel-apis-from-custom-function.md) article for details on using the Excel JavaScript APIs and custom Excel functions in a shared runtime.</span></span>
- <span data-ttu-id="13965-147">Découvrez l’exemple de modèles et de pratiques [Gérer le ruban et l’interface utilisateur du volet des tâches, puis exécuter le code sur un document ouvert](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) pour afficher un exemple plus complet de l’exécution JavaScript partagée en action.</span><span class="sxs-lookup"><span data-stu-id="13965-147">Explore the patterns-and-practices sample [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) to see a larger example of the shared JavaScript runtime in action.</span></span>

## <a name="see-also"></a><span data-ttu-id="13965-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="13965-148">See also</span></span>

- [<span data-ttu-id="13965-149">Vue d’ensemble : exécutez votre code de complément dans un runtime JavaScript partagé (préversion)</span><span class="sxs-lookup"><span data-stu-id="13965-149">Overview: Run your add-in code in a shared JavaScript runtime (preview)</span></span>](custom-functions-shared-overview.md)
