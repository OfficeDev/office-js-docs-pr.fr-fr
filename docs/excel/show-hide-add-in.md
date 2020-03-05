---
title: Afficher ou masquer un complément Office dans un runtime partagé
description: Découvrez comment masquer ou afficher par programme l’interface utilisateur d’un complément pendant qu’il s’exécute en continu
ms.date: 03/02/2020
localization_priority: Normal
ms.openlocfilehash: c028823be165723cad3c0b314b53fe7e618188b2
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/04/2020
ms.locfileid: "42413798"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime-preview"></a><span data-ttu-id="f8aa4-103">Afficher ou masquer un complément Office dans un runtime partagé (aperçu)</span><span class="sxs-lookup"><span data-stu-id="f8aa4-103">Show or hide an Office Add-in in a shared runtime (preview)</span></span>

<span data-ttu-id="f8aa4-104">Un complément Office peut inclure n’importe lequel des éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="f8aa4-104">An Office Add-in can include any of the following parts:</span></span>

- <span data-ttu-id="f8aa4-105">Un volet Office</span><span class="sxs-lookup"><span data-stu-id="f8aa4-105">A task pane</span></span>
- <span data-ttu-id="f8aa4-106">Fichier de fonctions sans interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="f8aa4-106">A UI-less function file</span></span>
- <span data-ttu-id="f8aa4-107">Une fonction personnalisée Excel</span><span class="sxs-lookup"><span data-stu-id="f8aa4-107">An Excel custom function</span></span>

<span data-ttu-id="f8aa4-108">Par défaut, chaque partie s’exécute dans son propre Runtime JavaScript distinct, avec son propre objet global et ses propres variables globales.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-108">By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables.</span></span> 

<span data-ttu-id="f8aa4-109">Il est possible que des compléments avec deux ou plusieurs composants partagent un Runtime JavaScript commun.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-109">It's possible for add-ins with two or more parts to share a common JavaScript runtime.</span></span> <span data-ttu-id="f8aa4-110">Cette fonctionnalité d’exécution partagée permet aux nouvelles API d’aperçu de masquer et de rouvrir le volet Office pendant l’exécution du complément.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-110">This shared runtime feature enables new preview APIs that hide and reopen the task pane while the add-in runs.</span></span>

> [!INCLUDE [Information about using preview APIs](../includes/excel-shared-runtime-preview-note.md)]

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="f8aa4-111">Configurer un complément pour utiliser un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="f8aa4-111">Configure an add-in to use a shared runtime</span></span>

<span data-ttu-id="f8aa4-112">Pour configurer le complément afin qu’il utilise un runtime partagé, reportez-vous à la rubrique [Configure Your Office Add-in use a Shared Runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="f8aa4-112">To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="show-and-hide-the-task-pane"></a><span data-ttu-id="f8aa4-113">Afficher et masquer le volet Office</span><span class="sxs-lookup"><span data-stu-id="f8aa4-113">Show and hide the task pane</span></span>

<span data-ttu-id="f8aa4-114">Les nouvelles API se trouvent dans `Office.addin` la propriété.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-114">The new APIs are in the `Office.addin` property.</span></span> <span data-ttu-id="f8aa4-115">Pour afficher le volet Office, votre code appelle `Office.addin.showAsTaskpane()`.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-115">To show the task pane, your code calls `Office.addin.showAsTaskpane()`.</span></span> <span data-ttu-id="f8aa4-116">Office affiche dans un volet des tâches la page que vous avez affectée à l’ID de`resid`ressource () pour le volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-116">Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane.</span></span> <span data-ttu-id="f8aa4-117">Il s’agit `resid` du que vous avez affecté `<SourceLocation>` à l `<Action xsi:type="ShowTaskpane">` 'du dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-117">This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest.</span></span> <span data-ttu-id="f8aa4-118">(Consultez [la rubrique Configure Your Office Add-in to use a Shared Runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span><span class="sxs-lookup"><span data-stu-id="f8aa4-118">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span></span>

<span data-ttu-id="f8aa4-119">Il s’agit d’une méthode asynchrone, de sorte que votre code doit l’attendre lorsque le code suivant ne doit pas s’exécuter tant qu’il n’est pas terminé.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-119">This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes.</span></span> <span data-ttu-id="f8aa4-120">Attendez la fin de cette opération avec `await` le mot clé `then()` ou une méthode, en fonction de la syntaxe JavaScript que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-120">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span> <span data-ttu-id="f8aa4-121">Voici un exemple de feuille de calcul Excel nommée **CurrentQuarterSales**.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-121">The following assumes that there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="f8aa4-122">Le complément doit faire apparaître le volet Office chaque fois que cette feuille de calcul est activée.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-122">The add-in should make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="f8aa4-123">La méthode `onCurrentQuarter` est un gestionnaire pour l’événement [Office. Worksheet. onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) qui a été enregistré pour la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-123">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="f8aa4-124">Pour masquer le volet Office, votre code appelle `Office.addin.hide()`.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-124">To hide the task pane, your code calls `Office.addin.hide()`.</span></span> <span data-ttu-id="f8aa4-125">L’exemple suivant est un gestionnaire inscrit pour l’événement [Office. Worksheet. onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) .</span><span class="sxs-lookup"><span data-stu-id="f8aa4-125">The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="f8aa4-126">Conservation des écouteurs d’État et d’événement</span><span class="sxs-lookup"><span data-stu-id="f8aa4-126">Preservation of state and event listeners</span></span>

<span data-ttu-id="f8aa4-127">Les `hide()` méthodes `showAsTaskpane()` et modifient uniquement la *visibilité* du volet Office.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-127">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="f8aa4-128">Ils ne déchargent pas ou ne le rechargent pas (ou réinitialisent son état).</span><span class="sxs-lookup"><span data-stu-id="f8aa4-128">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="f8aa4-129">Prenons le scénario suivant : un volet Office est conçu avec des onglets.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-129">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="f8aa4-130">L’onglet **Accueil** est ouvert lors du premier lancement du complément.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-130">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="f8aa4-131">Supposons qu’un utilisateur ouvre l’onglet **paramètres** et, plus tard, le code dans `hide()` les appels de volet de tâches en réponse à un événement.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-131">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="f8aa4-132">Toujours des appels `showAsTaskpane()` de code plus récents en réponse à un autre événement.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-132">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="f8aa4-133">Le volet des tâches réapparaît et l’onglet **paramètres** est toujours sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-133">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![Capture d’écran de volet de tâches qui comporte quatre onglets intitulé Accueil, paramètres, favoris et comptes.](../images/TaskpaneWithTabs.png)

<span data-ttu-id="f8aa4-135">De plus, tout écouteur d’événement enregistré dans le volet Office continue de s’exécuter même si le volet Office est masqué.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-135">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="f8aa4-136">Prenons le scénario suivant : le volet Office dispose d’un gestionnaire enregistré pour Excel `Worksheet.onActivated` et `Worksheet.onDeactivated` des événements pour une feuille nommée **Sheet1**.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-136">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="f8aa4-137">Le gestionnaire activé provoque l’affichage d’un point vert dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-137">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="f8aa4-138">Le gestionnaire désactivé transforme le point rouge (il s’agit de son état par défaut).</span><span class="sxs-lookup"><span data-stu-id="f8aa4-138">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="f8aa4-139">Supposons que le code `hide()` appelle lorsque la **feuille Sheet1** n’est pas activée et que le point est rouge.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-139">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="f8aa4-140">Lorsque le volet Office est masqué, la **feuille Sheet1** est activée.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-140">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="f8aa4-141">Appels `showAsTaskpane()` de code ultérieurs en réponse à un événement.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-141">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="f8aa4-142">Lorsque le volet Office s’ouvre, le point est vert car les écouteurs et gestionnaires d’événements ont été exécutés même si le volet Office a été masqué.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-142">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

### <a name="handle-visibility-changed-event"></a><span data-ttu-id="f8aa4-143">Événement de modification de la visibilité des handles</span><span class="sxs-lookup"><span data-stu-id="f8aa4-143">Handle visibility changed event</span></span>

<span data-ttu-id="f8aa4-144">Lorsque votre code modifie la visibilité du volet Office avec `showAsTaskpane()` ou `hide()`, Office déclenche l' `VisibilityModeChanged` événement.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-144">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="f8aa4-145">Il peut être utile de gérer cet événement.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-145">It can be useful to handle this event.</span></span> <span data-ttu-id="f8aa4-146">Par exemple, supposons que le volet Office affiche une liste de toutes les feuilles dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-146">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="f8aa4-147">Si une nouvelle feuille de calcul est ajoutée alors que le volet Office est masqué, le fait de rendre le volet Office visible ne lui permet pas d’ajouter le nouveau nom de feuille de calcul à la liste.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-147">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="f8aa4-148">Toutefois, votre code peut répondre à `VisibilityModeChanged` l’événement pour recharger la propriété [Worksheet.Name](/javascript/api/excel/excel.worksheet#name) de toutes les feuilles de calcul dans la collection [Workbook. Worksheets](/javascript/api/excel/excel.workbook#worksheets) , comme illustré dans l’exemple de code ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-148">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="f8aa4-149">Pour enregistrer un gestionnaire pour l’événement, n’utilisez pas la méthode « Add Handler » comme vous le feriez dans la plupart des contextes JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-149">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="f8aa4-150">Au lieu de cela, il existe une fonction spéciale à laquelle vous transmettez votre gestionnaire : [Office. AddIn. onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span><span class="sxs-lookup"><span data-stu-id="f8aa4-150">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="f8aa4-151">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-151">The following is an example.</span></span> <span data-ttu-id="f8aa4-152">Notez que la `args.visibilityMode` propriété est de type [visibilityMode](/javascript/api/office/office.visibilitymode).</span><span class="sxs-lookup"><span data-stu-id="f8aa4-152">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="f8aa4-153">La fonction renvoie une autre fonction qui *annule l’enregistrement* du gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-153">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="f8aa4-154">Voici un exemple simple, mais non fiable :</span><span class="sxs-lookup"><span data-stu-id="f8aa4-154">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="f8aa4-155">La `onVisibilityModeChanged` méthode est asynchrone, ce qui signifie que si votre code \*\* appelle le gestionnaire d' `onVisibilityModeChanged` annulation qui est renvoyé, vous `onVisibilityModeChanged` devez vous assurer qu’il a été terminé avant d’appeler le gestionnaire d’annulation.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-155">The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler.</span></span> <span data-ttu-id="f8aa4-156">Pour ce faire, vous pouvez utiliser le `await` mot clé sur l’appel de la méthode, comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-156">One way to do that is to use the `await` keyword on the method call as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="f8aa4-157">Si vous souhaitez utiliser uniquement JavaScript pre-ES2015, votre code peut utiliser la `then` méthode pour attendre que l’objet promesse renvoyé ait été résolu et affecter la fonction renvoyée à une variable globale comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-157">If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="f8aa4-158">La fonction de désinscription est elle-même asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-158">The deregister function is itself asynchronous.</span></span> <span data-ttu-id="f8aa4-159">Par conséquent, si vous avez du code qui ne doit pas s’exécuter jusqu’à ce que la désinscription soit terminée, la fonction de désinscription doit également être attendue avec le `await` mot `then` clé ou une méthode, comme dans les exemples suivants.</span><span class="sxs-lookup"><span data-stu-id="f8aa4-159">So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.</span></span>

<span data-ttu-id="f8aa4-160">Pour annuler l’inscription du gestionnaire :</span><span class="sxs-lookup"><span data-stu-id="f8aa4-160">To deregister the handler:</span></span>

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
