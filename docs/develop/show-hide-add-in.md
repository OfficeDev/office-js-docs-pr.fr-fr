---
title: Afficher ou masquer le volet Office de votre add-in Office
description: Découvrez comment masquer ou afficher par programme l’interface utilisateur d’un add-in pendant qu’il s’exécute en continu.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789229"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a><span data-ttu-id="0bd69-103">Afficher ou masquer le volet Office de votre add-in Office</span><span class="sxs-lookup"><span data-stu-id="0bd69-103">Show or hide the task pane of your Office Add-in</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="0bd69-104">Vous pouvez afficher le volet Office de votre add-in Office en appelant la `Office.addin.showAsTaskpane()` fonction.</span><span class="sxs-lookup"><span data-stu-id="0bd69-104">You can show the task pane of your Office Add-in by calling the `Office.addin.showAsTaskpane()` function.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="0bd69-105">Le code précédent suppose un scénario dans lequel il existe une feuille de calcul Excel nommée **CurrentQuarterSales**.</span><span class="sxs-lookup"><span data-stu-id="0bd69-105">The previous code assumes a scenario where there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="0bd69-106">Le add-in rend le volet Des tâches visible chaque fois que cette feuille de calcul est activée.</span><span class="sxs-lookup"><span data-stu-id="0bd69-106">The add-in will make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="0bd69-107">La méthode est un handler pour `onCurrentQuarter` l’événement [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) qui a été inscrit pour la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="0bd69-107">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) event which has been registered for the worksheet.</span></span>

<span data-ttu-id="0bd69-108">Vous pouvez également masquer le volet Des tâches en appelant la `Office.addin.hide()` fonction.</span><span class="sxs-lookup"><span data-stu-id="0bd69-108">You can also hide the task pane by calling the `Office.addin.hide()` function.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

<span data-ttu-id="0bd69-109">Le code précédent est un handler inscrit pour [l’événement Office.Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="0bd69-109">The previous code is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) event.</span></span>

## <a name="additional-details-on-showing-the-task-pane"></a><span data-ttu-id="0bd69-110">Détails supplémentaires sur l’affichage du volet Des tâches</span><span class="sxs-lookup"><span data-stu-id="0bd69-110">Additional details on showing the task pane</span></span>

<span data-ttu-id="0bd69-111">Lorsque vous appelez, Office affiche dans un volet Office le fichier que vous avez affecté en tant qu’ID de ressource ( ) du `Office.addin.showAsTaskpane()` `resid` volet Office.</span><span class="sxs-lookup"><span data-stu-id="0bd69-111">When you call `Office.addin.showAsTaskpane()`, Office will display in a task pane the file that you assigned as the resource ID (`resid`) value of the task pane.</span></span> <span data-ttu-id="0bd69-112">Cette valeur peut être affectée ou modifiée en ouvrant votre fichiermanifest.xmlet en `resid` le localisant à  `<SourceLocation>` l’intérieur de `<Action xsi:type="ShowTaskpane">` l’élément.</span><span class="sxs-lookup"><span data-stu-id="0bd69-112">This `resid` value can be assigned or changed by opening your **manifest.xml** file and locating `<SourceLocation>` inside the `<Action xsi:type="ShowTaskpane">` element.</span></span>
<span data-ttu-id="0bd69-113">(Pour plus [d’informations,](configure-your-add-in-to-use-a-shared-runtime.md) voir Configurer votre complément Office pour utiliser un runtime partagé.)</span><span class="sxs-lookup"><span data-stu-id="0bd69-113">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md) for additional details.)</span></span>

<span data-ttu-id="0bd69-114">Étant `Office.addin.showAsTaskpane()` donné qu’il s’agit d’une méthode asynchrone, votre code continuera à s’exécute jusqu’à ce que la fonction soit terminée.</span><span class="sxs-lookup"><span data-stu-id="0bd69-114">Since `Office.addin.showAsTaskpane()` is an asynchronous method, your code will continue running until the function is complete.</span></span> <span data-ttu-id="0bd69-115">Attendez cette fin avec le mot clé ou une méthode, en fonction de la `await` `then()` syntaxe JavaScript que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="0bd69-115">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span>

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a><span data-ttu-id="0bd69-116">Configurer votre add-in pour utiliser le runtime partagé</span><span class="sxs-lookup"><span data-stu-id="0bd69-116">Configure your add-in to use the shared runtime</span></span>

<span data-ttu-id="0bd69-117">Pour utiliser les `showAsTaskpane()` méthodes et les `hide()` méthodes, votre add-in doit utiliser le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="0bd69-117">To use the `showAsTaskpane()` and `hide()` methods, your add-in must use the shared runtime.</span></span> <span data-ttu-id="0bd69-118">Pour plus d’informations, voir [Configurer votre add-in Office pour utiliser un runtime partagé.](configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="0bd69-118">For more information, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="0bd69-119">Conservation des écouteurs d’état et d’événements</span><span class="sxs-lookup"><span data-stu-id="0bd69-119">Preservation of state and event listeners</span></span>

<span data-ttu-id="0bd69-120">Les `hide()` méthodes et les méthodes `showAsTaskpane()` modifient uniquement la *visibilité* du volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="0bd69-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="0bd69-121">Ils ne déchargent pas ou ne rechargent pas (ou réinitialisent son état).</span><span class="sxs-lookup"><span data-stu-id="0bd69-121">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="0bd69-122">Envisagez le scénario suivant : un volet Des tâches est conçu avec des onglets.</span><span class="sxs-lookup"><span data-stu-id="0bd69-122">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="0bd69-123">**L’onglet** Accueil est ouvert lors du premier lancement du module.</span><span class="sxs-lookup"><span data-stu-id="0bd69-123">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="0bd69-124">Supposons qu’un utilisateur ouvre l’onglet **Paramètres** et, plus tard, code dans les appels du volet Des tâches en `hide()` réponse à un événement.</span><span class="sxs-lookup"><span data-stu-id="0bd69-124">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="0bd69-125">Appels de code ultérieurs `showAsTaskpane()` en réponse à un autre événement.</span><span class="sxs-lookup"><span data-stu-id="0bd69-125">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="0bd69-126">Le volet Des tâches réapparaît et l’onglet **Paramètres** est toujours sélectionné.</span><span class="sxs-lookup"><span data-stu-id="0bd69-126">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![Capture d’écran du volet Des tâches avec quatre onglets étiquetés Accueil, Paramètres, Favoris et Comptes.](../images/TaskpaneWithTabs.png)

<span data-ttu-id="0bd69-128">En outre, tous les écouteurs d’événements inscrits dans le volet Des tâches continuent de s’exécuter même lorsque le volet Des tâches est masqué.</span><span class="sxs-lookup"><span data-stu-id="0bd69-128">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="0bd69-129">Envisagez le scénario suivant : le volet Des tâches possède un handler inscrit pour Excel et des événements pour une `Worksheet.onActivated` `Worksheet.onDeactivated` feuille nommée **Sheet1**.</span><span class="sxs-lookup"><span data-stu-id="0bd69-129">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="0bd69-130">Le handler activé entraîne l’apparition d’un point vert dans le volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="0bd69-130">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="0bd69-131">Le handler désactivé transforme le point en rouge (qui est son état par défaut).</span><span class="sxs-lookup"><span data-stu-id="0bd69-131">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="0bd69-132">Supposons alors que le code appelle `hide()` **lorsque la feuille Sheet1 n’est** pas activée et que le point est rouge.</span><span class="sxs-lookup"><span data-stu-id="0bd69-132">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="0bd69-133">Bien que le volet Des tâches soit masqué, **la feuille Sheet1** est activée.</span><span class="sxs-lookup"><span data-stu-id="0bd69-133">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="0bd69-134">Appels de code `showAsTaskpane()` ultérieurs en réponse à un événement.</span><span class="sxs-lookup"><span data-stu-id="0bd69-134">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="0bd69-135">Lorsque le volet Des tâches s’ouvre, le point est vert, car les écouteurs et les handlers d’événements s’ouvrent même si le volet Des tâches a été masqué.</span><span class="sxs-lookup"><span data-stu-id="0bd69-135">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

## <a name="handle-the-visibility-changed-event"></a><span data-ttu-id="0bd69-136">Gérer l’événement de changement de visibilité</span><span class="sxs-lookup"><span data-stu-id="0bd69-136">Handle the visibility changed event</span></span>

<span data-ttu-id="0bd69-137">Lorsque votre code modifie la visibilité du volet Office avec `showAsTaskpane()` `hide()` ou, Office déclenche `VisibilityModeChanged` l’événement.</span><span class="sxs-lookup"><span data-stu-id="0bd69-137">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="0bd69-138">Il peut être utile de gérer cet événement.</span><span class="sxs-lookup"><span data-stu-id="0bd69-138">It can be useful to handle this event.</span></span> <span data-ttu-id="0bd69-139">Par exemple, supposons que le volet Des tâches affiche une liste de toutes les feuilles dans un workbook.</span><span class="sxs-lookup"><span data-stu-id="0bd69-139">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="0bd69-140">Si une nouvelle feuille de calcul est ajoutée alors que le volet Des tâches est masqué, le fait de rendre le volet Des tâches visible n’ajoute pas en soi le nouveau nom de feuille de calcul à la liste.</span><span class="sxs-lookup"><span data-stu-id="0bd69-140">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="0bd69-141">Toutefois, votre code peut répondre à l’événement pour recharger la propriété Worksheet.name de toutes les feuilles de calcul de la `VisibilityModeChanged` collection [](/javascript/api/excel/excel.worksheet#name) [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) comme illustré dans l’exemple de code ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="0bd69-141">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="0bd69-142">Pour inscrire un handler pour l’événement, vous n’utilisez pas de méthode « add handler » comme vous le feriez dans la plupart des contextes JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="0bd69-142">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="0bd69-143">Au lieu de cela, il existe une fonction spéciale à laquelle vous passez votre handler : [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span><span class="sxs-lookup"><span data-stu-id="0bd69-143">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="0bd69-144">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="0bd69-144">The following is an example.</span></span> <span data-ttu-id="0bd69-145">Notez que `args.visibilityMode` la propriété est de type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span><span class="sxs-lookup"><span data-stu-id="0bd69-145">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="0bd69-146">La fonction renvoie une autre fonction *qui désinsère* le handler.</span><span class="sxs-lookup"><span data-stu-id="0bd69-146">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="0bd69-147">Voici un exemple simple, mais non robuste :</span><span class="sxs-lookup"><span data-stu-id="0bd69-147">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="0bd69-148">La `onVisibilityModeChanged` méthode est asynchrone et renvoie une promesse, ce qui signifie que votre code  doit attendre la réalisation de la promesse avant de pouvoir appeler le sous-enregistré.</span><span class="sxs-lookup"><span data-stu-id="0bd69-148">The `onVisibilityModeChanged` method is asynchronous and returns a promise, which means that your code needs to await the fulfillment of the promise before it can call the **deregister** handler.</span></span>

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="0bd69-149">La fonction d’agrégation est également asynchrone et renvoie une promesse.</span><span class="sxs-lookup"><span data-stu-id="0bd69-149">The deregister function is also asynchronous and returns a promise.</span></span> <span data-ttu-id="0bd69-150">Ainsi, si vous avez du code qui ne doit pas s’exécuter tant que l’agrégation n’est pas terminée, vous devez attendre la promesse renvoyée par la fonction d’agrégation.</span><span class="sxs-lookup"><span data-stu-id="0bd69-150">So, if you have code that should not run until after the deregistration is complete, then you should await the promise returned by the deregister function.</span></span>

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a><span data-ttu-id="0bd69-151">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0bd69-151">See also</span></span>

- [<span data-ttu-id="0bd69-152">Configurer votre add-in Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="0bd69-152">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="0bd69-153">Exécuter du code dans votre add-in Office à l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="0bd69-153">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
