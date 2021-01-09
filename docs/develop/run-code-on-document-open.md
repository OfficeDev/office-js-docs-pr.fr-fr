---
title: Exécuter du code dans votre add-in Office à l’ouverture du document
description: Découvrez comment exécuter du code dans votre add-in Office à l’ouverture du document.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789227"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a><span data-ttu-id="fcb58-103">Exécuter du code dans votre add-in Office à l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="fcb58-103">Run code in your Office Add-in when the document opens</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="fcb58-104">Vous pouvez configurer votre add-in Office pour charger et exécuter du code dès que le document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="fcb58-104">You can configure your Office Add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="fcb58-105">Cela est utile si vous devez inscrire des handlers d’événements, pré-charger des données pour le volet Des tâches, synchroniser l’interface utilisateur ou effectuer d’autres tâches avant que le module ne soit visible.</span><span class="sxs-lookup"><span data-stu-id="fcb58-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="fcb58-106">Configurer votre add-in pour qu’il se charge à l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="fcb58-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="fcb58-107">Le code suivant configure votre add-in pour qu’il se charge et démarre l’exécution à l’ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="fcb58-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="fcb58-108">La `setStartupBehavior` méthode est asynchrone.</span><span class="sxs-lookup"><span data-stu-id="fcb58-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="fcb58-109">Configurer votre add-in pour qu’il n’y a aucun comportement de chargement à l’ouverture d’un document</span><span class="sxs-lookup"><span data-stu-id="fcb58-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="fcb58-110">Le code suivant configure votre add-in pour qu’il ne démarre pas lorsque le document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="fcb58-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="fcb58-111">Au lieu de cela, il démarre lorsque l’utilisateur l’engage d’une manière ou d’une autre, par exemple en choisissant un bouton de ruban ou en ouvrant le volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="fcb58-111">Instead, it will start when the user engages it in some way, such as choosing a ribbon button or opening the task pane.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="fcb58-112">Obtenir le comportement de chargement actuel</span><span class="sxs-lookup"><span data-stu-id="fcb58-112">Get the current load behavior</span></span>

<span data-ttu-id="fcb58-113">Pour déterminer le comportement de démarrage actuel, exécutez la fonction suivante, qui renvoie un `Office.StartupBehavior` objet.</span><span class="sxs-lookup"><span data-stu-id="fcb58-113">To determine what the current startup behavior is, run the following function, which returns an `Office.StartupBehavior` object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="fcb58-114">Comment exécuter du code à l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="fcb58-114">How to run code when the document opens</span></span>

<span data-ttu-id="fcb58-115">Lorsque votre add-in est configuré pour se charger à l’ouverture du document, il s’exécute immédiatement.</span><span class="sxs-lookup"><span data-stu-id="fcb58-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="fcb58-116">Le `Office.initialize` handler d’événements est appelé.</span><span class="sxs-lookup"><span data-stu-id="fcb58-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="fcb58-117">Placez votre code de démarrage dans `Office.initialize` le ou le `Office.onReady` handler d’événements.</span><span class="sxs-lookup"><span data-stu-id="fcb58-117">Place your startup code in the `Office.initialize` or `Office.onReady` event handler.</span></span>

<span data-ttu-id="fcb58-118">Le code de la feuille de calcul active montre comment inscrire un handler d’événements pour les événements de modification à partir de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="fcb58-118">The following Excel add-in code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="fcb58-119">Si vous configurez votre add-in pour qu’il se charge sur le document ouvert, ce code enregistre le handler d’événements lors de l’ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="fcb58-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="fcb58-120">Vous pouvez gérer les événements de modification avant l’ouverture du volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="fcb58-120">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}
```

<span data-ttu-id="fcb58-121">Le code de l’add-in PowerPoint suivant montre comment inscrire un handler d’événements pour les événements de modification de sélection à partir du document PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="fcb58-121">The following PowerPoint add-in code shows how to register an event handler for selection change events from the PowerPoint document.</span></span> <span data-ttu-id="fcb58-122">Si vous configurez votre add-in pour qu’il se charge sur le document ouvert, ce code enregistre le handler d’événements lors de l’ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="fcb58-122">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="fcb58-123">Vous pouvez gérer les événements de modification avant l’ouverture du volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="fcb58-123">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="see-also"></a><span data-ttu-id="fcb58-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fcb58-124">See also</span></span>

- [<span data-ttu-id="fcb58-125">Configurer votre add-in Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="fcb58-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="fcb58-126">Partager des données et des événements entre les fonctions personnalisées Excel et le didacticiel du volet Des tâches</span><span class="sxs-lookup"><span data-stu-id="fcb58-126">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="fcb58-127">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="fcb58-127">Work with Events using the Excel JavaScript API</span></span>](../excel/excel-add-ins-events.md)
