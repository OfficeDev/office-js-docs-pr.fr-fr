---
title: Exécuter du code dans votre complément Excel lorsque le document s’ouvre
description: Exécutez le code dans votre complément Excel lorsque le document s’ouvre.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: f544140a08a81c69aeb2fbdbe53afc098a738ad1
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/03/2020
ms.locfileid: "44275762"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens"></a><span data-ttu-id="ab583-103">Exécuter du code dans votre complément Excel lorsque le document s’ouvre</span><span class="sxs-lookup"><span data-stu-id="ab583-103">Run code in your Excel add-in when the document opens</span></span>

<span data-ttu-id="ab583-104">Vous pouvez configurer votre complément Excel de sorte qu’il charge et exécute le code dès que le document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="ab583-104">You can configure your Excel add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="ab583-105">Cette opération est utile si vous devez enregistrer des gestionnaires d’événements, précharger des données pour le volet Office, synchroniser l’interface utilisateur ou effectuer d’autres tâches avant que le complément ne soit visible.</span><span class="sxs-lookup"><span data-stu-id="ab583-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="ab583-106">Configurer votre complément pour qu’il se charge lors de l’ouverture du document</span><span class="sxs-lookup"><span data-stu-id="ab583-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="ab583-107">Le code suivant configure votre complément de sorte qu’il se charge et commence à s’exécuter à l’ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="ab583-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="ab583-108">La `setStartupBehavior` méthode est asynchrone.</span><span class="sxs-lookup"><span data-stu-id="ab583-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="ab583-109">Configurer votre complément pour aucun comportement de chargement sur le document ouvert</span><span class="sxs-lookup"><span data-stu-id="ab583-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="ab583-110">Le code suivant permet de configurer votre complément de sorte qu’il ne démarre pas lorsque le document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="ab583-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="ab583-111">Au lieu de cela, il démarre lorsque l’utilisateur l’engage (par exemple, en choisissant un bouton de ruban ou en ouvrant le volet Office).</span><span class="sxs-lookup"><span data-stu-id="ab583-111">Instead it will start when the user engages it in some way (such as choosing a ribbon button, or opening the task pane.)</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="ab583-112">Obtenir le comportement de chargement actuel</span><span class="sxs-lookup"><span data-stu-id="ab583-112">Get the current load behavior</span></span>

<span data-ttu-id="ab583-113">Pour déterminer le comportement de démarrage actuel, exécutez la fonction suivante, qui renvoie un objet Office. StartupBehavior.</span><span class="sxs-lookup"><span data-stu-id="ab583-113">To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="ab583-114">Procédure d’exécution du code lorsque le document s’ouvre</span><span class="sxs-lookup"><span data-stu-id="ab583-114">How to run code when the document opens</span></span>

<span data-ttu-id="ab583-115">Lorsque votre complément est configuré pour être chargé à l’ouverture d’un document, il s’exécutera immédiatement.</span><span class="sxs-lookup"><span data-stu-id="ab583-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="ab583-116">Le `Office.initialize` Gestionnaire d’événements est appelé.</span><span class="sxs-lookup"><span data-stu-id="ab583-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="ab583-117">Placez votre code de démarrage dans le `Office.initialize` Gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="ab583-117">Place your startup code in the `Office.initialize` event handler.</span></span>

<span data-ttu-id="ab583-118">Le code suivant montre comment enregistrer un gestionnaire d’événements pour les événements de modification à partir de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="ab583-118">The following code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="ab583-119">Si vous configurez le chargement de votre complément à l’ouverture du document, ce code enregistrera le gestionnaire d’événements à l’ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="ab583-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="ab583-120">Vous pouvez gérer les événements de modification avant l’ouverture du volet Office.</span><span class="sxs-lookup"><span data-stu-id="ab583-120">You can handle change events before the task pane is opened.</span></span>


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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

## <a name="see-also"></a><span data-ttu-id="ab583-121">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ab583-121">See also</span></span>

- [<span data-ttu-id="ab583-122">Partager des données et des événements entre des fonctions personnalisées Excel et un didacticiel de volet de tâches</span><span class="sxs-lookup"><span data-stu-id="ab583-122">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)