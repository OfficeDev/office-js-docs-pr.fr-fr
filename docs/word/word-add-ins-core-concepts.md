---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Word
description: L’API JavaScript pour Word permet de créer des compléments pour Word.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293092"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="23b2f-103">Concepts fondamentaux de programmation avec l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="23b2f-104">Cet article décrit les concepts de base de l’utilisation de [l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md) pour créer des compléments pour Word 2016 ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="23b2f-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="23b2f-105">Referencing Office.js</span><span class="sxs-lookup"><span data-stu-id="23b2f-105">Referencing Office.js</span></span>

<span data-ttu-id="23b2f-106">Vous pouvez référencer Office.js à partir des emplacements suivants :</span><span class="sxs-lookup"><span data-stu-id="23b2f-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="23b2f-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` : utilisez cette ressource pour les compléments de production.</span><span class="sxs-lookup"><span data-stu-id="23b2f-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="23b2f-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` : utilisez cette ressource pour essayer les fonctionnalités en préversion.</span><span class="sxs-lookup"><span data-stu-id="23b2f-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="23b2f-109">Ensembles de conditions requises de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="23b2f-110">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="23b2f-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="23b2f-111">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si une application Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="23b2f-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="23b2f-112">Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour Word, voir [Ensembles de conditions requises de l’API JavaScript pour Word](../reference/requirement-sets/word-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="23b2f-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="23b2f-113">Exécution de compléments Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-113">Running Word add-ins</span></span>

<span data-ttu-id="23b2f-114">Pour exécuter votre complément, utilisez un gestionnaire d’événements `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="23b2f-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="23b2f-115">Pour plus d’informations sur l’initialisation du complément, voir [Présentation de l’API](../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="23b2f-115">For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

<span data-ttu-id="23b2f-116">Les compléments qui ciblent Word 2016 ou une version ultérieure peuvent utiliser les API propres à Word.</span><span class="sxs-lookup"><span data-stu-id="23b2f-116">Add-ins that target Word 2016 or later can use the Word-specific APIs.</span></span> <span data-ttu-id="23b2f-117">Ils transmettent la logique d’interaction avec Word en tant que fonction dans la `Word.run()` méthode.</span><span class="sxs-lookup"><span data-stu-id="23b2f-117">They pass the Word-interaction logic as a function into the `Word.run()` method.</span></span> <span data-ttu-id="23b2f-118">Pour en savoir plus sur la manière d’interagir avec le document Word dans ce modèle de programmation, consultez [Utilisation du modèle API propre à l’application](../develop/application-specific-api-model.md) .</span><span class="sxs-lookup"><span data-stu-id="23b2f-118">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about how to interact with the Word document in this programming model.</span></span>

<span data-ttu-id="23b2f-119">L’exemple suivant montre comment initialiser et exécuter un complément Word à l’aide de la `Word.run()` méthode.</span><span class="sxs-lookup"><span data-stu-id="23b2f-119">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a><span data-ttu-id="23b2f-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="23b2f-120">See also</span></span>

- [<span data-ttu-id="23b2f-121">Présentation de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-121">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="23b2f-122">Créer votre premier complément Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-122">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="23b2f-123">Didacticiel sur les compléments Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-123">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="23b2f-124">Référence d’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="23b2f-124">Word JavaScript API reference</span></span>](/javascript/api/word)
