---
title: Appeler des API Microsoft Excel à partir d’une fonction personnalisée
description: Découvrez les API Microsoft Excel que vous pouvez appeler à partir de votre fonction personnalisée.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a24cdfba2d79b6e2ad165765d22cd77743047d34
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217878"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a><span data-ttu-id="3e09e-103">Appeler des API Microsoft Excel à partir d’une fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="3e09e-103">Call Microsoft Excel APIs from a custom function</span></span>

<span data-ttu-id="3e09e-104">Appelez les API Excel Office. js à partir de vos fonctions personnalisées pour obtenir des données de plage et obtenir davantage de contexte pour vos calculs.</span><span class="sxs-lookup"><span data-stu-id="3e09e-104">Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.</span></span>

<span data-ttu-id="3e09e-105">L’appel des API Office. js via une fonction personnalisée peut être utile dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="3e09e-105">Calling Office.js APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="3e09e-106">Une fonction personnalisée doit obtenir des informations à partir d’Excel avant le calcul.</span><span class="sxs-lookup"><span data-stu-id="3e09e-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="3e09e-107">Ces informations peuvent inclure des propriétés de document, des formats de plage, des parties XML personnalisées, un nom de classeur ou d’autres informations spécifiques à Excel.</span><span class="sxs-lookup"><span data-stu-id="3e09e-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="3e09e-108">Une fonction personnalisée définit le format numérique de la cellule pour les valeurs renvoyées après le calcul.</span><span class="sxs-lookup"><span data-stu-id="3e09e-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="code-sample"></a><span data-ttu-id="3e09e-109">Exemple de code</span><span class="sxs-lookup"><span data-stu-id="3e09e-109">Code sample</span></span>

<span data-ttu-id="3e09e-110">Pour appeler les API Office. js, vous avez d’abord besoin d’un contexte.</span><span class="sxs-lookup"><span data-stu-id="3e09e-110">To call into the Office.js APIs you first need a context.</span></span> <span data-ttu-id="3e09e-111">Utilisez l' `Excel.RequestContext` objet pour obtenir un contexte.</span><span class="sxs-lookup"><span data-stu-id="3e09e-111">Use the `Excel.RequestContext` object to get a context.</span></span> <span data-ttu-id="3e09e-112">Ensuite, utilisez le contexte pour appeler les API dont vous avez besoin dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="3e09e-112">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="3e09e-113">L’exemple de code suivant montre comment obtenir une plage de valeurs du classeur.</span><span class="sxs-lookup"><span data-stu-id="3e09e-113">The following code sample shows how to get a range of values from the workbook.</span></span>

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a><span data-ttu-id="3e09e-114">Limitations de l’appel d’Office. js via une fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="3e09e-114">Limitations of calling Office.js through a custom function</span></span>

<span data-ttu-id="3e09e-115">N’appelez pas les API Office. js à partir d’une fonction personnalisée qui modifie l’environnement d’Excel.</span><span class="sxs-lookup"><span data-stu-id="3e09e-115">Don't call Office.js APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="3e09e-116">Cela signifie que vos fonctions personnalisées ne doivent pas effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="3e09e-116">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="3e09e-117">Insérer, supprimer ou mettre en forme des cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="3e09e-117">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="3e09e-118">Modifier la valeur d’une autre cellule.</span><span class="sxs-lookup"><span data-stu-id="3e09e-118">Change another cell's value.</span></span>
- <span data-ttu-id="3e09e-119">Déplacer, renommer, supprimer ou ajouter des feuilles dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="3e09e-119">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="3e09e-120">Modifier les options d’environnement, telles que le mode de calcul ou les affichages d’écran.</span><span class="sxs-lookup"><span data-stu-id="3e09e-120">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="3e09e-121">Ajouter des noms à un classeur.</span><span class="sxs-lookup"><span data-stu-id="3e09e-121">Add names to a workbook.</span></span>
- <span data-ttu-id="3e09e-122">Définir des propriétés ou exécuter la plupart des méthodes.</span><span class="sxs-lookup"><span data-stu-id="3e09e-122">Set properties or execute most methods.</span></span>

<span data-ttu-id="3e09e-123">La modification d’Excel peut entraîner une dégradation des performances, des délais et des boucles infinies.</span><span class="sxs-lookup"><span data-stu-id="3e09e-123">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="3e09e-124">Les calculs de fonctions personnalisées ne doivent pas s’exécuter lorsqu’un recalcul Excel a lieu, car cela entraînera des résultats imprévisibles.</span><span class="sxs-lookup"><span data-stu-id="3e09e-124">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="3e09e-125">Au lieu de cela, modifiez Excel à partir du contexte d’un bouton de ruban ou d’un volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="3e09e-125">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="3e09e-126">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="3e09e-126">Next steps</span></span>

- [<span data-ttu-id="3e09e-127">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="3e09e-127">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="3e09e-128">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3e09e-128">See also</span></span>

- [<span data-ttu-id="3e09e-129">Partager des données et des événements entre des fonctions personnalisées Excel et un didacticiel de volet de tâches</span><span class="sxs-lookup"><span data-stu-id="3e09e-129">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)