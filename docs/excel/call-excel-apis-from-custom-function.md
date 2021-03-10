---
title: Appeler des API JavaScript Excel à partir d’une fonction personnalisée
description: Découvrez les API JavaScript Excel que vous pouvez appeler à partir de votre fonction personnalisée.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613905"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a><span data-ttu-id="a9288-103">Appeler des API JavaScript Excel à partir d’une fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="a9288-103">Call Excel JavaScript APIs from a custom function</span></span>

<span data-ttu-id="a9288-104">Appelez des API JavaScript Excel à partir de vos fonctions personnalisées pour obtenir des données de plage et obtenir plus de contexte pour vos calculs.</span><span class="sxs-lookup"><span data-stu-id="a9288-104">Call Excel JavaScript APIs from your custom functions to get range data and obtain more context for your calculations.</span></span> <span data-ttu-id="a9288-105">L’appel d’API JavaScript Pour Excel via une fonction personnalisée peut être utile dans les cas de :</span><span class="sxs-lookup"><span data-stu-id="a9288-105">Calling Excel JavaScript APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="a9288-106">Une fonction personnalisée doit obtenir des informations d’Excel avant le calcul.</span><span class="sxs-lookup"><span data-stu-id="a9288-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="a9288-107">Ces informations peuvent inclure des propriétés de document, des formats de plage, des parties XML personnalisées, un nom de workbook ou d’autres informations spécifiques à Excel.</span><span class="sxs-lookup"><span data-stu-id="a9288-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="a9288-108">Une fonction personnalisée définira le format numérique de la cellule pour les valeurs de retour après le calcul.</span><span class="sxs-lookup"><span data-stu-id="a9288-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a9288-109">Pour appeler des API JavaScript Excel à partir de votre fonction personnalisée, vous devez utiliser un runtime JavaScript partagé.</span><span class="sxs-lookup"><span data-stu-id="a9288-109">To call Excel JavaScript APIs from your custom function, you'll need to use a shared JavaScript runtime.</span></span> <span data-ttu-id="a9288-110">Pour plus d’information, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="a9288-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="code-sample"></a><span data-ttu-id="a9288-111">Exemple de code</span><span class="sxs-lookup"><span data-stu-id="a9288-111">Code sample</span></span>

<span data-ttu-id="a9288-112">Pour appeler des API JavaScript Excel à partir d’une fonction personnalisée, vous avez d’abord besoin d’un contexte.</span><span class="sxs-lookup"><span data-stu-id="a9288-112">To call Excel JavaScript APIs from a custom function, you first need a context.</span></span> <span data-ttu-id="a9288-113">Utilisez [l’objet Excel.RequestContext](/javascript/api/excel/excel.requestcontext) pour obtenir un contexte.</span><span class="sxs-lookup"><span data-stu-id="a9288-113">Use the [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to get a context.</span></span> <span data-ttu-id="a9288-114">Utilisez ensuite le contexte pour appeler les API dont vous avez besoin dans le workbook.</span><span class="sxs-lookup"><span data-stu-id="a9288-114">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="a9288-115">L’exemple de code suivant montre comment utiliser pour obtenir une valeur à partir `Excel.RequestContext` d’une cellule dans le workbook.</span><span class="sxs-lookup"><span data-stu-id="a9288-115">The following code sample shows how to use `Excel.RequestContext` to get a value from a cell in the workbook.</span></span> <span data-ttu-id="a9288-116">Dans cet exemple, le paramètre est transmis dans la méthode `address` [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) de l’API JavaScript pour Excel et doit être entré sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="a9288-116">In this sample, the `address` parameter is passed into the Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) method and must be entered as a string.</span></span> <span data-ttu-id="a9288-117">Par exemple, la fonction personnalisée entrée dans l’interface utilisateur Excel doit suivre le modèle , où est l’adresse de la cellule à partir de laquelle récupérer `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` la valeur.</span><span class="sxs-lookup"><span data-stu-id="a9288-117">For example, the custom function entered into the Excel UI must follow the pattern `=CONTOSO.GETRANGEVALUE("A1")`, where `"A1"` is the address of the cell from which to retrieve the value.</span></span>

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a><span data-ttu-id="a9288-118">Limitations de l’appel d’API JavaScript pour Excel via une fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="a9288-118">Limitations of calling Excel JavaScript APIs through a custom function</span></span>

<span data-ttu-id="a9288-119">N’appelez pas les API JavaScript pour Excel à partir d’une fonction personnalisée qui modifie l’environnement d’Excel.</span><span class="sxs-lookup"><span data-stu-id="a9288-119">Don't call Excel JavaScript APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="a9288-120">Cela signifie que vos fonctions personnalisées ne doivent pas faire l’une des choses suivantes :</span><span class="sxs-lookup"><span data-stu-id="a9288-120">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="a9288-121">Insérer, supprimer ou mettre en forme des cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="a9288-121">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="a9288-122">Modifiez la valeur d’une autre cellule.</span><span class="sxs-lookup"><span data-stu-id="a9288-122">Change another cell's value.</span></span>
- <span data-ttu-id="a9288-123">Déplacer, renommer, supprimer ou ajouter des feuilles à un workbook.</span><span class="sxs-lookup"><span data-stu-id="a9288-123">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="a9288-124">Modifiez l’une des options d’environnement, telles que le mode de calcul ou les affichages d’écran.</span><span class="sxs-lookup"><span data-stu-id="a9288-124">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="a9288-125">Ajoutez des noms à un workbook.</span><span class="sxs-lookup"><span data-stu-id="a9288-125">Add names to a workbook.</span></span>
- <span data-ttu-id="a9288-126">Définissez des propriétés ou exécutez la plupart des méthodes.</span><span class="sxs-lookup"><span data-stu-id="a9288-126">Set properties or execute most methods.</span></span>

<span data-ttu-id="a9288-127">La modification d’Excel peut entraîner des performances médiocres, des dépassements de délai et des boucles infinies.</span><span class="sxs-lookup"><span data-stu-id="a9288-127">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="a9288-128">Les calculs de fonction personnalisée ne doivent pas s’exécuter pendant un recalcul Excel, car ils entraînent des résultats imprévisibles.</span><span class="sxs-lookup"><span data-stu-id="a9288-128">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="a9288-129">A la place, a apporter des modifications à Excel à partir du contexte d’un bouton de ruban ou d’un volet De tâches.</span><span class="sxs-lookup"><span data-stu-id="a9288-129">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a9288-130">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="a9288-130">Next steps</span></span>

- [<span data-ttu-id="a9288-131">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="a9288-131">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="a9288-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a9288-132">See also</span></span>

- [<span data-ttu-id="a9288-133">Partager des données et des événements entre les fonctions personnalisées Excel et le didacticiel du volet Des tâches</span><span class="sxs-lookup"><span data-stu-id="a9288-133">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="a9288-134">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="a9288-134">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
