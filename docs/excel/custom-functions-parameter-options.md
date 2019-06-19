---
ms.date: 06/17/2019
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, et bien plus encore.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: f20fd00cb751cc1ab258db6442785f67f3460817
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059880"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="d7c90-103">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d7c90-103">Custom functions parameter options</span></span>

<span data-ttu-id="d7c90-104">Les fonctions personnalisées peuvent être configurées avec de nombreuses options différentes pour les paramètres:</span><span class="sxs-lookup"><span data-stu-id="d7c90-104">Custom functions are configurable with many different options for parameters:</span></span>
- [<span data-ttu-id="d7c90-105">Paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="d7c90-105">Optional parameters</span></span>](#custom-functions-optional-parameters)
- [<span data-ttu-id="d7c90-106">Paramètres de plage</span><span class="sxs-lookup"><span data-stu-id="d7c90-106">Range parameters</span></span>](#range-parameters)
- [<span data-ttu-id="d7c90-107">Paramètre de contexte d’invocation</span><span class="sxs-lookup"><span data-stu-id="d7c90-107">Invocation context parameter</span></span>](#invocation-parameter)

## <a name="custom-functions-optional-parameters"></a><span data-ttu-id="d7c90-108">Paramètres facultatifs de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d7c90-108">Custom functions optional parameters</span></span>

<span data-ttu-id="d7c90-109">Alors que les paramètres réguliers sont obligatoires, les paramètres facultatifs ne le sont pas.</span><span class="sxs-lookup"><span data-stu-id="d7c90-109">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="d7c90-110">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="d7c90-110">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="d7c90-111">Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre.</span><span class="sxs-lookup"><span data-stu-id="d7c90-111">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="d7c90-112">Cette fonction apparaît sous `=CONTOSO.ADD(first, second, [third])` la forme dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d7c90-112">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third !== undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="d7c90-113">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qu’il se passe lorsque les paramètres facultatifs ne sont pas définis.</span><span class="sxs-lookup"><span data-stu-id="d7c90-113">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="d7c90-114">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="d7c90-114">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="d7c90-115">Si le `zipCode` paramètre n’est pas défini, la valeur par défaut est définie `98052`sur.</span><span class="sxs-lookup"><span data-stu-id="d7c90-115">If the `zipCode` parameter is undefined, the default value is set to `98052`.</span></span> <span data-ttu-id="d7c90-116">Si le paramètre`dayOfWeek` n’est pas défini, la valeur par défaut est définie à mercredi.</span><span class="sxs-lookup"><span data-stu-id="d7c90-116">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a><span data-ttu-id="d7c90-117">Paramètres de plage</span><span class="sxs-lookup"><span data-stu-id="d7c90-117">Range parameters</span></span>

<span data-ttu-id="d7c90-118">Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée.</span><span class="sxs-lookup"><span data-stu-id="d7c90-118">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="d7c90-119">Une fonction peut également renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="d7c90-119">A function can also return a range of data.</span></span> <span data-ttu-id="d7c90-120">Excel passe une plage de données de cellule sous la forme d’un tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="d7c90-120">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="d7c90-121">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d7c90-121">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="d7c90-122">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="d7c90-122">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="d7c90-123">Notez que dans les métadonnées JSON pour cette fonction, la propriété `type` du paramètre est définie `matrix`sur.</span><span class="sxs-lookup"><span data-stu-id="d7c90-123">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a><span data-ttu-id="d7c90-124">Paramètre invocation</span><span class="sxs-lookup"><span data-stu-id="d7c90-124">Invocation parameter</span></span>

<span data-ttu-id="d7c90-125">Chaque fonction personnalisée reçoit automatiquement un `invocation` argument en tant que dernier argument.</span><span class="sxs-lookup"><span data-stu-id="d7c90-125">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="d7c90-126">Cet argument peut être utilisé pour récupérer un contexte supplémentaire, comme l’adresse de la cellule d’appel.</span><span class="sxs-lookup"><span data-stu-id="d7c90-126">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="d7c90-127">Ou elle peut être utilisée pour envoyer des informations à Excel, comme un gestionnaire de fonctions pour [annuler une fonction](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="d7c90-127">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="d7c90-128">Même si aucun paramètre n’est déclaré, votre fonction personnalisée a ce paramètre.</span><span class="sxs-lookup"><span data-stu-id="d7c90-128">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="d7c90-129">Cet argument n’apparaît pas pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d7c90-129">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="d7c90-130">Si vous souhaitez utiliser `invocation` dans votre fonction personnalisée, déclarez-le comme dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="d7c90-130">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="d7c90-131">Dans l’exemple de code suivant, `invocation` le contexte est explicitement indiqué pour votre référence.</span><span class="sxs-lookup"><span data-stu-id="d7c90-131">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="d7c90-132">Le paramètre vous permet d’obtenir le contexte de la cellule d’appel, ce qui peut être utile dans certains scénarios, notamment [la découverte de l’adresse d’une cellule qui appelle une fonction personnalisée](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="d7c90-132">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="d7c90-133">Paramètre de contexte de la cellule d’adressage</span><span class="sxs-lookup"><span data-stu-id="d7c90-133">Addressing cell's context parameter</span></span>

<span data-ttu-id="d7c90-134">Dans certains cas, vous devez obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="d7c90-134">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="d7c90-135">Cela est utile dans les scénarios suivants:</span><span class="sxs-lookup"><span data-stu-id="d7c90-135">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="d7c90-136">Mise en forme des plages: utilisez l’adresse de la cellule comme clé pour stocker des informations dans [OfficeRuntime. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="d7c90-136">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="d7c90-137">Utilisez ensuite [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) dans Excel pour charger la clé à partir de l’élément `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="d7c90-137">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="d7c90-138">Affichage de valeurs mises en cache : si votre fonction est utilisée en mode hors connexion, affichez les valeurs mises en cache à partir de l’élément `OfficeRuntime.storage` à l’aide de `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="d7c90-138">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="d7c90-139">Rapprochement : utilisez l’adresse de la cellule pour découvrir la cellule d’origine afin de vous aider à réaliser un rapprochement lors du traitement.</span><span class="sxs-lookup"><span data-stu-id="d7c90-139">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="d7c90-140">Pour demander le contexte d’une cellule d’adressage dans une fonction, vous devez utiliser une fonction pour Rechercher l’adresse de la cellule, comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="d7c90-140">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="d7c90-141">Les informations relatives à l’adresse d’une cellule ne sont `@requiresAddress` exposées que si elles sont balisées dans les commentaires de la fonction.</span><span class="sxs-lookup"><span data-stu-id="d7c90-141">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
CustomFunctions.associate("GETADDRESS", getAddress);
```

<span data-ttu-id="d7c90-142">Par défaut, les valeurs renvoyées par une fonction `getAddress` ont le format suivant : `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="d7c90-142">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="d7c90-143">Par exemple, si une fonction a été appelée à partir d’une feuille de calcul appelée Dépenses dans la cellule B2, la valeur renvoyée serait `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="d7c90-143">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d7c90-144">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="d7c90-144">Next steps</span></span>
<span data-ttu-id="d7c90-145">Découvrez comment [enregistrer l’État dans vos fonctions personnalisées](custom-functions-save-state.md) ou utiliser des [valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="d7c90-145">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d7c90-146">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d7c90-146">See also</span></span>

* [<span data-ttu-id="d7c90-147">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d7c90-147">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="d7c90-148">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d7c90-148">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="d7c90-149">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d7c90-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d7c90-150">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d7c90-150">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="d7c90-151">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="d7c90-151">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="d7c90-152">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="d7c90-152">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
