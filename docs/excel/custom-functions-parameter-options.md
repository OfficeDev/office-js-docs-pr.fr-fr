---
ms.date: 07/01/2019
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, et bien plus encore.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: 9416653d697bdf36ca698271e00d9742ff0e75a9
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617043"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="a70d7-103">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a70d7-103">Custom functions parameter options</span></span>

<span data-ttu-id="a70d7-104">Les fonctions personnalisées peuvent être configurées avec de nombreuses options différentes pour les paramètres.</span><span class="sxs-lookup"><span data-stu-id="a70d7-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="a70d7-105">Paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="a70d7-105">Optional parameters</span></span>

<span data-ttu-id="a70d7-106">Alors que les paramètres réguliers sont obligatoires, les paramètres facultatifs ne le sont pas.</span><span class="sxs-lookup"><span data-stu-id="a70d7-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="a70d7-107">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="a70d7-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="a70d7-108">Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre.</span><span class="sxs-lookup"><span data-stu-id="a70d7-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="a70d7-109">Cette fonction apparaît sous `=CONTOSO.ADD(first, second, [third])` la forme dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a70d7-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="a70d7-110">Script</span><span class="sxs-lookup"><span data-stu-id="a70d7-110">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
CustomFunctions.associate("ADD", add);
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="a70d7-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="a70d7-111">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
CustomFunctions.associate("ADD", add);
```

---

> [!NOTE]
> <span data-ttu-id="a70d7-112">Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="a70d7-113">Cela signifie que les paramètres initialisés par défaut dans la machine à écrire ne fonctionnent pas comme prévu.</span><span class="sxs-lookup"><span data-stu-id="a70d7-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="a70d7-114">Par conséquent, n’utilisez pas `function add(first:number, second:number, third=0):number` la syntaxe car elle ne peut `third` pas être initialisée à 0.</span><span class="sxs-lookup"><span data-stu-id="a70d7-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="a70d7-115">À la place, utilisez la syntaxe de la machine à écrire comme indiqué dans l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="a70d7-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="a70d7-116">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qui se produit lorsque les paramètres facultatifs sont null.</span><span class="sxs-lookup"><span data-stu-id="a70d7-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="a70d7-117">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="a70d7-118">Si le `zipCode` paramètre est null, la valeur par défaut est définie `98052`sur.</span><span class="sxs-lookup"><span data-stu-id="a70d7-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="a70d7-119">Si le `dayOfWeek` paramètre est null, il est défini sur mercredi.</span><span class="sxs-lookup"><span data-stu-id="a70d7-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="a70d7-120">Script</span><span class="sxs-lookup"><span data-stu-id="a70d7-120">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="a70d7-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="a70d7-121">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string
{
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a><span data-ttu-id="a70d7-122">Paramètres de plage</span><span class="sxs-lookup"><span data-stu-id="a70d7-122">Range parameters</span></span>

<span data-ttu-id="a70d7-123">Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée.</span><span class="sxs-lookup"><span data-stu-id="a70d7-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="a70d7-124">Une fonction peut également renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="a70d7-124">A function can also return a range of data.</span></span> <span data-ttu-id="a70d7-125">Excel passe une plage de données de cellule sous la forme d’un tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="a70d7-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="a70d7-126">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a70d7-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="a70d7-127">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="a70d7-128">Notez que dans les métadonnées JSON pour cette fonction, la propriété `type` du paramètre est définie `matrix`sur.</span><span class="sxs-lookup"><span data-stu-id="a70d7-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

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

## <a name="invocation-parameter"></a><span data-ttu-id="a70d7-129">Paramètre invocation</span><span class="sxs-lookup"><span data-stu-id="a70d7-129">Invocation parameter</span></span>

<span data-ttu-id="a70d7-130">Chaque fonction personnalisée reçoit automatiquement un `invocation` argument en tant que dernier argument.</span><span class="sxs-lookup"><span data-stu-id="a70d7-130">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="a70d7-131">Cet argument peut être utilisé pour récupérer un contexte supplémentaire, comme l’adresse de la cellule d’appel.</span><span class="sxs-lookup"><span data-stu-id="a70d7-131">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="a70d7-132">Ou elle peut être utilisée pour envoyer des informations à Excel, comme un gestionnaire de fonctions pour [annuler une fonction](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="a70d7-132">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="a70d7-133">Même si aucun paramètre n’est déclaré, votre fonction personnalisée a ce paramètre.</span><span class="sxs-lookup"><span data-stu-id="a70d7-133">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="a70d7-134">Cet argument n’apparaît pas pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a70d7-134">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="a70d7-135">Si vous souhaitez utiliser `invocation` dans votre fonction personnalisée, déclarez-le comme dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="a70d7-135">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="a70d7-136">Dans l’exemple de code suivant, `invocation` le contexte est explicitement indiqué pour votre référence.</span><span class="sxs-lookup"><span data-stu-id="a70d7-136">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

<span data-ttu-id="a70d7-137">Le paramètre vous permet d’obtenir le contexte de la cellule d’appel, ce qui peut être utile dans certains scénarios, notamment [la découverte de l’adresse d’une cellule qui appelle une fonction personnalisée](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="a70d7-137">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="a70d7-138">Paramètre de contexte de la cellule d’adressage</span><span class="sxs-lookup"><span data-stu-id="a70d7-138">Addressing cell's context parameter</span></span>

<span data-ttu-id="a70d7-139">Dans certains cas, vous devez obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="a70d7-139">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="a70d7-140">Cela est utile dans les scénarios suivants:</span><span class="sxs-lookup"><span data-stu-id="a70d7-140">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="a70d7-141">Mise en forme des plages: utilisez l’adresse de la cellule comme clé pour stocker des informations dans [OfficeRuntime. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="a70d7-141">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="a70d7-142">Utilisez ensuite [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) dans Excel pour charger la clé à partir de l’élément `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-142">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="a70d7-143">Affichage de valeurs mises en cache : si votre fonction est utilisée en mode hors connexion, affichez les valeurs mises en cache à partir de l’élément `OfficeRuntime.storage` à l’aide de `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-143">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="a70d7-144">Rapprochement : utilisez l’adresse de la cellule pour découvrir la cellule d’origine afin de vous aider à réaliser un rapprochement lors du traitement.</span><span class="sxs-lookup"><span data-stu-id="a70d7-144">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="a70d7-145">Pour demander le contexte d’une cellule d’adressage dans une fonction, vous devez utiliser une fonction pour Rechercher l’adresse de la cellule, comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="a70d7-145">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="a70d7-146">Les informations relatives à l’adresse d’une cellule ne sont `@requiresAddress` exposées que si elles sont balisées dans les commentaires de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a70d7-146">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

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

<span data-ttu-id="a70d7-147">Par défaut, les valeurs renvoyées par une fonction `getAddress` ont le format suivant : `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-147">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="a70d7-148">Par exemple, si une fonction a été appelée à partir d’une feuille de calcul appelée Dépenses dans la cellule B2, la valeur renvoyée serait `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="a70d7-148">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a70d7-149">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="a70d7-149">Next steps</span></span>
<span data-ttu-id="a70d7-150">Découvrez comment [enregistrer l’État dans vos fonctions personnalisées](custom-functions-save-state.md) ou utiliser des [valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="a70d7-150">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a70d7-151">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a70d7-151">See also</span></span>

* [<span data-ttu-id="a70d7-152">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a70d7-152">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="a70d7-153">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a70d7-153">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a70d7-154">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a70d7-154">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a70d7-155">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a70d7-155">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="a70d7-156">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="a70d7-156">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a70d7-157">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a70d7-157">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
