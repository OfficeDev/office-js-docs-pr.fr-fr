---
ms.date: 11/06/2020
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, et bien plus encore.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: 0a803a4d41354530584b25d2bf9df944af430909
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071619"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="d3da7-103">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d3da7-103">Custom functions parameter options</span></span>

<span data-ttu-id="d3da7-104">Les fonctions personnalisées peuvent être configurées avec de nombreuses options de paramètres différentes.</span><span class="sxs-lookup"><span data-stu-id="d3da7-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="d3da7-105">Paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="d3da7-105">Optional parameters</span></span>

<span data-ttu-id="d3da7-106">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="d3da7-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="d3da7-107">Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre.</span><span class="sxs-lookup"><span data-stu-id="d3da7-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="d3da7-108">Cette fonction apparaît sous la forme `=CONTOSO.ADD(first, second, [third])` dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d3da7-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="d3da7-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="d3da7-109">JavaScript</span></span>](#tab/javascript)

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
```

#### <a name="typescript"></a>[<span data-ttu-id="d3da7-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="d3da7-110">TypeScript</span></span>](#tab/typescript)

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
```

---

> [!NOTE]
> <span data-ttu-id="d3da7-111">Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null` .</span><span class="sxs-lookup"><span data-stu-id="d3da7-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="d3da7-112">Cela signifie que les paramètres initialisés par défaut dans la machine à écrire ne fonctionnent pas comme prévu.</span><span class="sxs-lookup"><span data-stu-id="d3da7-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="d3da7-113">N’utilisez pas la syntaxe `function add(first:number, second:number, third=0):number` car elle ne peut pas `third` être initialisée à 0.</span><span class="sxs-lookup"><span data-stu-id="d3da7-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="d3da7-114">À la place, utilisez la syntaxe de la machine à écrire comme indiqué dans l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="d3da7-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="d3da7-115">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se passe lorsque les paramètres facultatifs sont null.</span><span class="sxs-lookup"><span data-stu-id="d3da7-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="d3da7-116">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="d3da7-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="d3da7-117">Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052` .</span><span class="sxs-lookup"><span data-stu-id="d3da7-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="d3da7-118">Si le `dayOfWeek` paramètre est null, il est défini sur mercredi.</span><span class="sxs-lookup"><span data-stu-id="d3da7-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="d3da7-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="d3da7-119">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
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

#### <a name="typescript"></a>[<span data-ttu-id="d3da7-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="d3da7-120">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
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

## <a name="range-parameters"></a><span data-ttu-id="d3da7-121">Paramètres de plage</span><span class="sxs-lookup"><span data-stu-id="d3da7-121">Range parameters</span></span>

<span data-ttu-id="d3da7-122">Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée.</span><span class="sxs-lookup"><span data-stu-id="d3da7-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="d3da7-123">Une fonction peut également renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="d3da7-123">A function can also return a range of data.</span></span> <span data-ttu-id="d3da7-124">Excel passe une plage de données de cellule sous la forme d’un tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="d3da7-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="d3da7-125">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d3da7-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="d3da7-126">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="d3da7-126">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="d3da7-127">Notez que dans les métadonnées JSON pour cette fonction, la propriété du paramètre `type` est définie sur `matrix` .</span><span class="sxs-lookup"><span data-stu-id="d3da7-127">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a><span data-ttu-id="d3da7-128">Paramètres répétitifs</span><span class="sxs-lookup"><span data-stu-id="d3da7-128">Repeating parameters</span></span>

<span data-ttu-id="d3da7-129">Un paramètre extensible permet à un utilisateur d’entrer une série d’arguments facultatifs dans une fonction.</span><span class="sxs-lookup"><span data-stu-id="d3da7-129">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="d3da7-130">Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3da7-130">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="d3da7-131">Si le nom du paramètre se termine par un nombre, le numéro de chaque argument augmente de manière incrémentielle, par exemple `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="d3da7-131">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="d3da7-132">Cela correspond à la Convention utilisée pour les fonctions Excel intégrées.</span><span class="sxs-lookup"><span data-stu-id="d3da7-132">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="d3da7-133">La fonction suivante additionne le total des nombres, des adresses de cellules, ainsi que des plages, si elles sont entrées.</span><span class="sxs-lookup"><span data-stu-id="d3da7-133">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

<span data-ttu-id="d3da7-134">Cette fonction s’affiche `=CONTOSO.ADD([operands], [operands]...)` dans le classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="d3da7-134">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="d3da7-135">Paramètre extensible à valeur unique</span><span class="sxs-lookup"><span data-stu-id="d3da7-135">Repeating single value parameter</span></span>

<span data-ttu-id="d3da7-136">Un paramètre de valeur unique extensible permet de transmettre plusieurs valeurs uniques.</span><span class="sxs-lookup"><span data-stu-id="d3da7-136">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="d3da7-137">Par exemple, l’utilisateur peut entrer ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="d3da7-137">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="d3da7-138">L’exemple suivant montre comment déclarer un seul paramètre de valeur.</span><span class="sxs-lookup"><span data-stu-id="d3da7-138">The following sample shows how to declare a single value parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a><span data-ttu-id="d3da7-139">Paramètre de plage unique</span><span class="sxs-lookup"><span data-stu-id="d3da7-139">Single range parameter</span></span>

<span data-ttu-id="d3da7-140">Un paramètre de plage unique n’est pas techniquement un paramètre répétitif, mais il est inclus ici, car la déclaration est très similaire aux paramètres répétitifs.</span><span class="sxs-lookup"><span data-stu-id="d3da7-140">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="d3da7-141">Il apparaîtrait à l’utilisateur sous la forme ADD (a2 : B3) où une seule plage est passée d’Excel.</span><span class="sxs-lookup"><span data-stu-id="d3da7-141">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="d3da7-142">L’exemple suivant montre comment déclarer un paramètre de plage unique.</span><span class="sxs-lookup"><span data-stu-id="d3da7-142">The following sample shows how to declare a single range parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a><span data-ttu-id="d3da7-143">Paramètre de plage extensible</span><span class="sxs-lookup"><span data-stu-id="d3da7-143">Repeating range parameter</span></span>

<span data-ttu-id="d3da7-144">Un paramètre de plage extensible permet de transmettre plusieurs plages ou nombres.</span><span class="sxs-lookup"><span data-stu-id="d3da7-144">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="d3da7-145">Par exemple, l’utilisateur peut entrer ADD (5, B2, C3, 8, E5 : E8).</span><span class="sxs-lookup"><span data-stu-id="d3da7-145">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="d3da7-146">Les plages extensibles sont généralement spécifiées avec le type `number[][][]` comme il s’agit de matrices en trois dimensions.</span><span class="sxs-lookup"><span data-stu-id="d3da7-146">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="d3da7-147">Pour un exemple, reportez-vous à l’exemple principal ci-dessous pour les paramètres de répétition (paramètres #repeating).</span><span class="sxs-lookup"><span data-stu-id="d3da7-147">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="d3da7-148">Déclaration de paramètres répétitifs</span><span class="sxs-lookup"><span data-stu-id="d3da7-148">Declaring repeating parameters</span></span>
<span data-ttu-id="d3da7-149">Dans la machine à écrire, indiquez que le paramètre est à plusieurs dimensions.</span><span class="sxs-lookup"><span data-stu-id="d3da7-149">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="d3da7-150">Par exemple,  `ADD(values: number[])` un tableau à une dimension indiquerait `ADD(values:number[][])` un tableau à deux dimensions, et ainsi de suite.</span><span class="sxs-lookup"><span data-stu-id="d3da7-150">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="d3da7-151">En JavaScript, utilisez `@param values {number[]}` pour les tableaux à une dimension, `@param <name> {number[][]}` pour les tableaux à deux dimensions, et ainsi de suite pour d’autres dimensions.</span><span class="sxs-lookup"><span data-stu-id="d3da7-151">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="d3da7-152">Pour le format JSON dynamique, vérifiez que votre paramètre est spécifié en tant que `"repeating": true` dans votre fichier JSON, et vérifiez que vos paramètres sont marqués comme `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="d3da7-152">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="d3da7-153">Paramètre invocation</span><span class="sxs-lookup"><span data-stu-id="d3da7-153">Invocation parameter</span></span>

<span data-ttu-id="d3da7-154">Chaque fonction personnalisée reçoit automatiquement un `invocation` argument en tant que dernier argument.</span><span class="sxs-lookup"><span data-stu-id="d3da7-154">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="d3da7-155">Cet argument peut être utilisé pour récupérer un contexte supplémentaire, comme l’adresse de la cellule d’appel.</span><span class="sxs-lookup"><span data-stu-id="d3da7-155">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="d3da7-156">Ou elle peut être utilisée pour envoyer des informations à Excel, comme un gestionnaire de fonctions pour [annuler une fonction](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="d3da7-156">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="d3da7-157">Même si aucun paramètre n’est déclaré, votre fonction personnalisée a ce paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3da7-157">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="d3da7-158">Cet argument n’apparaît pas pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d3da7-158">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="d3da7-159">Si vous souhaitez utiliser `invocation` dans votre fonction personnalisée, déclarez-le comme dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3da7-159">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="d3da7-160">Dans l’exemple de code suivant, le `invocation` contexte est explicitement indiqué pour votre référence.</span><span class="sxs-lookup"><span data-stu-id="d3da7-160">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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
```

## <a name="next-steps"></a><span data-ttu-id="d3da7-161">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="d3da7-161">Next steps</span></span>

<span data-ttu-id="d3da7-162">Découvrez comment utiliser [des valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="d3da7-162">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d3da7-163">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d3da7-163">See also</span></span>

* [<span data-ttu-id="d3da7-164">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d3da7-164">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="d3da7-165">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d3da7-165">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="d3da7-166">Créer manuellement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d3da7-166">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d3da7-167">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="d3da7-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="d3da7-168">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="d3da7-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
