---
ms.date: 12/21/2020
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, et bien plus encore.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: 312046551236e96e67de6f63f3e3511aba6f50ce
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735528"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="1ce6e-103">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1ce6e-103">Custom functions parameter options</span></span>

<span data-ttu-id="1ce6e-104">Les fonctions personnalisées peuvent être configurées avec de nombreuses options de paramètres différentes.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="1ce6e-105">Paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="1ce6e-105">Optional parameters</span></span>

<span data-ttu-id="1ce6e-106">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="1ce6e-107">Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="1ce6e-108">Cette fonction apparaît sous la forme `=CONTOSO.ADD(first, second, [third])` dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="1ce6e-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="1ce6e-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="1ce6e-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="1ce6e-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="1ce6e-111">Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="1ce6e-112">Cela signifie que les paramètres initialisés par défaut dans la machine à écrire ne fonctionnent pas comme prévu.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="1ce6e-113">N’utilisez pas la syntaxe `function add(first:number, second:number, third=0):number` car elle ne peut pas `third` être initialisée à 0.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="1ce6e-114">À la place, utilisez la syntaxe de la machine à écrire comme indiqué dans l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="1ce6e-115">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se passe lorsque les paramètres facultatifs sont null.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="1ce6e-116">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="1ce6e-117">Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="1ce6e-118">Si le `dayOfWeek` paramètre est null, il est défini sur mercredi.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="1ce6e-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="1ce6e-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="1ce6e-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="1ce6e-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="1ce6e-121">Paramètres de plage</span><span class="sxs-lookup"><span data-stu-id="1ce6e-121">Range parameters</span></span>

<span data-ttu-id="1ce6e-122">Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="1ce6e-123">Une fonction peut également renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-123">A function can also return a range of data.</span></span> <span data-ttu-id="1ce6e-124">Excel passe une plage de données de cellule sous la forme d’un tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="1ce6e-125">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="1ce6e-126">La fonction suivante accepte le paramètre `values` , et la syntaxe JSDOC `number[][]` définit la propriété du paramètre `dimensionality` sur `matrix` dans les métadonnées JSON pour cette fonction.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="1ce6e-127">Paramètres répétitifs</span><span class="sxs-lookup"><span data-stu-id="1ce6e-127">Repeating parameters</span></span>

<span data-ttu-id="1ce6e-128">Un paramètre extensible permet à un utilisateur d’entrer une série d’arguments facultatifs dans une fonction.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="1ce6e-129">Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="1ce6e-130">Si le nom du paramètre se termine par un nombre, le numéro de chaque argument augmente de manière incrémentielle, par exemple `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="1ce6e-131">Cela correspond à la Convention utilisée pour les fonctions Excel intégrées.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="1ce6e-132">La fonction suivante additionne le total des nombres, des adresses de cellules, ainsi que des plages, si elles sont entrées.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="1ce6e-133">Cette fonction s’affiche `=CONTOSO.ADD([operands], [operands]...)` dans le classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="1ce6e-134">Paramètre extensible à valeur unique</span><span class="sxs-lookup"><span data-stu-id="1ce6e-134">Repeating single value parameter</span></span>

<span data-ttu-id="1ce6e-135">Un paramètre de valeur unique extensible permet de transmettre plusieurs valeurs uniques.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="1ce6e-136">Par exemple, l’utilisateur peut entrer ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="1ce6e-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="1ce6e-137">L’exemple suivant montre comment déclarer un seul paramètre de valeur.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="1ce6e-138">Paramètre de plage unique</span><span class="sxs-lookup"><span data-stu-id="1ce6e-138">Single range parameter</span></span>

<span data-ttu-id="1ce6e-139">Un paramètre de plage unique n’est pas techniquement un paramètre répétitif, mais il est inclus ici, car la déclaration est très similaire aux paramètres répétitifs.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="1ce6e-140">Il apparaîtrait à l’utilisateur sous la forme ADD (a2 : B3) où une seule plage est passée d’Excel.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="1ce6e-141">L’exemple suivant montre comment déclarer un paramètre de plage unique.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="1ce6e-142">Paramètre de plage extensible</span><span class="sxs-lookup"><span data-stu-id="1ce6e-142">Repeating range parameter</span></span>

<span data-ttu-id="1ce6e-143">Un paramètre de plage extensible permet de transmettre plusieurs plages ou nombres.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="1ce6e-144">Par exemple, l’utilisateur peut entrer ADD (5, B2, C3, 8, E5 : E8).</span><span class="sxs-lookup"><span data-stu-id="1ce6e-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="1ce6e-145">Les plages extensibles sont généralement spécifiées avec le type `number[][][]` comme il s’agit de matrices en trois dimensions.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="1ce6e-146">Pour un exemple, reportez-vous à l’exemple principal ci-dessous pour les paramètres de répétition (paramètres #repeating).</span><span class="sxs-lookup"><span data-stu-id="1ce6e-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="1ce6e-147">Déclaration de paramètres répétitifs</span><span class="sxs-lookup"><span data-stu-id="1ce6e-147">Declaring repeating parameters</span></span>
<span data-ttu-id="1ce6e-148">Dans la machine à écrire, indiquez que le paramètre est à plusieurs dimensions.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="1ce6e-149">Par exemple,  `ADD(values: number[])` un tableau à une dimension indiquerait `ADD(values:number[][])` un tableau à deux dimensions, et ainsi de suite.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="1ce6e-150">En JavaScript, utilisez `@param values {number[]}` pour les tableaux à une dimension, `@param <name> {number[][]}` pour les tableaux à deux dimensions, et ainsi de suite pour d’autres dimensions.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="1ce6e-151">Pour le format JSON dynamique, vérifiez que votre paramètre est spécifié en tant que `"repeating": true` dans votre fichier JSON, et vérifiez que vos paramètres sont marqués comme `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="1ce6e-152">Paramètre invocation</span><span class="sxs-lookup"><span data-stu-id="1ce6e-152">Invocation parameter</span></span>

<span data-ttu-id="1ce6e-153">Chaque fonction personnalisée reçoit automatiquement un `invocation` argument en tant que dernier paramètre d’entrée, même si elle n’est pas explicitement déclarée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="1ce6e-154">Ce `invocation` paramètre correspond à l’objet [invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="1ce6e-155">L' `Invocation` objet peut être utilisé pour récupérer un contexte supplémentaire, comme l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="1ce6e-156">Pour accéder à l' `Invocation` objet, vous devez déclarer `invocation` le dernier paramètre de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="1ce6e-157">Le `invocation` paramètre n’apparaît pas en tant qu’argument de fonction personnalisée pour les utilisateurs dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="1ce6e-158">L’exemple suivant montre comment utiliser le `invocation` paramètre pour renvoyer l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="1ce6e-159">Cet exemple utilise la propriété [Address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) de l' `Invocation` objet.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="1ce6e-160">Pour accéder à l' `Invocation` objet, déclarez tout d’abord `CustomFunctions.Invocation` en tant que paramètre dans votre JSDoc.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="1ce6e-161">Ensuite, déclarez `@requiresAddress` dans votre JSDoc pour accéder à la `address` propriété de l' `Invocation` objet.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="1ce6e-162">Enfin, dans la fonction, récupérez et renvoyez la `address` propriété.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<span data-ttu-id="1ce6e-163">Dans Excel, une fonction personnalisée qui appelle la `address` propriété de l' `Invocation` objet renvoie l’adresse absolue suivant le format `SheetName!RelativeCellAddress` dans la cellule qui a appelé la fonction.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="1ce6e-164">Par exemple, si le paramètre d’entrée se trouve sur une feuille appelée **prix** dans la cellule F6, la valeur de l’adresse du paramètre renvoyé sera `Prices!F6` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="1ce6e-165">Le `invocation` paramètre peut également être utilisé pour envoyer des informations à Excel.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="1ce6e-166">Pour en savoir plus, consultez [la rubrique créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function) .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="1ce6e-167">Détection de l’adresse d’un paramètre</span><span class="sxs-lookup"><span data-stu-id="1ce6e-167">Detect the address of a parameter</span></span>

<span data-ttu-id="1ce6e-168">En combinaison avec le [paramètre invocation](#invocation-parameter), vous pouvez utiliser l’objet [invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) pour récupérer l’adresse d’un paramètre d’entrée de fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="1ce6e-169">Lorsqu’elle est appelée, la propriété [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) de l' `Invocation` objet permet à une fonction de renvoyer les adresses de tous les paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="1ce6e-170">Cela est utile dans les scénarios où les types de données d’entrée peuvent varier.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="1ce6e-171">L’adresse d’un paramètre d’entrée peut être utilisée pour vérifier le format numérique de la valeur d’entrée.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="1ce6e-172">Le format de nombre peut ensuite être ajusté avant l’entrée, si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="1ce6e-173">L’adresse d’un paramètre d’entrée peut également être utilisée pour détecter si la valeur d’entrée possède des propriétés connexes susceptibles de concerner les calculs ultérieurs.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!IMPORTANT]
> <span data-ttu-id="1ce6e-174">La `parameterAddresses` propriété ne fonctionne actuellement qu’avec des [métadonnées JSON créées manuellement](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="1ce6e-174">The `parameterAddresses` property currently only works with [manually-created JSON metadata](custom-functions-json.md).</span></span> <span data-ttu-id="1ce6e-175">Pour renvoyer des adresses de paramètres, la `options` propriété de l’objet doit être `requiresParameterAddresses` définie sur `true` , et l' `result` objet doit avoir la `dimensionality` propriété définie sur `matrix` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-175">To return parameter addresses, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="1ce6e-176">La fonction personnalisée suivante accepte trois paramètres d’entrée, récupère la `parameterAddresses` propriété de l' `Invocation` objet pour chaque paramètre, puis renvoie ces adresses.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-176">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<span data-ttu-id="1ce6e-177">Lorsqu’une fonction personnalisée qui appelle la `parameterAddresses` propriété est exécutée, l’adresse du paramètre est renvoyée suivant le format `SheetName!RelativeCellAddress` dans la cellule qui a appelé la fonction.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-177">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="1ce6e-178">Par exemple, si le paramètre d’entrée se trouve sur une feuille appelée **coûts** dans la cellule D8, la valeur de l’adresse du paramètre renvoyé sera `Costs!D8` .</span><span class="sxs-lookup"><span data-stu-id="1ce6e-178">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="1ce6e-179">Si la fonction personnalisée possède plusieurs paramètres et que plusieurs adresses de paramètres sont renvoyées, les adresses renvoyées s’affichent dans plusieurs cellules, décroissant verticalement, à partir de la cellule qui a appelé la fonction.</span><span class="sxs-lookup"><span data-stu-id="1ce6e-179">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="1ce6e-180">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="1ce6e-180">Next steps</span></span>

<span data-ttu-id="1ce6e-181">Découvrez comment utiliser [des valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="1ce6e-181">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1ce6e-182">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1ce6e-182">See also</span></span>

* [<span data-ttu-id="1ce6e-183">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1ce6e-183">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="1ce6e-184">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1ce6e-184">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="1ce6e-185">Créer manuellement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1ce6e-185">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="1ce6e-186">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="1ce6e-186">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="1ce6e-187">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="1ce6e-187">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
