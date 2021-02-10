---
ms.date: 02/04/2021
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, tels que les plages Excel, les paramètres facultatifs, le contexte d’appel, etc.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: afe6947b1a1b9022a0284535b9ab1d68c9777c14
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173905"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="eae78-103">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="eae78-103">Custom functions parameter options</span></span>

<span data-ttu-id="eae78-104">Les fonctions personnalisées sont configurables avec de nombreuses options de paramètre différentes.</span><span class="sxs-lookup"><span data-stu-id="eae78-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="eae78-105">Paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="eae78-105">Optional parameters</span></span>

<span data-ttu-id="eae78-106">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="eae78-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="eae78-107">Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre.</span><span class="sxs-lookup"><span data-stu-id="eae78-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="eae78-108">Cette fonction apparaît comme `=CONTOSO.ADD(first, second, [third])` dans Excel.</span><span class="sxs-lookup"><span data-stu-id="eae78-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="eae78-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="eae78-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="eae78-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="eae78-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="eae78-111">Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null` .</span><span class="sxs-lookup"><span data-stu-id="eae78-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="eae78-112">Cela signifie que les paramètres initialisés par défaut dans TypeScript ne fonctionneront pas comme prévu.</span><span class="sxs-lookup"><span data-stu-id="eae78-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="eae78-113">N’utilisez pas la `function add(first:number, second:number, third=0):number` syntaxe, car elle ne s’initialisera pas sur `third` 0.</span><span class="sxs-lookup"><span data-stu-id="eae78-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="eae78-114">Utilisez plutôt la syntaxe TypeScript comme illustré dans l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="eae78-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="eae78-115">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se produit lorsque les paramètres facultatifs sont null.</span><span class="sxs-lookup"><span data-stu-id="eae78-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="eae78-116">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="eae78-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="eae78-117">Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052` .</span><span class="sxs-lookup"><span data-stu-id="eae78-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="eae78-118">Si le `dayOfWeek` paramètre est null, il est paramétrable mercredi.</span><span class="sxs-lookup"><span data-stu-id="eae78-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="eae78-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="eae78-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="eae78-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="eae78-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="eae78-121">Paramètres de plage</span><span class="sxs-lookup"><span data-stu-id="eae78-121">Range parameters</span></span>

<span data-ttu-id="eae78-122">Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée.</span><span class="sxs-lookup"><span data-stu-id="eae78-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="eae78-123">Une fonction peut également renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="eae78-123">A function can also return a range of data.</span></span> <span data-ttu-id="eae78-124">Excel passe une plage de données de cellule sous forme de tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="eae78-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="eae78-125">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="eae78-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="eae78-126">La fonction suivante accepte le paramètre et la syntaxe JSDOC définit la propriété du paramètre dans les métadonnées `values` `number[][]` `dimensionality` `matrix` JSON pour cette fonction.</span><span class="sxs-lookup"><span data-stu-id="eae78-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="eae78-127">Paramètres répétés</span><span class="sxs-lookup"><span data-stu-id="eae78-127">Repeating parameters</span></span>

<span data-ttu-id="eae78-128">Un paramètre exercissable permet à un utilisateur d’entrer une série d’arguments facultatifs à une fonction.</span><span class="sxs-lookup"><span data-stu-id="eae78-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="eae78-129">Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre.</span><span class="sxs-lookup"><span data-stu-id="eae78-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="eae78-130">Si le nom du paramètre se termine par un nombre, le nombre de chaque argument augmente de manière incrémentielle, par `ADD(number1, [number2], [number3],…)` exemple.</span><span class="sxs-lookup"><span data-stu-id="eae78-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="eae78-131">Cela correspond à la convention utilisée pour les fonctions Excel intégrées.</span><span class="sxs-lookup"><span data-stu-id="eae78-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="eae78-132">La fonction suivante additione le total des nombres, des adresses de cellule, ainsi que des plages, si elles sont entrées.</span><span class="sxs-lookup"><span data-stu-id="eae78-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="eae78-133">Cette fonction `=CONTOSO.ADD([operands], [operands]...)` s’affiche dans le livre de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="eae78-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="eae78-134">Paramètre de valeur unique répété</span><span class="sxs-lookup"><span data-stu-id="eae78-134">Repeating single value parameter</span></span>

<span data-ttu-id="eae78-135">Un paramètre à valeur unique exercissable permet de passer plusieurs valeurs simples.</span><span class="sxs-lookup"><span data-stu-id="eae78-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="eae78-136">Par exemple, l’utilisateur peut entrer ADD(1,B2,3).</span><span class="sxs-lookup"><span data-stu-id="eae78-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="eae78-137">L’exemple suivant montre comment déclarer un paramètre de valeur unique.</span><span class="sxs-lookup"><span data-stu-id="eae78-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="eae78-138">Paramètre de plage unique</span><span class="sxs-lookup"><span data-stu-id="eae78-138">Single range parameter</span></span>

<span data-ttu-id="eae78-139">Un paramètre de plage unique n’est techniquement pas un paramètre exercissable, mais il est inclus ici, car la déclaration est très similaire aux paramètres ext ments ex r us.</span><span class="sxs-lookup"><span data-stu-id="eae78-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="eae78-140">Il apparaît à l’utilisateur comme ADD(A2:B3) où une seule plage est passée à partir d’Excel.</span><span class="sxs-lookup"><span data-stu-id="eae78-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="eae78-141">L’exemple suivant montre comment déclarer un paramètre de plage unique.</span><span class="sxs-lookup"><span data-stu-id="eae78-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="eae78-142">Paramètre de plage répétée</span><span class="sxs-lookup"><span data-stu-id="eae78-142">Repeating range parameter</span></span>

<span data-ttu-id="eae78-143">Un paramètre de plage exercidable permet de passer plusieurs plages ou nombres.</span><span class="sxs-lookup"><span data-stu-id="eae78-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="eae78-144">Par exemple, l’utilisateur peut entrer ADD(5,B2,C3,8,E5:E8).</span><span class="sxs-lookup"><span data-stu-id="eae78-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="eae78-145">Les plages exercidées sont généralement spécifiées avec le type, car il s’agit de `number[][][]` matrices en trois dimensions.</span><span class="sxs-lookup"><span data-stu-id="eae78-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="eae78-146">Pour obtenir un exemple, consultez le principal exemple répertorié pour les paramètres répétés(#repeating-parameters).</span><span class="sxs-lookup"><span data-stu-id="eae78-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="eae78-147">Déclaration de paramètres répétés</span><span class="sxs-lookup"><span data-stu-id="eae78-147">Declaring repeating parameters</span></span>
<span data-ttu-id="eae78-148">Dans Typescript, indiquez que le paramètre est multidimensionnel.</span><span class="sxs-lookup"><span data-stu-id="eae78-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="eae78-149">Par exemple, cela indiquerait un tableau à une dimension, un tableau à  `ADD(values: number[])` `ADD(values:number[][])` deux dimensions, etc.</span><span class="sxs-lookup"><span data-stu-id="eae78-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="eae78-150">Dans JavaScript, utilisez pour les tableaux à une dimension, pour les tableaux à deux dimensions, et ainsi de `@param values {number[]}` suite pour plus de `@param <name> {number[][]}` dimensions.</span><span class="sxs-lookup"><span data-stu-id="eae78-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="eae78-151">Pour JSON écrit à la main, assurez-vous que votre paramètre est spécifié comme dans votre fichier JSON, et vérifiez que vos paramètres sont `"repeating": true` marqués comme `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="eae78-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="eae78-152">Paramètre d’appel</span><span class="sxs-lookup"><span data-stu-id="eae78-152">Invocation parameter</span></span>

<span data-ttu-id="eae78-153">Chaque fonction personnalisée est automatiquement passée un argument comme dernier paramètre `invocation` d’entrée, même s’il n’est pas explicitement déclaré.</span><span class="sxs-lookup"><span data-stu-id="eae78-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="eae78-154">Ce `invocation` paramètre correspond à l’objet [Invocation.](/javascript/api/custom-functions-runtime/customfunctions.invocation)</span><span class="sxs-lookup"><span data-stu-id="eae78-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="eae78-155">L’objet peut être utilisé pour récupérer un contexte supplémentaire, tel que l’adresse de la cellule `Invocation` qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="eae78-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="eae78-156">Pour accéder à `Invocation` l’objet, vous devez déclarer `invocation` comme dernier paramètre de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="eae78-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="eae78-157">Le `invocation` paramètre n’apparaît pas en tant qu’argument de fonction personnalisée pour les utilisateurs dans Excel.</span><span class="sxs-lookup"><span data-stu-id="eae78-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="eae78-158">L’exemple suivant montre comment utiliser le paramètre pour renvoyer l’adresse de la cellule `invocation` qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="eae78-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="eae78-159">Cet exemple utilise la propriété [d’adresse](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) de `Invocation` l’objet.</span><span class="sxs-lookup"><span data-stu-id="eae78-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="eae78-160">Pour accéder à `Invocation` l’objet, déclarez d’abord `CustomFunctions.Invocation` en tant que paramètre dans votre JSDoc.</span><span class="sxs-lookup"><span data-stu-id="eae78-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="eae78-161">Ensuite, déclarez `@requiresAddress` dans votre JSDoc pour accéder à la `address` propriété de `Invocation` l’objet.</span><span class="sxs-lookup"><span data-stu-id="eae78-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="eae78-162">Enfin, dans la fonction, récupérez et renvoyez la `address` propriété.</span><span class="sxs-lookup"><span data-stu-id="eae78-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

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

<span data-ttu-id="eae78-163">Dans Excel, une fonction personnalisée appelant la propriété de l’objet retourne l’adresse absolue en suivant le format de la cellule qui a `address` `Invocation` appelé la `SheetName!RelativeCellAddress` fonction.</span><span class="sxs-lookup"><span data-stu-id="eae78-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="eae78-164">Par exemple, si le paramètre d’entrée se trouve dans une feuille appelée **Prix** dans la cellule F6, la valeur d’adresse du paramètre renvoyé est `Prices!F6` .</span><span class="sxs-lookup"><span data-stu-id="eae78-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="eae78-165">Le `invocation` paramètre peut également être utilisé pour envoyer des informations à Excel.</span><span class="sxs-lookup"><span data-stu-id="eae78-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="eae78-166">Pour en [savoir plus, voir](custom-functions-web-reqs.md#make-a-streaming-function) Faire une fonction de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="eae78-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="eae78-167">Détecter l’adresse d’un paramètre</span><span class="sxs-lookup"><span data-stu-id="eae78-167">Detect the address of a parameter</span></span>

<span data-ttu-id="eae78-168">En combinaison avec le paramètre [d’appel,](#invocation-parameter)vous pouvez utiliser l’objet [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) pour récupérer l’adresse d’un paramètre d’entrée de fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="eae78-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="eae78-169">Lorsqu’elle est invoquée, [la propriété parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) de l’objet permet à une fonction de renvoyer les adresses de tous `Invocation` les paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="eae78-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="eae78-170">Cela est utile dans les scénarios où les types de données d’entrée peuvent varier.</span><span class="sxs-lookup"><span data-stu-id="eae78-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="eae78-171">L’adresse d’un paramètre d’entrée peut être utilisée pour vérifier le format numérique de la valeur d’entrée.</span><span class="sxs-lookup"><span data-stu-id="eae78-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="eae78-172">Le format numérique peut ensuite être ajusté avant l’entrée, si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="eae78-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="eae78-173">L’adresse d’un paramètre d’entrée peut également être utilisée pour détecter si la valeur d’entrée possède des propriétés associées qui peuvent être pertinentes pour les calculs suivants.</span><span class="sxs-lookup"><span data-stu-id="eae78-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!NOTE]
> <span data-ttu-id="eae78-174">Si vous travaillez avec des métadonnées [JSON](custom-functions-json.md) créées manuellement pour renvoyer des adresses de paramètre au lieu du générateur Yo Office, l’objet doit avoir la propriété définie sur , et l’objet doit avoir la propriété définie sur `options` `requiresParameterAddresses` `true` `result` `dimensionality` `matrix` .</span><span class="sxs-lookup"><span data-stu-id="eae78-174">If you're working with [manually-created JSON metadata](custom-functions-json.md) to return parameter addresses instead of the Yo Office generator, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="eae78-175">La fonction personnalisée suivante prend trois paramètres d’entrée, récupère la propriété de l’objet pour chaque paramètre, puis `parameterAddresses` `Invocation` renvoie les adresses.</span><span class="sxs-lookup"><span data-stu-id="eae78-175">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

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

<span data-ttu-id="eae78-176">Lorsqu’une fonction personnalisée appelant la propriété s’exécute, l’adresse du paramètre est renvoyée en suivant le format de la cellule `parameterAddresses` qui a appelé la `SheetName!RelativeCellAddress` fonction.</span><span class="sxs-lookup"><span data-stu-id="eae78-176">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="eae78-177">Par exemple, si le paramètre d’entrée se trouve dans une feuille appelée **Costs** dans la cellule D8, la valeur d’adresse du paramètre renvoyé est `Costs!D8` .</span><span class="sxs-lookup"><span data-stu-id="eae78-177">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="eae78-178">Si la fonction personnalisée possède plusieurs paramètres et que plusieurs adresses de paramètre sont renvoyées, les adresses renvoyées se renverront sur plusieurs cellules, décroit verticalement à partir de la cellule qui a appelé la fonction.</span><span class="sxs-lookup"><span data-stu-id="eae78-178">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="eae78-179">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="eae78-179">Next steps</span></span>

<span data-ttu-id="eae78-180">Découvrez comment utiliser des [valeurs volatiles dans vos fonctions personnalisées.](custom-functions-volatile.md)</span><span class="sxs-lookup"><span data-stu-id="eae78-180">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="eae78-181">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eae78-181">See also</span></span>

* [<span data-ttu-id="eae78-182">Recevoir et gérer des données à l’aide de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="eae78-182">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="eae78-183">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="eae78-183">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="eae78-184">Créer manuellement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="eae78-184">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="eae78-185">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="eae78-185">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="eae78-186">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="eae78-186">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
