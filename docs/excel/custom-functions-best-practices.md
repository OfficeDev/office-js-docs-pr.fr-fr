---
ms.date: 11/29/2018
description: Découvrez les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.
title: Meilleures pratiques des fonctions personnalisées
ms.openlocfilehash: c1be1d01a88d50bb0f3aee8af1aea7c47658bc10
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724885"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="80eba-103">Meilleures pratiques de fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="80eba-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="80eba-104">Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="80eba-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="80eba-105">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="80eba-105">Error handling</span></span>

<span data-ttu-id="80eba-106">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="80eba-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="80eba-107">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="80eba-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="80eba-108">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="80eba-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="troubleshooting"></a><span data-ttu-id="80eba-109">Résolution des problèmes</span><span class="sxs-lookup"><span data-stu-id="80eba-109">Troubleshooting</span></span>

<span data-ttu-id="80eba-110">Si vous testez votre complément dans Office sur Windows, vous devez autoriser la \*\* [connexion d’exécution](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) \*\* à résoudre les problèmes XML du fichier manifeste de votre complément, ainsi que plusieurs conditions d’installation et exécution.</span><span class="sxs-lookup"><span data-stu-id="80eba-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="80eba-111">La connexion d’exécution écrit les`console.log`instructions vers un fichier journal pour vous aider à découvrir des problèmes.</span><span class="sxs-lookup"><span data-stu-id="80eba-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="80eba-112">Pour signaler des commentaires à l’équipe Excel des fonctions personnalisées sur cette méthode de résolution des problèmes, envoyez des commentaires à l’équipe.</span><span class="sxs-lookup"><span data-stu-id="80eba-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="80eba-113">Pour ce faire, sélectionnez **Fichier | Commentaires | Envoyer un smiley mécontent**.</span><span class="sxs-lookup"><span data-stu-id="80eba-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="80eba-114">Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.</span><span class="sxs-lookup"><span data-stu-id="80eba-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="debugging"></a><span data-ttu-id="80eba-115">Débogage</span><span class="sxs-lookup"><span data-stu-id="80eba-115">Debugging</span></span>

<span data-ttu-id="80eba-116">Pour l’instant, la méthode optimale pour le débogage de fonctions personnalisées Excel consiste à [charger](../testing/sideload-office-add-ins-for-testing.md) votre complément au sein d’**Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="80eba-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="80eba-117">Vous pouvez ensuite déboguer vos fonctions personnalisées à l’aide de l’ [outil natif F12 de débogage de votre navigateur](../testing/debug-add-ins-in-office-online.md) en combinaison avec les techniques suivantes :</span><span class="sxs-lookup"><span data-stu-id="80eba-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="80eba-118">Utilisez les`console.log` instructions au sein de votre code de fonctions personnalisées pour envoyer la sortie à la console en temps réel.</span><span class="sxs-lookup"><span data-stu-id="80eba-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="80eba-119">Utilisez les `debugger;` instructions au sein de votre code de fonctions personnalisées pour spécifier les points d'arrêt où l’exécution sera suspendue lorsque la fenêtre F12 est ouverte.</span><span class="sxs-lookup"><span data-stu-id="80eba-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="80eba-120">Par exemple, si la fonction suivante s’exécute lorsque la fenêtre F12 est ouverte, l’exécution sera suspendue sur la`debugger;` déclaration, vous permettant d’inspecter manuellement les valeurs de paramètres avant le retour de la fonction.</span><span class="sxs-lookup"><span data-stu-id="80eba-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="80eba-121">L’`debugger;` instruction n’a aucun effet dans Excel Online lorsque la fenêtre F12 n’est pas ouverte.</span><span class="sxs-lookup"><span data-stu-id="80eba-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="80eba-122">Pour l’instant, l’`debugger;` instruction n’a aucun effet dans Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="80eba-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="80eba-123">Si votre complément ne parvient pas à s’enregistrer, [vérifier que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web hébergeant votre application complément.</span><span class="sxs-lookup"><span data-stu-id="80eba-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="80eba-124">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="80eba-124">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="80eba-125">Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="80eba-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="80eba-126">Par ailleurs, dans le fichier JavaScript qui définit vos fonctions personnalisées, vous devez fournir des informations pour spécifier quel objet fonction dans le fichier de métadonnées JSON correspond à chaque fonction personnalisée dans le fichier JavaScript.</span><span class="sxs-lookup"><span data-stu-id="80eba-126">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="80eba-127">Par exemple, l’exemple de code suivant définit la fonction personnalisée `add` et puis indique que la fonction `add` correspond à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="80eba-127">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="80eba-128">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="80eba-128">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="80eba-129">Dans le fichier JavaScript, spécifiez les noms de fonction dans camelCase.</span><span class="sxs-lookup"><span data-stu-id="80eba-129">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="80eba-130">Par exemple, le nom de fonction `addTenToInput` écrit dans camelCase : le premier mot dans le nom commence par une lettre en minuscule et chaque mot suivant dans le nom commence par une lettre en majuscule.</span><span class="sxs-lookup"><span data-stu-id="80eba-130">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="80eba-131">Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque `name` propriété en majuscules.</span><span class="sxs-lookup"><span data-stu-id="80eba-131">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="80eba-132">La `name` propriété définit le nom de la fonction que les utilisateurs finaux verront dans Excel.</span><span class="sxs-lookup"><span data-stu-id="80eba-132">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="80eba-133">Utiliser des lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonction intégrée sont en majuscules.</span><span class="sxs-lookup"><span data-stu-id="80eba-133">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="80eba-134">Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque `id` propriété en majuscules.</span><span class="sxs-lookup"><span data-stu-id="80eba-134">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="80eba-135">Cette opération souligne quelle partie de l’`CustomFunctionMappings` instruction dans votre code JavaScript correspond à la `id` propriété dans le fichier métadonnées JSON (à condition que votre nom de fonction utilise camelCase, comme recommandé précédemment).</span><span class="sxs-lookup"><span data-stu-id="80eba-135">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="80eba-136">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="80eba-136">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="80eba-137">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="80eba-137">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="80eba-138">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit avoir la même`id` valeur.</span><span class="sxs-lookup"><span data-stu-id="80eba-138">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="80eba-139">En outre, n’indiquez pas deux `id` valeurs dans le fichier de métadonnées qui diffèrent uniquement par la casse.</span><span class="sxs-lookup"><span data-stu-id="80eba-139">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="80eba-140">Par exemple, ne définissez pas un objet fonction avec une `id` valeur **ajouter** et un autre objet fonction avec une`id` valeur de **AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="80eba-140">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="80eba-141">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="80eba-141">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="80eba-142">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="80eba-142">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="80eba-143">Dans le fichier JavaScript, spécifiez tous les mappages de fonctions personnalisées dans le même emplacement.</span><span class="sxs-lookup"><span data-stu-id="80eba-143">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="80eba-144">Par exemple, le code suivant définit deux fonctions personnalisées et indique ensuite les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="80eba-144">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    <span data-ttu-id="80eba-145">L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="80eba-145">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

    ```json
    {
      "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
      "functions": [
        {
          "id": "ADD",
          "name": "ADD",
          ...
        },
        {
          "id": "INCREMENT",
          "name": "INCREMENT",
          ...
        }
      ]
    }
    ```

## <a name="declaring-optional-parameters"></a><span data-ttu-id="80eba-146">Déclarer des paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="80eba-146">Declaring optional parameters</span></span> 
<span data-ttu-id="80eba-147">Dans Excel pour Windows (version 1812 ou version ultérieure), vous pouvez déclarer des paramètres facultatifs pour vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="80eba-147">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="80eba-148">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="80eba-148">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="80eba-149">Par exemple, une fonction `FOO` avec un paramètre obligatoire appelé`parameter1` et un autre paramètre facultatif appelé `parameter2` apparaîtra sous la forme `=FOO(parameter1, [parameter2])` dans Excel.</span><span class="sxs-lookup"><span data-stu-id="80eba-149">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="80eba-150">Pour rendre un paramètre facultatif, ajouter `"optional": true` au paramètre dans le fichier de métadonnées JSON qui définit la fonction.</span><span class="sxs-lookup"><span data-stu-id="80eba-150">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="80eba-151">L’exemple suivant montre comment cela peut se présenter pour la fonction `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="80eba-151">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="80eba-152">Vous pouvez remarquer que le paramètre facultatif `[third]` suit deux paramètres requis.</span><span class="sxs-lookup"><span data-stu-id="80eba-152">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="80eba-153">Les paramètres obligatoires apparaissent en premier dans l’interface utilisateur formule d’Excel.</span><span class="sxs-lookup"><span data-stu-id="80eba-153">Required parameters will appear first in Excel’s Formula UI.</span></span>

```json
{
    "id": "add",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

<span data-ttu-id="80eba-154">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qu’il se passe lorsque les paramètres facultatifs ne sont pas définis.</span><span class="sxs-lookup"><span data-stu-id="80eba-154">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="80eba-155">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="80eba-155">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="80eba-156">Si le paramètre`zipCode` n’est pas défini, la valeur par défaut est définie sur 98052.</span><span class="sxs-lookup"><span data-stu-id="80eba-156">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="80eba-157">Si le paramètre`dayOfWeek` n’est pas défini, la valeur par défaut est définie à mercredi.</span><span class="sxs-lookup"><span data-stu-id="80eba-157">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a><span data-ttu-id="80eba-158">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="80eba-158">Additional considerations</span></span>

<span data-ttu-id="80eba-159">Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM.</span><span class="sxs-lookup"><span data-stu-id="80eba-159">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="80eba-160">Sur Excel pour Windows, où les fonctions personnalisées utilisent l’[exécution JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="80eba-160">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="80eba-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="80eba-161">See also</span></span>

* [<span data-ttu-id="80eba-162">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="80eba-162">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="80eba-163">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="80eba-163">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="80eba-164">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="80eba-164">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="80eba-165">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="80eba-165">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
