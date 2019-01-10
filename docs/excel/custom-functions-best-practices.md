---
ms.date: 01/08/2019
description: Découvrez les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.
title: Meilleures pratiques de fonctions personnalisées (aperçu)
ms.openlocfilehash: 45618a61d0d1fdd0398ecec3aa0db21e493787fd
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770650"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="71b01-103">Meilleures pratiques de fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="71b01-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="71b01-104">Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="71b01-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="71b01-105">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="71b01-105">Error handling</span></span>

<span data-ttu-id="71b01-106">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="71b01-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="71b01-107">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="71b01-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="71b01-108">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="71b01-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="troubleshooting"></a><span data-ttu-id="71b01-109">Résolution des problèmes</span><span class="sxs-lookup"><span data-stu-id="71b01-109">Troubleshooting</span></span>

<span data-ttu-id="71b01-110">Si vous testez votre complément dans Office sur Windows, vous devez autoriser la \*\* [connexion d’exécution](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) \*\* à résoudre les problèmes XML du fichier manifeste de votre complément, ainsi que plusieurs conditions d’installation et exécution.</span><span class="sxs-lookup"><span data-stu-id="71b01-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="71b01-111">La connexion d’exécution écrit les`console.log`instructions vers un fichier journal pour vous aider à découvrir des problèmes.</span><span class="sxs-lookup"><span data-stu-id="71b01-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="71b01-112">Pour signaler des commentaires à l’équipe Excel des fonctions personnalisées sur cette méthode de résolution des problèmes, envoyez des commentaires à l’équipe.</span><span class="sxs-lookup"><span data-stu-id="71b01-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="71b01-113">Pour ce faire, sélectionnez **Fichier | Commentaires | Envoyer un smiley mécontent**.</span><span class="sxs-lookup"><span data-stu-id="71b01-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="71b01-114">Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.</span><span class="sxs-lookup"><span data-stu-id="71b01-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="debugging"></a><span data-ttu-id="71b01-115">Débogage</span><span class="sxs-lookup"><span data-stu-id="71b01-115">Debugging</span></span>

<span data-ttu-id="71b01-116">Pour l’instant, la méthode optimale pour le débogage de fonctions personnalisées Excel consiste à [charger](../testing/sideload-office-add-ins-for-testing.md) votre complément au sein d’**Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="71b01-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="71b01-117">Vous pouvez ensuite déboguer vos fonctions personnalisées à l’aide de l’ [outil natif F12 de débogage de votre navigateur](../testing/debug-add-ins-in-office-online.md) en combinaison avec les techniques suivantes :</span><span class="sxs-lookup"><span data-stu-id="71b01-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="71b01-118">Utilisez les`console.log` instructions au sein de votre code de fonctions personnalisées pour envoyer la sortie à la console en temps réel.</span><span class="sxs-lookup"><span data-stu-id="71b01-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="71b01-119">Utilisez les `debugger;` instructions au sein de votre code de fonctions personnalisées pour spécifier les points d'arrêt où l’exécution sera suspendue lorsque la fenêtre F12 est ouverte.</span><span class="sxs-lookup"><span data-stu-id="71b01-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="71b01-120">Par exemple, si la fonction suivante s’exécute lorsque la fenêtre F12 est ouverte, l’exécution sera suspendue sur la`debugger;` déclaration, vous permettant d’inspecter manuellement les valeurs de paramètres avant le retour de la fonction.</span><span class="sxs-lookup"><span data-stu-id="71b01-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="71b01-121">L’`debugger;` instruction n’a aucun effet dans Excel Online lorsque la fenêtre F12 n’est pas ouverte.</span><span class="sxs-lookup"><span data-stu-id="71b01-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="71b01-122">Pour l’instant, l’`debugger;` instruction n’a aucun effet dans Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="71b01-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="71b01-123">Si votre complément ne parvient pas à s’enregistrer, [vérifier que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web hébergeant votre application complément.</span><span class="sxs-lookup"><span data-stu-id="71b01-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="71b01-124">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="71b01-124">Associating function names with JSON metadata</span></span>

<span data-ttu-id="71b01-125">Comme décrit dans l’article[vue d’ensemble de fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier de script (JavaScript ou machine à écrire) pour former une fonction complète.</span><span class="sxs-lookup"><span data-stu-id="71b01-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="71b01-126">Pour qu’une fonction s’exécute correctement, vous devez lier le nom de la fonction dans le fichier de script à l’id répertorié dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="71b01-126">For a function to work properly, you'll need to bind the name of the function in the script file to the id listed in the JSON file.</span></span> <span data-ttu-id="71b01-127">Ce processus est appelé association.</span><span class="sxs-lookup"><span data-stu-id="71b01-127">This process is called association.</span></span> <span data-ttu-id="71b01-128">Pensez à inclure les associations à la fin de vos fichiers de code JavaScript ; dans le cas contraire, les fonctions ne fonctionneront pas.</span><span class="sxs-lookup"><span data-stu-id="71b01-128">Make a note to include associations at the end of your JavaScript code files; otherwise, your functions will not work.</span></span>

<span data-ttu-id="71b01-129">L’exemple de code suivant montre comment procéder à cette association.</span><span class="sxs-lookup"><span data-stu-id="71b01-129">The following multi-part POST code shows how to do this.</span></span> <span data-ttu-id="71b01-130">L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="71b01-130">For example, the following code sample defines the custom function  and then specifies that the function  corresponds to the object in the JSON metadata file where the value of the  property is ADD.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add); 
```

<span data-ttu-id="71b01-131">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="71b01-131">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="71b01-132">Utilisez uniquement des lettres majuscules d’une fonction `name` et `id` dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="71b01-132">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="71b01-133">N’utilisez pas un mélange de cas ou uniquement des lettres minuscules.</span><span class="sxs-lookup"><span data-stu-id="71b01-133">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="71b01-134">Si vous le faites, vous risquez de finir avec deux valeurs différentes uniquement par la casse ,cela entraînera un remplacement involontaire de vos fonctions.</span><span class="sxs-lookup"><span data-stu-id="71b01-134">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="71b01-135">Par exemple, un objet de fonction à une `id` valeur**ajouter** peut être remplacé par déclaration plus loin dans le fichier d’objet de fonction avec une`id` valeur**AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="71b01-135">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="71b01-136">De plus, la `name` propriété définit le nom de la fonction que les utilisateurs finaux verront dans Excel.</span><span class="sxs-lookup"><span data-stu-id="71b01-136">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="71b01-137">Utiliser des lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonction intégrée sont en majuscules.</span><span class="sxs-lookup"><span data-stu-id="71b01-137">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="71b01-138">Toutefois, il n’est pas nécessaire de tirer profit de la fonction `name` lors de l’association.</span><span class="sxs-lookup"><span data-stu-id="71b01-138">However, it is not necessary to capitalize the function's `name` when associating.</span></span> <span data-ttu-id="71b01-139">Dans l’exemple,`CustomFunctions.associate("add", add)`équivaut à`CustomFunctions.associate("ADD", add)`.</span><span class="sxs-lookup"><span data-stu-id="71b01-139">For example, `CustomFunctions.associate("add", add)` is equivalent to `CustomFunctions.associate("ADD", add)`.</span></span>

* <span data-ttu-id="71b01-140">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="71b01-140">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="71b01-141">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="71b01-141">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="71b01-142">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.</span><span class="sxs-lookup"><span data-stu-id="71b01-142">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="71b01-143">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="71b01-143">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="71b01-144">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="71b01-144">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="71b01-145">Dans le fichier JavaScript, spécifiez tous les mappages de fonctions personnalisées dans le même emplacement.</span><span class="sxs-lookup"><span data-stu-id="71b01-145">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="71b01-146">Par exemple, le code suivant définit deux fonctions personnalisées et indique ensuite les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="71b01-146">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    <span data-ttu-id="71b01-147">L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="71b01-147">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="71b01-148">Notez que les propriétés`id` et `name`sont en majuscules dans ce fichier.</span><span class="sxs-lookup"><span data-stu-id="71b01-148">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="71b01-149">Déclarer des paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="71b01-149">Declaring optional parameters</span></span> 
<span data-ttu-id="71b01-150">Dans Excel pour Windows (version 1812 ou version ultérieure), vous pouvez déclarer des paramètres facultatifs pour vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="71b01-150">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="71b01-151">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="71b01-151">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="71b01-152">Par exemple, une fonction `FOO` avec un paramètre obligatoire appelé`parameter1` et un autre paramètre facultatif appelé `parameter2` apparaîtra sous la forme `=FOO(parameter1, [parameter2])` dans Excel.</span><span class="sxs-lookup"><span data-stu-id="71b01-152">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="71b01-153">Pour rendre un paramètre facultatif, ajouter `"optional": true` au paramètre dans le fichier de métadonnées JSON qui définit la fonction.</span><span class="sxs-lookup"><span data-stu-id="71b01-153">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="71b01-154">L’exemple suivant montre comment cela peut se présenter pour la fonction `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="71b01-154">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="71b01-155">Vous pouvez remarquer que le paramètre facultatif `[third]` suit deux paramètres requis.</span><span class="sxs-lookup"><span data-stu-id="71b01-155">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="71b01-156">Les paramètres obligatoires apparaissent en premier dans l’interface utilisateur formule d’Excel.</span><span class="sxs-lookup"><span data-stu-id="71b01-156">Required parameters will appear first in Excel’s Formula UI.</span></span>

```json
{
    "id": "ADD",
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

<span data-ttu-id="71b01-157">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qu’il se passe lorsque les paramètres facultatifs ne sont pas définis.</span><span class="sxs-lookup"><span data-stu-id="71b01-157">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="71b01-158">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="71b01-158">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="71b01-159">Si le paramètre`zipCode` n’est pas défini, la valeur par défaut est définie sur 98052.</span><span class="sxs-lookup"><span data-stu-id="71b01-159">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="71b01-160">Si le paramètre`dayOfWeek` n’est pas défini, la valeur par défaut est définie à mercredi.</span><span class="sxs-lookup"><span data-stu-id="71b01-160">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="71b01-161">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="71b01-161">Additional considerations</span></span>

<span data-ttu-id="71b01-162">Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM.</span><span class="sxs-lookup"><span data-stu-id="71b01-162">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="71b01-163">Sur Excel pour Windows, où les fonctions personnalisées utilisent l’[exécution JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="71b01-163">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="71b01-164">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="71b01-164">See also</span></span>

* [<span data-ttu-id="71b01-165">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="71b01-165">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="71b01-166">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="71b01-166">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="71b01-167">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="71b01-167">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="71b01-168">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="71b01-168">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="71b01-169">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="71b01-169">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
