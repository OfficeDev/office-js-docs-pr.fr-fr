---
ms.date: 01/08/2019
description: Découvrez les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.
title: Meilleures pratiques de fonctions personnalisées (aperçu)
localization_priority: Normal
ms.openlocfilehash: 4efcd0ba5efb0dc7450192694e8f0750de43b8a8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448608"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="ae659-103">Meilleures pratiques de fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="ae659-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="ae659-104">Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ae659-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a><span data-ttu-id="ae659-105">Résolution des problèmes</span><span class="sxs-lookup"><span data-stu-id="ae659-105">Troubleshooting</span></span>

1. <span data-ttu-id="ae659-106">Si vous testez votre complément dans Office sur Windows, vous devez autoriser la \*\* [connexion d’exécution](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) \*\* à résoudre les problèmes XML du fichier manifeste de votre complément, ainsi que plusieurs conditions d’installation et exécution.</span><span class="sxs-lookup"><span data-stu-id="ae659-106">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="ae659-107">La connexion d’exécution écrit les`console.log`instructions vers un fichier journal pour vous aider à découvrir des problèmes.</span><span class="sxs-lookup"><span data-stu-id="ae659-107">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

2. <span data-ttu-id="ae659-108">Votre complément ne se charge pas si une ou plusieurs fonctions personnalisées sont en conflit avec les fonctions personnalisées d'un complément enregistré précédemment.</span><span class="sxs-lookup"><span data-stu-id="ae659-108">Your add-in will not load if one or more custom functions conflicts with a previously registered add-in's custom functions.</span></span> <span data-ttu-id="ae659-109">Dans ce cas, vous pouvez supprimer le complément existant ou, si vous rencontrez cette erreur lors du développement d'un complément, vous pouvez spécifier un autre nom d'espace de noms dans votre manifeste.</span><span class="sxs-lookup"><span data-stu-id="ae659-109">In this case, you can either remove the existing add-in, or if you encounter this error while developing an add-in, you can specify a different namespace name in your manifest.</span></span>

3. <span data-ttu-id="ae659-110">Pour signaler des commentaires à l’équipe Excel des fonctions personnalisées sur cette méthode de résolution des problèmes, envoyez des commentaires à l’équipe.</span><span class="sxs-lookup"><span data-stu-id="ae659-110">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="ae659-111">Pour ce faire, sélectionnez **Fichier | Commentaires | Envoyer un smiley mécontent**.</span><span class="sxs-lookup"><span data-stu-id="ae659-111">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="ae659-112">Envoyer un smiley mécontent fournira les journaux nécessaires pour comprendre le problème que vous rencontrez.</span><span class="sxs-lookup"><span data-stu-id="ae659-112">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="ae659-113">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="ae659-113">Associating function names with JSON metadata</span></span>

<span data-ttu-id="ae659-114">Comme décrit dans l’article[vue d’ensemble de fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier de script (JavaScript ou machine à écrire) pour former une fonction complète.</span><span class="sxs-lookup"><span data-stu-id="ae659-114">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="ae659-115">Pour qu'une fonction fonctionne correctement, vous devez associer l'ID à l'implémentation JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ae659-115">For a function to work properly, you need to associate the id with the JavaScript implementation.</span></span> <span data-ttu-id="ae659-116">Vérifiez qu'il existe une association, sinon la fonction ne sera pas appelée.</span><span class="sxs-lookup"><span data-stu-id="ae659-116">Make sure there is an association, otherwise the function will not be called.</span></span>

<span data-ttu-id="ae659-117">L’exemple de code suivant montre comment procéder à cette association.</span><span class="sxs-lookup"><span data-stu-id="ae659-117">The following code sample shows how to do this association.</span></span> <span data-ttu-id="ae659-118">L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="ae659-118">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="ae659-119">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="ae659-119">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="ae659-120">Utilisez uniquement des lettres majuscules d’une fonction `name` et `id` dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="ae659-120">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="ae659-121">N’utilisez pas un mélange de cas ou uniquement des lettres minuscules.</span><span class="sxs-lookup"><span data-stu-id="ae659-121">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="ae659-122">Si vous le faites, vous risquez de finir avec deux valeurs différentes uniquement par la casse ,cela entraînera un remplacement involontaire de vos fonctions.</span><span class="sxs-lookup"><span data-stu-id="ae659-122">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="ae659-123">Par exemple, un objet de fonction à une `id` valeur**ajouter** peut être remplacé par déclaration plus loin dans le fichier d’objet de fonction avec une`id` valeur**AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="ae659-123">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="ae659-124">De plus, la `name` propriété définit le nom de la fonction que les utilisateurs finaux verront dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ae659-124">Additionally, the `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="ae659-125">Utiliser des lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonction intégrée sont en majuscules.</span><span class="sxs-lookup"><span data-stu-id="ae659-125">Using uppercase letters for the name of each custom function provides a consistent experience in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="ae659-126">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="ae659-126">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="ae659-127">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="ae659-127">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="ae659-128">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.</span><span class="sxs-lookup"><span data-stu-id="ae659-128">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="ae659-129">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="ae659-129">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="ae659-130">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="ae659-130">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="ae659-131">Dans le fichier JavaScript, spécifiez tous les mappages de fonctions personnalisées dans le même emplacement.</span><span class="sxs-lookup"><span data-stu-id="ae659-131">In the JavaScript file, specify all custom function associations in the same location.</span></span> <span data-ttu-id="ae659-132">Par exemple, le code suivant définit deux fonctions personnalisées et indique ensuite les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="ae659-132">For example, the following code sample defines two custom functions and then specifies the association information for both functions.</span></span>

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

    <span data-ttu-id="ae659-133">L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ae659-133">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="ae659-134">Notez que les propriétés`id` et `name`sont en majuscules dans ce fichier.</span><span class="sxs-lookup"><span data-stu-id="ae659-134">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="ae659-135">Déclarer des paramètres facultatifs</span><span class="sxs-lookup"><span data-stu-id="ae659-135">Declaring optional parameters</span></span> 

<span data-ttu-id="ae659-136">Dans Excel pour Windows (version 1812 ou version ultérieure), vous pouvez déclarer des paramètres facultatifs pour vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="ae659-136">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="ae659-137">Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses.</span><span class="sxs-lookup"><span data-stu-id="ae659-137">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="ae659-138">Par exemple, une fonction `FOO` avec un paramètre obligatoire appelé`parameter1` et un autre paramètre facultatif appelé `parameter2` apparaîtra sous la forme `=FOO(parameter1, [parameter2])` dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ae659-138">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="ae659-139">Pour rendre un paramètre facultatif, ajouter `"optional": true` au paramètre dans le fichier de métadonnées JSON qui définit la fonction.</span><span class="sxs-lookup"><span data-stu-id="ae659-139">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="ae659-140">L’exemple suivant montre comment cela peut se présenter pour la fonction `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="ae659-140">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="ae659-141">Vous pouvez remarquer que le paramètre facultatif `[third]` suit deux paramètres requis.</span><span class="sxs-lookup"><span data-stu-id="ae659-141">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="ae659-142">Les paramètres obligatoires apparaissent en premier dans l’interface utilisateur formule d’Excel.</span><span class="sxs-lookup"><span data-stu-id="ae659-142">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="ae659-143">Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qu’il se passe lorsque les paramètres facultatifs ne sont pas définis.</span><span class="sxs-lookup"><span data-stu-id="ae659-143">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="ae659-144">Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="ae659-144">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="ae659-145">Si le paramètre`zipCode` n’est pas défini, la valeur par défaut est définie sur 98052.</span><span class="sxs-lookup"><span data-stu-id="ae659-145">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="ae659-146">Si le paramètre`dayOfWeek` n’est pas défini, la valeur par défaut est définie à mercredi.</span><span class="sxs-lookup"><span data-stu-id="ae659-146">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="ae659-147">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="ae659-147">Additional considerations</span></span>

<span data-ttu-id="ae659-148">Pour créer un complément qui s’exécute sur plusieurs plateformes (l’un des clients clés des compléments Office), vous ne devez pas accéder au Document DOM (Object Model) dans les fonctions personnalisées ou utiliser de bibliothèques comme jQuery qui dépendent du DOM.</span><span class="sxs-lookup"><span data-stu-id="ae659-148">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="ae659-149">Sur Excel pour Windows, où les fonctions personnalisées utilisent l’[exécution JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="ae659-149">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="ae659-150">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ae659-150">See also</span></span>

* [<span data-ttu-id="ae659-151">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="ae659-151">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="ae659-152">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ae659-152">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ae659-153">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="ae659-153">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="ae659-154">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="ae659-154">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="ae659-155">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="ae659-155">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
