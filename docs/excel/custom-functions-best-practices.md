---
ms.date: 10/03/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: 218e62cd074ccf3f3708bba90c938f7ddef059cb
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579820"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="e3764-103">Meilleures pratiques pour les fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="e3764-103">Custom functions best practices</span></span>

<span data-ttu-id="e3764-104">Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="e3764-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="e3764-105">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="e3764-105">Error handling</span></span>

<span data-ttu-id="e3764-p101">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="e3764-p101">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="e3764-109">Débogage</span><span class="sxs-lookup"><span data-stu-id="e3764-109">Debugging</span></span>

<span data-ttu-id="e3764-p102">Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à d’abord [charger en parallèle](../testing/sideload-office-add-ins-for-testing.md) votre complément dans **Excel Online**. Vous pouvez ensuite déboguer vos fonctions personnalisées à l’aide de l' [outil F12 de débogage natif de votre navigateur](../testing/debug-add-ins-in-office-online.md) en combinaison avec les techniques suivantes :</span><span class="sxs-lookup"><span data-stu-id="e3764-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="e3764-112">Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.</span><span class="sxs-lookup"><span data-stu-id="e3764-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="e3764-p103">Utilisez les instructions `debugger;` au sein de votre code des fonctions personnalisées pour spécifier les points d’arrêt où l’exécution s’interrompra lorsque la fenêtre F12 est ouverte. Par exemple, si la fonction suivante s’exécute alors que la fenêtre F12 est ouverte, l’exécution s’interrompra sur l’instruction `debugger;`, ce qui vous permettra d’inspecter manuellement les valeurs de paramètre avant le retour de la fonction.L’instruction `debugger;` n’a aucun effet dans Excel Online lorsque la fenêtre F12 n’est pas ouverte. Actuellement, les instructions `debugger;` n’ont aucun effet dans Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="e3764-p103">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open. For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns. The `debugger;` statement has no effect in Excel Online when the F12 window is not open. Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="e3764-117">Si votre complément ne parvient pas à s’enregistrer, [vérifiez que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web qui héberge votre application de complément.</span><span class="sxs-lookup"><span data-stu-id="e3764-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="e3764-118">Si vous testez votre complément dans Office sur le bureau Windows, vous pouvez activer la [journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) pour résoudre les problèmes du fichier manifeste XML de votre complément, ainsi que plusieurs conditions d’installation et d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e3764-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="e3764-119">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="e3764-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="e3764-p104">Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires à Excel pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux. En outre, dans le fichier JavaScript qui définit vos fonctions personnalisées, vous devez fournir les informations pour spécifier l’objet de fonction dans le fichier de métadonnées JSON correspondant à chaque fonction personnalisée dans le fichier JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e3764-p104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="e3764-122">Par exemple, l’exemple de code suivant définit la fonction personnalisée `add`, puis spécifie que la fonction `add` correspond à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **ADD**.</span><span class="sxs-lookup"><span data-stu-id="e3764-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="e3764-123">Gardez à l’esprit les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et en spécifiant les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="e3764-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="e3764-p105">Dans le fichier JavaScript, spécifiez les noms de fonction en casse mixte. Par exemple, le nom de la fonction `addTenToInput` est écrit en casse mixte : le premier mot dans le nom commence par une lettre minuscule, et chaque mot suivant dans le nom commence par une lettre majuscule.</span><span class="sxs-lookup"><span data-stu-id="e3764-p105">In the JavaScript file, specify function names in camelCase. For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="e3764-p106">Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque propriété `name` en majuscules. La propriété `name`  définit le nom de la fonction que les utilisateurs finaux verront s’afficher dans Excel. L’utilisation de lettres majuscules pour le nom de chaque fonction personnalisée fournit une expérience cohérente pour les utilisateurs finaux dans Excel, où tous les noms de fonctions intégrées sont en majuscules.</span><span class="sxs-lookup"><span data-stu-id="e3764-p106">In the JSON metadata file, specify the value of each `name` property in uppercase. The `name` property defines the function name that end users will see in Excel. Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="e3764-p107">Dans le fichier de métadonnées JSON, spécifiez la valeur de chaque propriété `id` en majuscules. Ainsi, il est évident quelle partie de l’instruction `CustomFunctionMappings`  dans votre code JavaScript correspond à la propriété `id`    dans le fichier de métadonnées JSON (à condition que votre nom de la fonction utilise CamelCase, comme indiqué précédemment).</span><span class="sxs-lookup"><span data-stu-id="e3764-p107">In the JSON metadata file, specify the value of each `id` property in uppercase. Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="e3764-131">Dans le fichier de métadonnées JSON, assurez-vous que la valeur de chaque propriété`id` contient uniquement des caractères alphanumériques et des points.</span><span class="sxs-lookup"><span data-stu-id="e3764-131">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="e3764-p108">Dans le fichier de métadonnées JSON, assurez-vous que la valeur de chaque propriété `id` est unique dans l’étendue du fichier. Autrement dit, deux objets fonctions dans le fichier de métadonnées ne doivent pas avoir la même valeur `id`. En outre, ne spécifiez pas deux valeurs `id`  dans le fichier de métadonnées qui diffèrent uniquement par la casse. Par exemple, ne définissez pas un objet fonction avec une valeur `id`  de **add** et un autre objet fonction avec une valeur `id`  de **ADD**.</span><span class="sxs-lookup"><span data-stu-id="e3764-p108">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value. Additionally, do not specify two `id` values in the metadata file that only differ by case. For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="e3764-p109">Ne modifiez pas la valeur d’une propriété `id` dans le fichier de métadonnées JSON après qu’il a été mappé à un nom de fonction JavaScript correspondant. Vous pouvez modifier le nom de la fonction que les utilisateurs voient dans Excel en mettant à jour la propriété `name`  dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une propriété `id`  une fois établie.</span><span class="sxs-lookup"><span data-stu-id="e3764-p109">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="e3764-p110">Dans le fichier JavaScript, spécifiez tous les mappages de fonctions personnalisées au même endroit. Par exemple, l’exemple de code suivant définit deux fonctions personnalisées puis spécifie les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="e3764-p110">In the JavaScript file, specify all custom function mappings in the same location. For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="e3764-140">L’exemple suivant montre les métadonnées JSON qui correspondent aux fonctions définies dans cet exemple de code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e3764-140">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="e3764-141">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="e3764-141">Additional considerations</span></span>

<span data-ttu-id="e3764-142">Pour créer un complément qui s’exécute sur plusieurs plates-formes (l’un des principaux clients des compléments Office), vous ne devez pas accéder au DOM (Document Object Model) dans les fonctions personnalisées ni utiliser des bibliothèques comme jQuery qui s’appuient sur le modèle DOM.</span><span class="sxs-lookup"><span data-stu-id="e3764-142">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="e3764-143">Dans Excel pour Windows, où les fonctions personnalisées utilisent le [runtime JavaScript](custom-functions-runtime.md), des fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="e3764-143">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="e3764-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e3764-144">See also</span></span>

* [<span data-ttu-id="e3764-145">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="e3764-145">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="e3764-146">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e3764-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="e3764-147">Runtime de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="e3764-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="e3764-148">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="e3764-148">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
