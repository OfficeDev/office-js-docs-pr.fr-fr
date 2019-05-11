---
ms.date: 05/08/2019
description: Découvrez les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.
title: Meilleures pratiques pour l’utilisation des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: d825f5a9f14e240ca5af3c3325cb646248d99ca9
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952102"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="3d120-103">Meilleures pratiques pour l’utilisation des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3d120-103">Custom functions best practices</span></span>

<span data-ttu-id="3d120-104">Cet article décrit les meilleures pratiques pour le développement des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="3d120-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="3d120-105">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="3d120-105">Associating function names with JSON metadata</span></span>

<span data-ttu-id="3d120-106">Comme décrit dans l’article[vue d’ensemble de fonctions personnalisées](custom-functions-overview.md), un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier de script (JavaScript ou machine à écrire) pour former une fonction complète.</span><span class="sxs-lookup"><span data-stu-id="3d120-106">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="3d120-107">Si vous utilisez `yo office` les métadonnées JSON, vous pouvez les générer à partir des commentaires de code.</span><span class="sxs-lookup"><span data-stu-id="3d120-107">If you are using `yo office` the JSON metadata can be generated from the code comments.</span></span> <span data-ttu-id="3d120-108">Dans le cas contraire, vous devez générer le fichier de métadonnées JSON manuellement.</span><span class="sxs-lookup"><span data-stu-id="3d120-108">Otherwise you need to build the JSON metadata file manually.</span></span>

<span data-ttu-id="3d120-109">Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de `id` la fonction à l’implémentation JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3d120-109">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="3d120-110">Vérifiez qu’il existe une association, sinon la fonction ne sera pas appelée.</span><span class="sxs-lookup"><span data-stu-id="3d120-110">Make sure there is an association, otherwise the function will not be called.</span></span> <span data-ttu-id="3d120-111">L’exemple de code suivant montre comment effectuer l’Association à l' `CustomFunctions.associate()` aide de la méthode.</span><span class="sxs-lookup"><span data-stu-id="3d120-111">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="3d120-112">L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="3d120-112">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="3d120-113">Le code JSON suivant illustre les métadonnées JSON associées au code JavaScript de fonction personnalisée précédent.</span><span class="sxs-lookup"><span data-stu-id="3d120-113">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


<span data-ttu-id="3d120-114">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="3d120-114">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="3d120-115">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="3d120-115">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="3d120-116">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="3d120-116">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="3d120-117">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.</span><span class="sxs-lookup"><span data-stu-id="3d120-117">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

* <span data-ttu-id="3d120-118">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="3d120-118">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="3d120-119">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="3d120-119">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="3d120-120">Dans le fichier JavaScript, spécifiez une association de fonctions `CustomFunctions.associate` personnalisées à l’aide de after each.</span><span class="sxs-lookup"><span data-stu-id="3d120-120">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="3d120-121">L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3d120-121">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="3d120-122">Les `id` valeurs `name` de la propriété et sont en majuscules, ce qui est recommandé lors de la description de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="3d120-122">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="3d120-123">Vous n’avez besoin d’ajouter ce JSON que si vous préparez votre propre fichier JSON manuellement et non à l’aide de la génération automatique.</span><span class="sxs-lookup"><span data-stu-id="3d120-123">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="3d120-124">Pour plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="3d120-124">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="3d120-125">Considérations supplémentaires</span><span class="sxs-lookup"><span data-stu-id="3d120-125">Additional considerations</span></span>

<span data-ttu-id="3d120-126">Évitez d’accéder directement ou indirectement au modèle DOM (Document Object Model) (par exemple, à l’aide de jQuery) à partir de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="3d120-126">Avoid accessing the Document Object Model (DOM) directly or indirectly (for example, using jQuery) from your custom function.</span></span> <span data-ttu-id="3d120-127">Dans Excel sur Windows, où les fonctions personnalisées utilisent le [Runtime JavaScript](custom-functions-runtime.md), les fonctions personnalisées ne peuvent pas accéder au DOM.</span><span class="sxs-lookup"><span data-stu-id="3d120-127">In Excel on Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="3d120-128">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="3d120-128">Next steps</span></span>
<span data-ttu-id="3d120-129">Découvrez comment [effectuer des requêtes Web avec des fonctions personnalisées](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="3d120-129">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3d120-130">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3d120-130">See also</span></span>

* [<span data-ttu-id="3d120-131">Générer automatiquement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3d120-131">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="3d120-132">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="3d120-132">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3d120-133">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="3d120-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
