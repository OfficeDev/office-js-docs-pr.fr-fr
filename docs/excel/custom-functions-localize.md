---
ms.date: 04/29/2020
description: Localisez vos fonctions personnalisées Excel.
title: Localiser des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 001045f82634d7e96c4d4515ccd87b5cfaf2cd1c
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275965"
---
# <a name="localize-custom-functions"></a><span data-ttu-id="400af-103">Localiser des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="400af-103">Localize custom functions</span></span>

<span data-ttu-id="400af-104">Vous pouvez localiser votre complément et vos noms de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="400af-104">You can localize both your add-in and your custom function names.</span></span> <span data-ttu-id="400af-105">Pour ce faire, fournissez des noms de fonction localisés dans le fichier JSON des fonctions et des informations de paramètres régionaux dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="400af-105">To do so, provide localized function names in the functions' JSON file and locale information in the XML manifest file.</span></span>

>[!IMPORTANT]
> <span data-ttu-id="400af-106">Les métadonnées générées automatiquement ne fonctionnent pas pour la localisation, c’est pourquoi vous devez mettre à jour le fichier JSON manuellement.</span><span class="sxs-lookup"><span data-stu-id="400af-106">Auto-generated metadata doesn't work for localization so you need to update the JSON file manually.</span></span> <span data-ttu-id="400af-107">Pour savoir comment procéder, consultez la rubrique [Metadata for Custom Functions in Excel](custom-functions-json.md)</span><span class="sxs-lookup"><span data-stu-id="400af-107">To learn how to do this, see [Metadata for custom functions in Excel](custom-functions-json.md)</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a><span data-ttu-id="400af-108">Noms des fonctions de localisation</span><span class="sxs-lookup"><span data-stu-id="400af-108">Localize function names</span></span>

<span data-ttu-id="400af-109">Pour localiser vos fonctions personnalisées, créez un nouveau fichier de métadonnées JSON pour chaque langue.</span><span class="sxs-lookup"><span data-stu-id="400af-109">To localize your custom functions, create a new JSON metadata file for each language.</span></span> <span data-ttu-id="400af-110">Dans chaque fichier JSON de langue, créez `name` et des `description` Propriétés dans la langue cible.</span><span class="sxs-lookup"><span data-stu-id="400af-110">In each language JSON file, create `name` and `description` properties in the target language.</span></span> <span data-ttu-id="400af-111">Le fichier par défaut pour l’anglais est nommé **functions. JSON**.</span><span class="sxs-lookup"><span data-stu-id="400af-111">The default file for English is named **functions.json**.</span></span> <span data-ttu-id="400af-112">Utilisez les paramètres régionaux dans le nom de fichier de tous les fichiers JSON supplémentaires, tels que les **fonctions-de-JSON** pour les identifier.</span><span class="sxs-lookup"><span data-stu-id="400af-112">Use the locale in the filename for each additional JSON file, such as **functions-de.json** to help identify them.</span></span>

<span data-ttu-id="400af-113">Le `name` et `description` s’affichent dans Excel et sont localisés.</span><span class="sxs-lookup"><span data-stu-id="400af-113">The `name` and `description` appear in Excel and are localized.</span></span> <span data-ttu-id="400af-114">Toutefois, la `id` de chaque fonction n’est pas localisée.</span><span class="sxs-lookup"><span data-stu-id="400af-114">However, the `id` of each function isn't localized.</span></span> <span data-ttu-id="400af-115">La `id` propriété indique comment Excel identifie votre fonction comme étant unique et ne doit pas être modifiée une fois qu’elle a été définie.</span><span class="sxs-lookup"><span data-stu-id="400af-115">The `id` property is how Excel identifies your function as unique and shouldn't be changed once it is set.</span></span>

<span data-ttu-id="400af-116">Le code JSON suivant montre comment définir une fonction avec la `id` propriété « Multiply ».</span><span class="sxs-lookup"><span data-stu-id="400af-116">The following JSON shows how to define a function with the `id` property "MULTIPLY."</span></span> <span data-ttu-id="400af-117">La `name` `description` propriété et de la fonction est localisée pour l’allemand.</span><span class="sxs-lookup"><span data-stu-id="400af-117">The `name` and `description` property of the function is localized for German.</span></span> <span data-ttu-id="400af-118">Chaque paramètre `name` et `description` est également localisé pour l’allemand.</span><span class="sxs-lookup"><span data-stu-id="400af-118">Each parameter `name` and `description` is also localized for German.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

<span data-ttu-id="400af-119">Comparez le JSON précédent avec le JSON suivant pour l’anglais.</span><span class="sxs-lookup"><span data-stu-id="400af-119">Compare the previous JSON with the following JSON for English.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a><span data-ttu-id="400af-120">Localiser votre complément</span><span class="sxs-lookup"><span data-stu-id="400af-120">Localize your add-in</span></span>

<span data-ttu-id="400af-121">Après avoir créé un fichier JSON pour chaque langue, mettez à jour votre fichier manifeste XML avec une valeur de remplacement pour chaque paramètre régional qui spécifie l’URL de chaque fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="400af-121">After creating a JSON file for each language, update your XML manifest file with an override value for each locale that specifies the URL of each JSON metadata file.</span></span> <span data-ttu-id="400af-122">Le code XML de manifeste suivant affiche les `en-us` paramètres régionaux par défaut avec une URL de fichier JSON de remplacement pour `de-de` (Allemagne).</span><span class="sxs-lookup"><span data-stu-id="400af-122">The following manifest XML shows a default `en-us` locale with an override JSON file URL for `de-de` (Germany).</span></span> <span data-ttu-id="400af-123">Le fichier **Functions-de. JSON** contient les noms et les ID des fonctions localisées en allemand.</span><span class="sxs-lookup"><span data-stu-id="400af-123">The **functions-de.json** file contains the localized German function names and ids.</span></span>

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

<span data-ttu-id="400af-124">Pour plus d’informations sur le processus de localisation d’un complément, reportez-vous à la rubrique [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="400af-124">For more information on the process of localizing an add-in, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span></span>

## <a name="next-steps"></a><span data-ttu-id="400af-125">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="400af-125">Next steps</span></span>
<span data-ttu-id="400af-126">Découvrez [les conventions d’affectation de noms pour les fonctions personnalisées](custom-functions-naming.md) ou découvrir les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="400af-126">Learn about [naming conventions for custom functions](custom-functions-naming.md) or discover [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="400af-127">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="400af-127">See also</span></span>

* [<span data-ttu-id="400af-128">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="400af-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="400af-129">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="400af-129">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="400af-130">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="400af-130">Create custom functions in Excel</span></span>](custom-functions-overview.md)
