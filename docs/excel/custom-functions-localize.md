---
ms.date: 05/03/2019
description: Localisez vos fonctions personnalisées Excel.
title: Localiser des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 5dbe2f78f1d24c3d8c8214f4e604e66f097adba3
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628031"
---
# <a name="localize-custom-functions"></a><span data-ttu-id="ad30f-103">Localiser des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ad30f-103">Localize custom functions</span></span>

<span data-ttu-id="ad30f-104">Vous pouvez localiser votre complément et vos noms de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="ad30f-104">You can localize both your add-in and your custom function names.</span></span> <span data-ttu-id="ad30f-105">Vous devez fournir des noms de fonctions localisées dans le fichier JSON des fonctions et fournir des informations de paramètres régionaux dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="ad30f-105">You need to provide localized function names in the functions' JSON file and provide locale information in the XML manifest file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!IMPORTANT]
> <span data-ttu-id="ad30f-106">Les métadonnées générées automatiquement ne fonctionnent pas pour la localisation, c’est pourquoi vous devez mettre à jour le fichier JSON manuellement.</span><span class="sxs-lookup"><span data-stu-id="ad30f-106">Autogenerated metadata doesn't work for localization so you need to update the JSON file manually.</span></span>

## <a name="localize-function-names"></a><span data-ttu-id="ad30f-107">Noms des fonctions de localisation</span><span class="sxs-lookup"><span data-stu-id="ad30f-107">Localize function names</span></span>

<span data-ttu-id="ad30f-108">Pour localiser vos fonctions personnalisées, créez un nouveau fichier de métadonnées JSON pour chaque langue.</span><span class="sxs-lookup"><span data-stu-id="ad30f-108">To localize your custom functions, create a new JSON metadata file for each language.</span></span> <span data-ttu-id="ad30f-109">Dans chaque fichier JSON de langue, `name` créez `description` et des propriétés dans la langue cible.</span><span class="sxs-lookup"><span data-stu-id="ad30f-109">In each language JSON file, create `name` and `description` properties in the target language.</span></span> <span data-ttu-id="ad30f-110">Le fichier par défaut pour l’anglais est nommé **functions. JSON**.</span><span class="sxs-lookup"><span data-stu-id="ad30f-110">The default file for English is named **functions.json**.</span></span> <span data-ttu-id="ad30f-111">Il est recommandé d’utiliser les paramètres régionaux dans le nom de fichier de chaque fichier JSON supplémentaire, comme les **fonctions-de-JSON** pour les identifier.</span><span class="sxs-lookup"><span data-stu-id="ad30f-111">It's recommended you use the locale in the filename for each additional JSON file, such as **functions-de.json** to help identify them.</span></span>

<span data-ttu-id="ad30f-112">Le `name` et `description` s’affichent dans Excel et sont localisés.</span><span class="sxs-lookup"><span data-stu-id="ad30f-112">The `name` and `description` appear in Excel and are localized.</span></span> <span data-ttu-id="ad30f-113">Toutefois, la `id` de chaque fonction n’est pas localisée.</span><span class="sxs-lookup"><span data-stu-id="ad30f-113">However, the `id` of each function is not localized.</span></span> <span data-ttu-id="ad30f-114">La `id` propriété indique comment Excel identifie votre fonction comme étant unique et ne doit pas être modifiée une fois qu’elle a été définie.</span><span class="sxs-lookup"><span data-stu-id="ad30f-114">The `id` property is how Excel identifies your function as unique and should not be changed once it is set.</span></span>

<span data-ttu-id="ad30f-115">Le code JSON suivant montre comment définir une fonction avec la `id` propriété «Multiply».</span><span class="sxs-lookup"><span data-stu-id="ad30f-115">The following JSON shows how to define a function with the `id` property "MULTIPLY."</span></span> <span data-ttu-id="ad30f-116">La `name` propriété `description` et de la fonction est localisée pour l’allemand.</span><span class="sxs-lookup"><span data-stu-id="ad30f-116">The `name` and `description` property of the function is localized for German.</span></span> <span data-ttu-id="ad30f-117">Chaque paramètre `name` et `description` est également localisé pour l’allemand.</span><span class="sxs-lookup"><span data-stu-id="ad30f-117">Each parameter `name` and `description` is also localized for German.</span></span>

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

<span data-ttu-id="ad30f-118">Comparez le JSON précédent avec le JSON suivant pour l’anglais.</span><span class="sxs-lookup"><span data-stu-id="ad30f-118">Compare the previous JSON with the following JSON for English.</span></span>

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

## <a name="localize-your-add-in"></a><span data-ttu-id="ad30f-119">Localiser votre complément</span><span class="sxs-lookup"><span data-stu-id="ad30f-119">Localize your add-in</span></span>

<span data-ttu-id="ad30f-120">Après avoir créé un fichier JSON pour chaque langue, vous devez mettre à jour votre fichier manifeste XML avec une valeur de remplacement pour chaque paramètre régional qui spécifie l’URL de chaque fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="ad30f-120">After creating a JSON file for each language, you need to update your XML manifest file with an override value for each locale that specifies the URL of each JSON metadata file.</span></span> <span data-ttu-id="ad30f-121">Le code XML de manifeste suivant affiche `en-us` les paramètres régionaux par défaut avec une URL de fichier `de-de` JSON de remplacement pour (Allemagne).</span><span class="sxs-lookup"><span data-stu-id="ad30f-121">The following manifest XML shows a default `en-us` locale with an override JSON file URL for `de-de` (Germany).</span></span> <span data-ttu-id="ad30f-122">Le fichier **Functions-de. JSON** contient les noms et les ID des fonctions localisées en allemand.</span><span class="sxs-lookup"><span data-stu-id="ad30f-122">The **functions-de.json** file contains the localized German function names and ids.</span></span>

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

<span data-ttu-id="ad30f-123">Pour plus d’informations sur le processus de localisation d’un complément, reportez-vous à la rubrique [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="ad30f-123">For more information on the process of localizing an add-in, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span></span>

## <a name="next-steps"></a><span data-ttu-id="ad30f-124">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="ad30f-124">Next steps</span></span>
<span data-ttu-id="ad30f-125">Découvrez [les conventions d’affectation de noms pour les fonctions personnalisées](custom-functions-naming.md) ou découvrir les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="ad30f-125">Learn about [naming conventions for custom functions](custom-functions-naming.md) or discover [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ad30f-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ad30f-126">See also</span></span>

* [<span data-ttu-id="ad30f-127">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ad30f-127">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ad30f-128">Générer automatiquement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ad30f-128">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="ad30f-129">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ad30f-129">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="ad30f-130">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="ad30f-130">Create custom functions in Excel</span></span>](custom-functions-overview.md)