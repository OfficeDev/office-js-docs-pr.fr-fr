---
title: Utilisation des substitutions étendues du manifeste
description: Découvrez comment configurer des fonctionnalités d’extensibilité avec des substitutions étendues du manifeste.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 4eb8936e8a01b81a3883f848446d20ebf4ecf863
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505569"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a><span data-ttu-id="e3bc3-103">Utilisation des substitutions étendues du manifeste</span><span class="sxs-lookup"><span data-stu-id="e3bc3-103">Work with Extended Overrides of the manifest</span></span>

<span data-ttu-id="e3bc3-104">Certaines fonctionnalités d’extensibilité des add-ins Office sont configurées avec des fichiers JSON hébergés sur votre serveur, et non avec le manifeste XML du module.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-104">Some extensibility features of Office Add-ins are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="e3bc3-105">Cet article part du principe que vous êtes familiarisé avec les manifestes de add-in Office et leur rôle dans les add-ins. Veuillez lire [le manifeste XML des add-ins Office,](add-in-manifests.md)si ce n’est pas le cas récemment.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-105">This article assumes that you're familiar with Office add-in manifests and their role in add-ins. Please read [Office Add-ins XML manifest](add-in-manifests.md), if you haven't recently.</span></span>

<span data-ttu-id="e3bc3-106">Le tableau suivant spécifie les fonctionnalités d’extensibilité qui nécessitent une substitution étendue, ainsi que des liens vers la documentation de la fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-106">The following table specifies the extensibility features that require an extended override along with links to documentation of the feature.</span></span>

| <span data-ttu-id="e3bc3-107">Fonctionnalité</span><span class="sxs-lookup"><span data-stu-id="e3bc3-107">Feature</span></span> | <span data-ttu-id="e3bc3-108">Instructions de développement</span><span class="sxs-lookup"><span data-stu-id="e3bc3-108">Development Instructions</span></span> |
| :----- | :----- |
| <span data-ttu-id="e3bc3-109">Raccourcis clavier</span><span class="sxs-lookup"><span data-stu-id="e3bc3-109">Keyboard shortcuts</span></span> | [<span data-ttu-id="e3bc3-110">Ajouter des raccourcis clavier personnalisés à vos add-ins Office</span><span class="sxs-lookup"><span data-stu-id="e3bc3-110">Add Custom keyboard shortcuts to your Office Add-ins</span></span>](../design/keyboard-shortcuts.md) |

<span data-ttu-id="e3bc3-111">Le schéma qui définit le format JSON est [un schéma de manifeste étendu.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="e3bc3-111">The schema that defines the JSON format is [extended-manifest schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!TIP]
> <span data-ttu-id="e3bc3-112">Cet article est quelque peu abstrait.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-112">This article is somewhat abstract.</span></span> <span data-ttu-id="e3bc3-113">Envisagez de lire l’un des articles du tableau pour clarifier les concepts.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-113">Consider reading one of the articles in the table to add clarity to the concepts.</span></span>

## <a name="tell-office-where-to-find-the-json-file"></a><span data-ttu-id="e3bc3-114">Indiquer à Office où trouver le fichier JSON</span><span class="sxs-lookup"><span data-stu-id="e3bc3-114">Tell Office where to find the JSON file</span></span>

<span data-ttu-id="e3bc3-115">Utilisez le manifeste pour indiquer à Office où trouver le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-115">Use the manifest to tell Office where to find the JSON file.</span></span> <span data-ttu-id="e3bc3-116">Juste *en dessous* (pas à l’intérieur) de l’élément dans le manifeste, ajoutez un élément `<VersionOverrides>` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="e3bc3-116">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="e3bc3-117">Définissez `Url` l’attribut sur l’URL complète d’un fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-117">Set the `Url` attribute to the full URL of a JSON file.</span></span> <span data-ttu-id="e3bc3-118">Voici un exemple de l’élément le plus `<ExtendedOverrides>` simple possible.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-118">The following is an example of the simplest possible `<ExtendedOverrides>` element.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="e3bc3-119">Voici un exemple de fichier JSON de remplacements étendu très simple.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-119">The following is an example of a very simple extended overrides JSON file.</span></span> <span data-ttu-id="e3bc3-120">Il affecte le raccourci clavier Ctrl+Shift+A à une fonction (définie ailleurs) qui ouvre le volet Des tâches du module.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-120">It assigns keyboard shortcut CTRL+SHIFT+A to a function (defined elsewhere) that opens the add-in's task pane.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a><span data-ttu-id="e3bc3-121">Localiser le fichier de remplacements étendu</span><span class="sxs-lookup"><span data-stu-id="e3bc3-121">Localize the extended overrides file</span></span>

<span data-ttu-id="e3bc3-122">Si votre add-in prend en charge plusieurs paramètres régionaux, vous pouvez utiliser l’attribut de l’élément pour pointer Office vers un `ResourceUrl` `<ExtendedOverrides>` fichier de ressources localisées.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-122">If your add-in supports multiple locales, you can use the `ResourceUrl` attribute of the `<ExtendedOverrides>` element to point Office to a file of localized resources.</span></span> <span data-ttu-id="e3bc3-123">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="e3bc3-123">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="e3bc3-124">Pour plus d’informations sur la création et l’utilisation du fichier de ressources, sur la façon de faire référence à ses ressources dans le fichier de remplacements étendu et pour les options supplémentaires non abordées ici, voir [Localize extended overrides](localization.md#localize-extended-overrides).</span><span class="sxs-lookup"><span data-stu-id="e3bc3-124">For more details about how to create and use the resources file, how to refer to its resources in the extended overrides file, and for additional options not discussed here, see [Localize extended overrides](localization.md#localize-extended-overrides).</span></span>
