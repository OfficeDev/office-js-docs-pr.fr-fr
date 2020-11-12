---
title: Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement
description: Détails sur l’ensemble de conditions requises pour ExcelApiOnline.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 16c96f413424d5fc85a21419fb72cf6580c1ac18
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996528"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="2ec3c-103">Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement</span><span class="sxs-lookup"><span data-stu-id="2ec3c-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="2ec3c-104">L' `ExcelApiOnline` ensemble de conditions requises est un ensemble de conditions requises spéciales qui inclut des fonctionnalités qui sont disponibles uniquement pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="2ec3c-105">Les API de cet ensemble de conditions requises sont considérées comme des API de production (non soumises à des modifications structurelles ou comportementales non documentées) pour l’application Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="2ec3c-106">`ExcelApiOnline` sont considérés comme des API de « préversion » pour les autres plateformes (Windows, Mac, iOS) et ne sont peut-être pas pris en charge par aucune de ces plateformes.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="2ec3c-107">Lorsque les API dans l' `ExcelApiOnline` ensemble de conditions requises sont prises en charge sur toutes les plateformes, elles seront ajoutées à l’ensemble de conditions requises publié suivant ( `ExcelApi 1.[NEXT]` ).</span><span class="sxs-lookup"><span data-stu-id="2ec3c-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="2ec3c-108">Une fois que cette nouvelle exigence est publique, ces API seront supprimées de `ExcelApiOnline` .</span><span class="sxs-lookup"><span data-stu-id="2ec3c-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="2ec3c-109">Imaginez qu’il s’agit d’un processus de promotion similaire, qui passe de l’aperçu à la version Release.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2ec3c-110">`ExcelApiOnline` est un sur-ensemble du jeu de conditions requises le plus récent.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2ec3c-111">`ExcelApiOnline 1.1` est la seule version des API en ligne uniquement.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="2ec3c-112">En effet, Excel sur le Web disposera toujours d’une seule version disponible pour les utilisateurs qui est la version la plus récente.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="2ec3c-113">Utilisation recommandée</span><span class="sxs-lookup"><span data-stu-id="2ec3c-113">Recommended usage</span></span>

<span data-ttu-id="2ec3c-114">Étant donné que `ExcelApiOnline` les API sont uniquement prises en charge par Excel sur le Web, votre complément doit vérifier si l’ensemble de conditions requises est pris en charge avant d’appeler ces API.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="2ec3c-115">Cela évite d’appeler une API en ligne uniquement sur une autre plateforme.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="2ec3c-116">Une fois que l’API se trouve dans un ensemble de conditions requises entre plateformes, vous devez supprimer ou modifier la `isSetSupported` vérification.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="2ec3c-117">Cette opération active la fonctionnalité de votre complément sur d’autres plateformes.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="2ec3c-118">Veillez à tester la fonctionnalité sur ces plateformes lors de l’exécution de cette modification.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2ec3c-119">Votre manifeste ne peut pas spécifier `ExcelApiOnline 1.1` comme condition d’activation.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="2ec3c-120">Il ne s’agit pas d’une valeur valide à utiliser dans l' [élément Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="2ec3c-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="2ec3c-121">Liste des API</span><span class="sxs-lookup"><span data-stu-id="2ec3c-121">API list</span></span>

| <span data-ttu-id="2ec3c-122">Class</span><span class="sxs-lookup"><span data-stu-id="2ec3c-122">Class</span></span> | <span data-ttu-id="2ec3c-123">Champs</span><span class="sxs-lookup"><span data-stu-id="2ec3c-123">Fields</span></span> | <span data-ttu-id="2ec3c-124">Description</span><span class="sxs-lookup"><span data-stu-id="2ec3c-124">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="2ec3c-125">Range</span><span class="sxs-lookup"><span data-stu-id="2ec3c-125">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="2ec3c-126">getMergedAreas()</span><span class="sxs-lookup"><span data-stu-id="2ec3c-126">getMergedAreas()</span></span>](/javascript/api/excel/excel.range#getmergedareas--)|<span data-ttu-id="2ec3c-127">Renvoie un objet RangeAreas qui représente les zones fusionnées dans cette plage.</span><span class="sxs-lookup"><span data-stu-id="2ec3c-127">Returns a RangeAreas object that represents the merged areas in this range.</span></span>|

## <a name="see-also"></a><span data-ttu-id="2ec3c-128">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2ec3c-128">See also</span></span>

- [<span data-ttu-id="2ec3c-129">Documentation référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="2ec3c-129">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="2ec3c-130">Version d’évaluation API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="2ec3c-130">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="2ec3c-131">Ensembles de conditions requises de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="2ec3c-131">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)
