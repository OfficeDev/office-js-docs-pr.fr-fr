---
title: Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement
description: Détails sur l’ensemble de conditions requises pour ExcelApiOnline
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757491"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="9f82f-103">Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement</span><span class="sxs-lookup"><span data-stu-id="9f82f-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="9f82f-104">L' `ExcelApiOnline` ensemble de conditions requises est un ensemble de conditions requises spéciales qui inclut des fonctionnalités qui sont disponibles uniquement pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="9f82f-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="9f82f-105">Les API de cet ensemble de conditions requises sont considérées comme des API de production (non soumises à des modifications structurelles ou comportementales non documentées) pour l’hôte Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="9f82f-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host.</span></span> <span data-ttu-id="9f82f-106">`ExcelApiOnline`sont considérés comme des API de « préversion » pour les autres plateformes (Windows, Mac, iOS) et ne sont peut-être pas pris en charge par aucune de ces plateformes.</span><span class="sxs-lookup"><span data-stu-id="9f82f-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="9f82f-107">Lorsque les API dans `ExcelApiOnline` l’ensemble de conditions requises sont prises en charge sur toutes les plateformes, elles seront ajoutées`ExcelApi 1.[NEXT]`à l’ensemble de conditions requises publié suivant ().</span><span class="sxs-lookup"><span data-stu-id="9f82f-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="9f82f-108">Une fois que cette nouvelle exigence est publique, ces API seront supprimées de `ExcelApiOnline`.</span><span class="sxs-lookup"><span data-stu-id="9f82f-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="9f82f-109">Imaginez qu’il s’agit d’un processus de promotion similaire, qui passe de l’aperçu à la version Release.</span><span class="sxs-lookup"><span data-stu-id="9f82f-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9f82f-110">`ExcelApiOnline`est un sur-ensemble du jeu de conditions requises le plus récent.</span><span class="sxs-lookup"><span data-stu-id="9f82f-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9f82f-111">`ExcelApiOnline 1.1`est la seule version des API en ligne uniquement.</span><span class="sxs-lookup"><span data-stu-id="9f82f-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="9f82f-112">En effet, Excel sur le Web disposera toujours d’une seule version disponible pour les utilisateurs qui est la version la plus récente.</span><span class="sxs-lookup"><span data-stu-id="9f82f-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="9f82f-113">Utilisation recommandée</span><span class="sxs-lookup"><span data-stu-id="9f82f-113">Recommended usage</span></span>

<span data-ttu-id="9f82f-114">Étant `ExcelApiOnline` donné que les API sont uniquement prises en charge par Excel sur le Web, votre complément doit vérifier si l’ensemble de conditions requises est pris en charge avant d’appeler ces API.</span><span class="sxs-lookup"><span data-stu-id="9f82f-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="9f82f-115">Cela évite d’appeler une API en ligne uniquement sur une autre plateforme.</span><span class="sxs-lookup"><span data-stu-id="9f82f-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="9f82f-116">Une fois que l’API se trouve dans un ensemble de conditions requises entre plateformes, vous `isSetSupported` devez supprimer ou modifier la vérification.</span><span class="sxs-lookup"><span data-stu-id="9f82f-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="9f82f-117">Cette opération active la fonctionnalité de votre complément sur d’autres plateformes.</span><span class="sxs-lookup"><span data-stu-id="9f82f-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="9f82f-118">Veillez à tester la fonctionnalité sur ces plateformes lors de l’exécution de cette modification.</span><span class="sxs-lookup"><span data-stu-id="9f82f-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9f82f-119">Votre manifeste ne peut `ExcelApiOnline 1.1` pas spécifier comme condition d’activation.</span><span class="sxs-lookup"><span data-stu-id="9f82f-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="9f82f-120">Il ne s’agit pas d’une valeur valide à utiliser dans l' [élément Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="9f82f-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="9f82f-121">Liste des API</span><span class="sxs-lookup"><span data-stu-id="9f82f-121">API list</span></span>

<span data-ttu-id="9f82f-122">Il n’existe actuellement aucune API en ligne uniquement.</span><span class="sxs-lookup"><span data-stu-id="9f82f-122">There are currently no online-only APIs.</span></span> <span data-ttu-id="9f82f-123">Vérifiez à nouveau que de nouvelles fonctionnalités sont ajoutées à Excel sur le Web et prises en charge par les API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="9f82f-123">Check back as new features are added to Excel on the web and supported by the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="9f82f-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9f82f-124">See also</span></span>

- [<span data-ttu-id="9f82f-125">Documentation référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="9f82f-125">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="9f82f-126">Version d’évaluation API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="9f82f-126">Excel JavaScript preview APIs</span></span>](./excel-preview-apis.md)
- [<span data-ttu-id="9f82f-127">Ensembles de conditions requises de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="9f82f-127">Excel JavaScript API requirement sets</span></span>](./excel-api-requirement-sets.md)