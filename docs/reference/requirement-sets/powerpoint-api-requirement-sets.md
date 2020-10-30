---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour PowerPoint.
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774727"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="080e9-103">Séries de conditions requises de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="080e9-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="080e9-p101">Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="080e9-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="080e9-107">Le tableau suivant répertorie les ensembles de conditions requises pour PowerPoint, les applications clientes Office qui prennent en charge ces ensembles de conditions requises et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="080e9-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="080e9-108">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="080e9-108">Requirement set</span></span>  |  <span data-ttu-id="080e9-109">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="080e9-109">Office on Windows</span></span><br><span data-ttu-id="080e9-110">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="080e9-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="080e9-111">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="080e9-111">Office on iPad</span></span><br><span data-ttu-id="080e9-112">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="080e9-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="080e9-113">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="080e9-113">Office on Mac</span></span><br><span data-ttu-id="080e9-114">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="080e9-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="080e9-115">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="080e9-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="080e9-116">Aperçu</span><span class="sxs-lookup"><span data-stu-id="080e9-116">Preview</span></span>](powerpoint-preview-apis.md)  | <span data-ttu-id="080e9-117">Veuillez utiliser la dernière version d’Office pour tester les API d’aperçu (vous devrez peut-être adhérer au [programme Office Insider](https://insider.office.com)).</span><span class="sxs-lookup"><span data-stu-id="080e9-117">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)).</span></span> |
| <span data-ttu-id="080e9-118">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="080e9-118">PowerPointApi 1.1</span></span> | <span data-ttu-id="080e9-119">Version 1810 (Build 11001.20074) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="080e9-119">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="080e9-120">2.17 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="080e9-120">2.17 or later</span></span> | <span data-ttu-id="080e9-121">16.19 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="080e9-121">16.19 or later</span></span> | <span data-ttu-id="080e9-122">Octobre 2018</span><span class="sxs-lookup"><span data-stu-id="080e9-122">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="080e9-123">Numéros de version et de build d’Office</span><span class="sxs-lookup"><span data-stu-id="080e9-123">Office versions and build numbers</span></span>

<span data-ttu-id="080e9-124">Pour plus d’informations sur les versions et les numéros de build d’Office, voir :</span><span class="sxs-lookup"><span data-stu-id="080e9-124">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="080e9-125">API JavaScript pour PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="080e9-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="080e9-126">L’API JavaScript PowerPoint 1.1 inclut une [seule API pour créer une nouvelle présentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span><span class="sxs-lookup"><span data-stu-id="080e9-126">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="080e9-127">Pour plus d’informations sur l’API, consultez [Créer une présentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span><span class="sxs-lookup"><span data-stu-id="080e9-127">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="080e9-128">Utiliser les conditions requises PowerPoint au moment de l’exécution et dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="080e9-128">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="080e9-129">Cette section suppose que vous êtes familiarisé avec les rubriques [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md) et [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="080e9-129">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="080e9-130">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="080e9-130">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="080e9-131">Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si une application Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="080e9-131">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="080e9-132">Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="080e9-132">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="080e9-133">L’exemple de code suivant montre comment déterminer si l’application Office dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.</span><span class="sxs-lookup"><span data-stu-id="080e9-133">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="080e9-134">Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="080e9-134">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="080e9-135">Vous pouvez utiliser l’[élément Requirements](../manifest/requirements.md) dans le manifeste de complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API que votre complément doit activer.</span><span class="sxs-lookup"><span data-stu-id="080e9-135">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="080e9-136">Si la plateforme ou l’application Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiées dans l’élément `Requirements` du manifeste, le complément ne s’exécute pas dans cette application ou plateforme et ne s’affiche pas dans la liste de compléments dans **Mes compléments** .Si votre complément requiert une configuration spécifique pour les fonctionnalités complètes, mais qu’il peut fournir une valeur même pour les utilisateurs sur les plateformes qui ne prennent pas en charge la condition requise, nous vous recommandons de vérifier la prise en charge des exigences au moment de l’exécution, comme décrit ci-dessus, au lieu de définir la prise en charge de la condition requise dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="080e9-136">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins** . If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="080e9-137">L’exemple de code suivant montre l’élément `Requirements` dans un manifeste indiquant que le complément doit être chargé dans toutes les applications clientes Office prenant en charge l’ensemble de conditions requises PowerPointApi version 1.1 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="080e9-137">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="080e9-138">Séries de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="080e9-138">Office Common API requirement sets</span></span>

<span data-ttu-id="080e9-139">La plupart des fonctionnalités du complément PowerPoint proviennent de la série courante d’API.</span><span class="sxs-lookup"><span data-stu-id="080e9-139">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="080e9-140">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="080e9-140">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="080e9-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="080e9-141">See also</span></span>

- [<span data-ttu-id="080e9-142">Documentation référence de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="080e9-142">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="080e9-143">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="080e9-143">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="080e9-144">Spécifier les applications Office et les exigences de l’API</span><span class="sxs-lookup"><span data-stu-id="080e9-144">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="080e9-145">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="080e9-145">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
