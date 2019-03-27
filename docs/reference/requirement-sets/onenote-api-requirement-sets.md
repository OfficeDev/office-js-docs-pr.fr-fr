---
title: Ensembles de conditions requises de l’API JavaScript pour OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 287e405955477a98854b1df4a81fe90ec16e5bbc
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871597"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="2f837-102">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="2f837-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="2f837-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2f837-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="2f837-106">Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="2f837-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="2f837-107">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f837-107">Requirement set</span></span>  |  <span data-ttu-id="2f837-108">Office Online</span><span class="sxs-lookup"><span data-stu-id="2f837-108">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="2f837-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="2f837-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="2f837-110">Septembre 2016</span><span class="sxs-lookup"><span data-stu-id="2f837-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="2f837-111">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="2f837-111">Office Common API requirement sets</span></span>

<span data-ttu-id="2f837-112">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="2f837-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="2f837-113">API JavaScript pour OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="2f837-113">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="2f837-114">L’API JavaScript 1.1 pour OneNote est la première version de l’API.</span><span class="sxs-lookup"><span data-stu-id="2f837-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="2f837-115">Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[Récapitulatif de programmation API JavaScript pour OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="2f837-115">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="2f837-116">Vérification de la prise en charge d’un ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="2f837-116">Runtime requirement support check</span></span>

<span data-ttu-id="2f837-117">Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge un ensemble de conditions requises d’API en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="2f837-117">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="2f837-118">Vérification de la prise en charge des conditions requises basée sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="2f837-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="2f837-p103">Utilisez l’élément Conditions requises dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément Conditions requises, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="2f837-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="2f837-121">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="2f837-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="2f837-122">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2f837-122">See also</span></span>

- [<span data-ttu-id="2f837-123">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f837-123">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="2f837-124">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="2f837-124">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="2f837-125">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="2f837-125">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
