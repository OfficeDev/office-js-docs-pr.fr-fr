---
title: Résolution des problèmes liés aux compléments Excel
description: Découvrez comment résoudre les problèmes liés aux erreurs de développement dans les compléments Excel.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409388"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="c4a8f-103">Résolution des problèmes liés aux compléments Excel</span><span class="sxs-lookup"><span data-stu-id="c4a8f-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="c4a8f-104">Cet article traite de la résolution des problèmes propres à Excel.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="c4a8f-105">Veuillez utiliser l’outil de commentaires en bas de la page pour suggérer d’autres problèmes pouvant être ajoutés à l’article.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="c4a8f-106">Limitations de l’API lorsque le classeur actif bascule</span><span class="sxs-lookup"><span data-stu-id="c4a8f-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="c4a8f-107">Les compléments pour Excel sont conçus pour fonctionner sur un seul classeur à la fois.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="c4a8f-108">Des erreurs peuvent se produire lorsqu’un classeur distinct de celui qui exécute le complément obtient le focus.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="c4a8f-109">Cela se produit uniquement lorsque des méthodes particulières sont en cours d’appel lorsque le focus est modifié.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="c4a8f-110">Les API suivantes sont affectées par ce commutateur de classeurs :</span><span class="sxs-lookup"><span data-stu-id="c4a8f-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="c4a8f-111">sur les API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c4a8f-111">Excel JavaScript API</span></span> | <span data-ttu-id="c4a8f-112">Erreur générée</span><span class="sxs-lookup"><span data-stu-id="c4a8f-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="c4a8f-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="c4a8f-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="c4a8f-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="c4a8f-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c4a8f-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="c4a8f-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c4a8f-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="c4a8f-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c4a8f-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="c4a8f-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="c4a8f-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c4a8f-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="c4a8f-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="c4a8f-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="c4a8f-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="c4a8f-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="c4a8f-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="c4a8f-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="c4a8f-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="c4a8f-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c4a8f-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="c4a8f-129">Cela s’applique uniquement à plusieurs classeurs Excel ouverts sous Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="c4a8f-130">Co-édition</span><span class="sxs-lookup"><span data-stu-id="c4a8f-130">Coauthoring</span></span>

<span data-ttu-id="c4a8f-131">Consultez la rubrique [co-authoring in Excel Add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with Events in a CoAuthoring Environment.</span><span class="sxs-lookup"><span data-stu-id="c4a8f-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="c4a8f-132">L’article aborde également les conflits de fusion potentiels lors de l’utilisation de certaines API, telles que [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="c4a8f-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="c4a8f-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c4a8f-133">See also</span></span>

- [<span data-ttu-id="c4a8f-134">Résoudre les erreurs de développement avec les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c4a8f-134">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="c4a8f-135">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="c4a8f-135">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
