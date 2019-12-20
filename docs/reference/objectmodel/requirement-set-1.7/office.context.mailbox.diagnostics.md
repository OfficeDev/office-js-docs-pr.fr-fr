---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3baf192dc209d015ff888ff5067d2cafbaee3181
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814625"
---
# <a name="diagnostics"></a><span data-ttu-id="e3f67-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="e3f67-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="e3f67-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="e3f67-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="e3f67-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3f67-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3f67-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3f67-105">Requirements</span></span>

|<span data-ttu-id="e3f67-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3f67-106">Requirement</span></span>| <span data-ttu-id="e3f67-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3f67-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f67-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3f67-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e3f67-109">1.1</span><span class="sxs-lookup"><span data-stu-id="e3f67-109">1.1</span></span>|
|[<span data-ttu-id="e3f67-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3f67-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f67-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f67-111">ReadItem</span></span>|
|[<span data-ttu-id="e3f67-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3f67-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f67-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3f67-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="e3f67-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e3f67-114">Properties</span></span>

| <span data-ttu-id="e3f67-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="e3f67-115">Property</span></span> | <span data-ttu-id="e3f67-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="e3f67-116">Minimum</span></span><br><span data-ttu-id="e3f67-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="e3f67-117">permission level</span></span> | <span data-ttu-id="e3f67-118">Modes</span><span class="sxs-lookup"><span data-stu-id="e3f67-118">Modes</span></span> | <span data-ttu-id="e3f67-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e3f67-119">Return type</span></span> | <span data-ttu-id="e3f67-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="e3f67-120">Minimum</span></span><br><span data-ttu-id="e3f67-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3f67-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="e3f67-122">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="e3f67-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7#hostname) | <span data-ttu-id="e3f67-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f67-123">ReadItem</span></span> | <span data-ttu-id="e3f67-124">Composition</span><span class="sxs-lookup"><span data-stu-id="e3f67-124">Compose</span></span><br><span data-ttu-id="e3f67-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="e3f67-125">Read</span></span> | <span data-ttu-id="e3f67-126">String</span><span class="sxs-lookup"><span data-stu-id="e3f67-126">String</span></span> | [<span data-ttu-id="e3f67-127">1.1</span><span class="sxs-lookup"><span data-stu-id="e3f67-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e3f67-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="e3f67-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7#hostversion) | <span data-ttu-id="e3f67-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f67-129">ReadItem</span></span> | <span data-ttu-id="e3f67-130">Composition</span><span class="sxs-lookup"><span data-stu-id="e3f67-130">Compose</span></span><br><span data-ttu-id="e3f67-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="e3f67-131">Read</span></span> | <span data-ttu-id="e3f67-132">String</span><span class="sxs-lookup"><span data-stu-id="e3f67-132">String</span></span> | [<span data-ttu-id="e3f67-133">1.1</span><span class="sxs-lookup"><span data-stu-id="e3f67-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e3f67-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="e3f67-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7#owaview) | <span data-ttu-id="e3f67-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f67-135">ReadItem</span></span> | <span data-ttu-id="e3f67-136">Composition</span><span class="sxs-lookup"><span data-stu-id="e3f67-136">Compose</span></span><br><span data-ttu-id="e3f67-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="e3f67-137">Read</span></span> | <span data-ttu-id="e3f67-138">String</span><span class="sxs-lookup"><span data-stu-id="e3f67-138">String</span></span> | [<span data-ttu-id="e3f67-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e3f67-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
