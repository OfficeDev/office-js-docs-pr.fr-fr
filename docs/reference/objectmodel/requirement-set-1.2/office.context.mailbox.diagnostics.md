---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dad9d35c397351938944d89bf98e450427cb74a3
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814982"
---
# <a name="diagnostics"></a><span data-ttu-id="b9346-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="b9346-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="b9346-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="b9346-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="b9346-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9346-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9346-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9346-105">Requirements</span></span>

|<span data-ttu-id="b9346-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9346-106">Requirement</span></span>| <span data-ttu-id="b9346-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9346-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9346-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9346-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9346-109">1.1</span><span class="sxs-lookup"><span data-stu-id="b9346-109">1.1</span></span>|
|[<span data-ttu-id="b9346-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9346-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9346-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9346-111">ReadItem</span></span>|
|[<span data-ttu-id="b9346-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9346-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9346-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9346-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="b9346-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b9346-114">Properties</span></span>

| <span data-ttu-id="b9346-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="b9346-115">Property</span></span> | <span data-ttu-id="b9346-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="b9346-116">Minimum</span></span><br><span data-ttu-id="b9346-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="b9346-117">permission level</span></span> | <span data-ttu-id="b9346-118">Modes</span><span class="sxs-lookup"><span data-stu-id="b9346-118">Modes</span></span> | <span data-ttu-id="b9346-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="b9346-119">Return type</span></span> | <span data-ttu-id="b9346-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="b9346-120">Minimum</span></span><br><span data-ttu-id="b9346-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9346-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="b9346-122">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="b9346-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#hostname) | <span data-ttu-id="b9346-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9346-123">ReadItem</span></span> | <span data-ttu-id="b9346-124">Composition</span><span class="sxs-lookup"><span data-stu-id="b9346-124">Compose</span></span><br><span data-ttu-id="b9346-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9346-125">Read</span></span> | <span data-ttu-id="b9346-126">String</span><span class="sxs-lookup"><span data-stu-id="b9346-126">String</span></span> | [<span data-ttu-id="b9346-127">1.1</span><span class="sxs-lookup"><span data-stu-id="b9346-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9346-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="b9346-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#hostversion) | <span data-ttu-id="b9346-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9346-129">ReadItem</span></span> | <span data-ttu-id="b9346-130">Composition</span><span class="sxs-lookup"><span data-stu-id="b9346-130">Compose</span></span><br><span data-ttu-id="b9346-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9346-131">Read</span></span> | <span data-ttu-id="b9346-132">String</span><span class="sxs-lookup"><span data-stu-id="b9346-132">String</span></span> | [<span data-ttu-id="b9346-133">1.1</span><span class="sxs-lookup"><span data-stu-id="b9346-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9346-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="b9346-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#owaview) | <span data-ttu-id="b9346-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9346-135">ReadItem</span></span> | <span data-ttu-id="b9346-136">Composition</span><span class="sxs-lookup"><span data-stu-id="b9346-136">Compose</span></span><br><span data-ttu-id="b9346-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="b9346-137">Read</span></span> | <span data-ttu-id="b9346-138">String</span><span class="sxs-lookup"><span data-stu-id="b9346-138">String</span></span> | [<span data-ttu-id="b9346-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b9346-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
