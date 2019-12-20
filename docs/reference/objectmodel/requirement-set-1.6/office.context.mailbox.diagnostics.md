---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ee10e511ed81a591e5e7b89c7650e16fca27da09
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814737"
---
# <a name="diagnostics"></a><span data-ttu-id="07a60-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="07a60-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="07a60-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="07a60-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="07a60-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="07a60-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07a60-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="07a60-105">Requirements</span></span>

|<span data-ttu-id="07a60-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="07a60-106">Requirement</span></span>| <span data-ttu-id="07a60-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="07a60-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="07a60-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="07a60-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="07a60-109">1.1</span><span class="sxs-lookup"><span data-stu-id="07a60-109">1.1</span></span>|
|[<span data-ttu-id="07a60-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="07a60-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07a60-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07a60-111">ReadItem</span></span>|
|[<span data-ttu-id="07a60-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="07a60-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07a60-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="07a60-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="07a60-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="07a60-114">Properties</span></span>

| <span data-ttu-id="07a60-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="07a60-115">Property</span></span> | <span data-ttu-id="07a60-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="07a60-116">Minimum</span></span><br><span data-ttu-id="07a60-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="07a60-117">permission level</span></span> | <span data-ttu-id="07a60-118">Modes</span><span class="sxs-lookup"><span data-stu-id="07a60-118">Modes</span></span> | <span data-ttu-id="07a60-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="07a60-119">Return type</span></span> | <span data-ttu-id="07a60-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="07a60-120">Minimum</span></span><br><span data-ttu-id="07a60-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="07a60-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="07a60-122">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="07a60-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6#hostname) | <span data-ttu-id="07a60-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07a60-123">ReadItem</span></span> | <span data-ttu-id="07a60-124">Composition</span><span class="sxs-lookup"><span data-stu-id="07a60-124">Compose</span></span><br><span data-ttu-id="07a60-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="07a60-125">Read</span></span> | <span data-ttu-id="07a60-126">String</span><span class="sxs-lookup"><span data-stu-id="07a60-126">String</span></span> | [<span data-ttu-id="07a60-127">1.1</span><span class="sxs-lookup"><span data-stu-id="07a60-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="07a60-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="07a60-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6#hostversion) | <span data-ttu-id="07a60-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07a60-129">ReadItem</span></span> | <span data-ttu-id="07a60-130">Composition</span><span class="sxs-lookup"><span data-stu-id="07a60-130">Compose</span></span><br><span data-ttu-id="07a60-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="07a60-131">Read</span></span> | <span data-ttu-id="07a60-132">String</span><span class="sxs-lookup"><span data-stu-id="07a60-132">String</span></span> | [<span data-ttu-id="07a60-133">1.1</span><span class="sxs-lookup"><span data-stu-id="07a60-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="07a60-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="07a60-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6#owaview) | <span data-ttu-id="07a60-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07a60-135">ReadItem</span></span> | <span data-ttu-id="07a60-136">Composition</span><span class="sxs-lookup"><span data-stu-id="07a60-136">Compose</span></span><br><span data-ttu-id="07a60-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="07a60-137">Read</span></span> | <span data-ttu-id="07a60-138">String</span><span class="sxs-lookup"><span data-stu-id="07a60-138">String</span></span> | [<span data-ttu-id="07a60-139">1.1</span><span class="sxs-lookup"><span data-stu-id="07a60-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
