---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 4284f5e2367e72700d1b34bbac18b08bb0bc77f8
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814835"
---
# <a name="diagnostics"></a><span data-ttu-id="1a9b3-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="1a9b3-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="1a9b3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="1a9b3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="1a9b3-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1a9b3-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9b3-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1a9b3-105">Requirements</span></span>

|<span data-ttu-id="1a9b3-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a9b3-106">Requirement</span></span>| <span data-ttu-id="1a9b3-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="1a9b3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9b3-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1a9b3-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1a9b3-109">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9b3-109">1.1</span></span>|
|[<span data-ttu-id="1a9b3-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1a9b3-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9b3-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9b3-111">ReadItem</span></span>|
|[<span data-ttu-id="1a9b3-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1a9b3-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9b3-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1a9b3-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="1a9b3-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="1a9b3-114">Properties</span></span>

| <span data-ttu-id="1a9b3-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="1a9b3-115">Property</span></span> | <span data-ttu-id="1a9b3-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="1a9b3-116">Minimum</span></span><br><span data-ttu-id="1a9b3-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="1a9b3-117">permission level</span></span> | <span data-ttu-id="1a9b3-118">Modes</span><span class="sxs-lookup"><span data-stu-id="1a9b3-118">Modes</span></span> | <span data-ttu-id="1a9b3-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="1a9b3-119">Return type</span></span> | <span data-ttu-id="1a9b3-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="1a9b3-120">Minimum</span></span><br><span data-ttu-id="1a9b3-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a9b3-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="1a9b3-122">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="1a9b3-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5#hostname) | <span data-ttu-id="1a9b3-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9b3-123">ReadItem</span></span> | <span data-ttu-id="1a9b3-124">Composition</span><span class="sxs-lookup"><span data-stu-id="1a9b3-124">Compose</span></span><br><span data-ttu-id="1a9b3-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a9b3-125">Read</span></span> | <span data-ttu-id="1a9b3-126">String</span><span class="sxs-lookup"><span data-stu-id="1a9b3-126">String</span></span> | [<span data-ttu-id="1a9b3-127">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9b3-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1a9b3-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="1a9b3-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5#hostversion) | <span data-ttu-id="1a9b3-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9b3-129">ReadItem</span></span> | <span data-ttu-id="1a9b3-130">Composition</span><span class="sxs-lookup"><span data-stu-id="1a9b3-130">Compose</span></span><br><span data-ttu-id="1a9b3-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a9b3-131">Read</span></span> | <span data-ttu-id="1a9b3-132">String</span><span class="sxs-lookup"><span data-stu-id="1a9b3-132">String</span></span> | [<span data-ttu-id="1a9b3-133">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9b3-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1a9b3-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="1a9b3-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5#owaview) | <span data-ttu-id="1a9b3-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9b3-135">ReadItem</span></span> | <span data-ttu-id="1a9b3-136">Composition</span><span class="sxs-lookup"><span data-stu-id="1a9b3-136">Compose</span></span><br><span data-ttu-id="1a9b3-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a9b3-137">Read</span></span> | <span data-ttu-id="1a9b3-138">String</span><span class="sxs-lookup"><span data-stu-id="1a9b3-138">String</span></span> | [<span data-ttu-id="1a9b3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9b3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
