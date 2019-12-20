---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 5ceafe65dedcb1db6c67ca28f9a1d9e05f805850
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814284"
---
# <a name="diagnostics"></a><span data-ttu-id="74c27-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="74c27-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="74c27-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="74c27-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="74c27-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="74c27-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="74c27-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="74c27-105">Requirements</span></span>

|<span data-ttu-id="74c27-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="74c27-106">Requirement</span></span>| <span data-ttu-id="74c27-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="74c27-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="74c27-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="74c27-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="74c27-109">1.1</span><span class="sxs-lookup"><span data-stu-id="74c27-109">1.1</span></span>|
|[<span data-ttu-id="74c27-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="74c27-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="74c27-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="74c27-111">ReadItem</span></span>|
|[<span data-ttu-id="74c27-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="74c27-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="74c27-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="74c27-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="74c27-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="74c27-114">Properties</span></span>

| <span data-ttu-id="74c27-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="74c27-115">Property</span></span> | <span data-ttu-id="74c27-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="74c27-116">Minimum</span></span><br><span data-ttu-id="74c27-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="74c27-117">permission level</span></span> | <span data-ttu-id="74c27-118">Modes</span><span class="sxs-lookup"><span data-stu-id="74c27-118">Modes</span></span> | <span data-ttu-id="74c27-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="74c27-119">Return type</span></span> | <span data-ttu-id="74c27-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="74c27-120">Minimum</span></span><br><span data-ttu-id="74c27-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="74c27-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="74c27-122">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="74c27-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4#hostname) | <span data-ttu-id="74c27-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="74c27-123">ReadItem</span></span> | <span data-ttu-id="74c27-124">Composition</span><span class="sxs-lookup"><span data-stu-id="74c27-124">Compose</span></span><br><span data-ttu-id="74c27-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="74c27-125">Read</span></span> | <span data-ttu-id="74c27-126">String</span><span class="sxs-lookup"><span data-stu-id="74c27-126">String</span></span> | [<span data-ttu-id="74c27-127">1.1</span><span class="sxs-lookup"><span data-stu-id="74c27-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="74c27-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="74c27-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4#hostversion) | <span data-ttu-id="74c27-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="74c27-129">ReadItem</span></span> | <span data-ttu-id="74c27-130">Composition</span><span class="sxs-lookup"><span data-stu-id="74c27-130">Compose</span></span><br><span data-ttu-id="74c27-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="74c27-131">Read</span></span> | <span data-ttu-id="74c27-132">String</span><span class="sxs-lookup"><span data-stu-id="74c27-132">String</span></span> | [<span data-ttu-id="74c27-133">1.1</span><span class="sxs-lookup"><span data-stu-id="74c27-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="74c27-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="74c27-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4#owaview) | <span data-ttu-id="74c27-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="74c27-135">ReadItem</span></span> | <span data-ttu-id="74c27-136">Composition</span><span class="sxs-lookup"><span data-stu-id="74c27-136">Compose</span></span><br><span data-ttu-id="74c27-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="74c27-137">Read</span></span> | <span data-ttu-id="74c27-138">String</span><span class="sxs-lookup"><span data-stu-id="74c27-138">String</span></span> | [<span data-ttu-id="74c27-139">1.1</span><span class="sxs-lookup"><span data-stu-id="74c27-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
