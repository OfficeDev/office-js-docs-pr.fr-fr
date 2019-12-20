---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7b7b9c7facd0542335094a42a3d1f53dab1f6aef
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814319"
---
# <a name="userprofile"></a><span data-ttu-id="5c56e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="5c56e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="5c56e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="5c56e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="5c56e-104">Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="5c56e-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5c56e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5c56e-105">Requirements</span></span>

|<span data-ttu-id="5c56e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5c56e-106">Requirement</span></span>| <span data-ttu-id="5c56e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="5c56e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5c56e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5c56e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5c56e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="5c56e-109">1.1</span></span>|
|[<span data-ttu-id="5c56e-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="5c56e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5c56e-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c56e-111">ReadItem</span></span>|
|[<span data-ttu-id="5c56e-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5c56e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5c56e-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5c56e-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="5c56e-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="5c56e-114">Properties</span></span>

| <span data-ttu-id="5c56e-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="5c56e-115">Property</span></span> | <span data-ttu-id="5c56e-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="5c56e-116">Minimum</span></span><br><span data-ttu-id="5c56e-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="5c56e-117">permission level</span></span> | <span data-ttu-id="5c56e-118">Modes</span><span class="sxs-lookup"><span data-stu-id="5c56e-118">Modes</span></span> | <span data-ttu-id="5c56e-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="5c56e-119">Return type</span></span> | <span data-ttu-id="5c56e-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="5c56e-120">Minimum</span></span><br><span data-ttu-id="5c56e-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="5c56e-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="5c56e-122">displayName</span><span class="sxs-lookup"><span data-stu-id="5c56e-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#displayname) | <span data-ttu-id="5c56e-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c56e-123">ReadItem</span></span> | <span data-ttu-id="5c56e-124">Composition</span><span class="sxs-lookup"><span data-stu-id="5c56e-124">Compose</span></span><br><span data-ttu-id="5c56e-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="5c56e-125">Read</span></span> | <span data-ttu-id="5c56e-126">String</span><span class="sxs-lookup"><span data-stu-id="5c56e-126">String</span></span> | [<span data-ttu-id="5c56e-127">1.1</span><span class="sxs-lookup"><span data-stu-id="5c56e-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5c56e-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5c56e-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#emailaddress) | <span data-ttu-id="5c56e-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c56e-129">ReadItem</span></span> | <span data-ttu-id="5c56e-130">Composition</span><span class="sxs-lookup"><span data-stu-id="5c56e-130">Compose</span></span><br><span data-ttu-id="5c56e-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="5c56e-131">Read</span></span> | <span data-ttu-id="5c56e-132">String</span><span class="sxs-lookup"><span data-stu-id="5c56e-132">String</span></span> | [<span data-ttu-id="5c56e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="5c56e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5c56e-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="5c56e-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#timezone) | <span data-ttu-id="5c56e-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5c56e-135">ReadItem</span></span> | <span data-ttu-id="5c56e-136">Composition</span><span class="sxs-lookup"><span data-stu-id="5c56e-136">Compose</span></span><br><span data-ttu-id="5c56e-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="5c56e-137">Read</span></span> | <span data-ttu-id="5c56e-138">String</span><span class="sxs-lookup"><span data-stu-id="5c56e-138">String</span></span> | [<span data-ttu-id="5c56e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="5c56e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
