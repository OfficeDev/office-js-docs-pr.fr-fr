---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 6b5229c1bc300d11714f3aa2cf8fa8ff2465667c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814263"
---
# <a name="userprofile"></a><span data-ttu-id="de43a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="de43a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="de43a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="de43a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="de43a-104">Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="de43a-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="de43a-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="de43a-105">Requirements</span></span>

|<span data-ttu-id="de43a-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="de43a-106">Requirement</span></span>| <span data-ttu-id="de43a-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="de43a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="de43a-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="de43a-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="de43a-109">1.1</span><span class="sxs-lookup"><span data-stu-id="de43a-109">1.1</span></span>|
|[<span data-ttu-id="de43a-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="de43a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de43a-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de43a-111">ReadItem</span></span>|
|[<span data-ttu-id="de43a-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="de43a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de43a-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="de43a-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="de43a-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="de43a-114">Properties</span></span>

| <span data-ttu-id="de43a-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="de43a-115">Property</span></span> | <span data-ttu-id="de43a-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="de43a-116">Minimum</span></span><br><span data-ttu-id="de43a-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="de43a-117">permission level</span></span> | <span data-ttu-id="de43a-118">Modes</span><span class="sxs-lookup"><span data-stu-id="de43a-118">Modes</span></span> | <span data-ttu-id="de43a-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="de43a-119">Return type</span></span> | <span data-ttu-id="de43a-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="de43a-120">Minimum</span></span><br><span data-ttu-id="de43a-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="de43a-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="de43a-122">displayName</span><span class="sxs-lookup"><span data-stu-id="de43a-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="de43a-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de43a-123">ReadItem</span></span> | <span data-ttu-id="de43a-124">Composition</span><span class="sxs-lookup"><span data-stu-id="de43a-124">Compose</span></span><br><span data-ttu-id="de43a-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="de43a-125">Read</span></span> | <span data-ttu-id="de43a-126">String</span><span class="sxs-lookup"><span data-stu-id="de43a-126">String</span></span> | [<span data-ttu-id="de43a-127">1.1</span><span class="sxs-lookup"><span data-stu-id="de43a-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="de43a-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="de43a-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="de43a-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de43a-129">ReadItem</span></span> | <span data-ttu-id="de43a-130">Composition</span><span class="sxs-lookup"><span data-stu-id="de43a-130">Compose</span></span><br><span data-ttu-id="de43a-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="de43a-131">Read</span></span> | <span data-ttu-id="de43a-132">String</span><span class="sxs-lookup"><span data-stu-id="de43a-132">String</span></span> | [<span data-ttu-id="de43a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="de43a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="de43a-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="de43a-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="de43a-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de43a-135">ReadItem</span></span> | <span data-ttu-id="de43a-136">Composition</span><span class="sxs-lookup"><span data-stu-id="de43a-136">Compose</span></span><br><span data-ttu-id="de43a-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="de43a-137">Read</span></span> | <span data-ttu-id="de43a-138">String</span><span class="sxs-lookup"><span data-stu-id="de43a-138">String</span></span> | [<span data-ttu-id="de43a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="de43a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
