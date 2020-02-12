---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950991"
---
# <a name="userprofile"></a><span data-ttu-id="78e84-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="78e84-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="78e84-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="78e84-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="78e84-104">Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="78e84-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78e84-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="78e84-105">Requirements</span></span>

|<span data-ttu-id="78e84-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="78e84-106">Requirement</span></span>| <span data-ttu-id="78e84-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="78e84-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="78e84-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="78e84-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="78e84-109">1.1</span><span class="sxs-lookup"><span data-stu-id="78e84-109">1.1</span></span>|
|[<span data-ttu-id="78e84-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="78e84-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="78e84-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78e84-111">ReadItem</span></span>|
|[<span data-ttu-id="78e84-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="78e84-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="78e84-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="78e84-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="78e84-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="78e84-114">Properties</span></span>

| <span data-ttu-id="78e84-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="78e84-115">Property</span></span> | <span data-ttu-id="78e84-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="78e84-116">Minimum</span></span><br><span data-ttu-id="78e84-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="78e84-117">permission level</span></span> | <span data-ttu-id="78e84-118">Modes</span><span class="sxs-lookup"><span data-stu-id="78e84-118">Modes</span></span> | <span data-ttu-id="78e84-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="78e84-119">Return type</span></span> | <span data-ttu-id="78e84-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="78e84-120">Minimum</span></span><br><span data-ttu-id="78e84-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="78e84-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="78e84-122">displayName</span><span class="sxs-lookup"><span data-stu-id="78e84-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="78e84-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78e84-123">ReadItem</span></span> | <span data-ttu-id="78e84-124">Composition</span><span class="sxs-lookup"><span data-stu-id="78e84-124">Compose</span></span><br><span data-ttu-id="78e84-125">Lire</span><span class="sxs-lookup"><span data-stu-id="78e84-125">Read</span></span> | <span data-ttu-id="78e84-126">String</span><span class="sxs-lookup"><span data-stu-id="78e84-126">String</span></span> | [<span data-ttu-id="78e84-127">1.1</span><span class="sxs-lookup"><span data-stu-id="78e84-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="78e84-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="78e84-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="78e84-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78e84-129">ReadItem</span></span> | <span data-ttu-id="78e84-130">Composition</span><span class="sxs-lookup"><span data-stu-id="78e84-130">Compose</span></span><br><span data-ttu-id="78e84-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="78e84-131">Read</span></span> | <span data-ttu-id="78e84-132">String</span><span class="sxs-lookup"><span data-stu-id="78e84-132">String</span></span> | [<span data-ttu-id="78e84-133">1.1</span><span class="sxs-lookup"><span data-stu-id="78e84-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="78e84-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="78e84-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="78e84-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78e84-135">ReadItem</span></span> | <span data-ttu-id="78e84-136">Composition</span><span class="sxs-lookup"><span data-stu-id="78e84-136">Compose</span></span><br><span data-ttu-id="78e84-137">Lire</span><span class="sxs-lookup"><span data-stu-id="78e84-137">Read</span></span> | <span data-ttu-id="78e84-138">String</span><span class="sxs-lookup"><span data-stu-id="78e84-138">String</span></span> | [<span data-ttu-id="78e84-139">1.1</span><span class="sxs-lookup"><span data-stu-id="78e84-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
