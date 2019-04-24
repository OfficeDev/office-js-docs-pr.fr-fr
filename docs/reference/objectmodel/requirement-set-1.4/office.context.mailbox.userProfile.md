---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450316"
---
# <a name="userprofile"></a><span data-ttu-id="547da-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="547da-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="547da-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="547da-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="547da-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="547da-104">Requirements</span></span>

|<span data-ttu-id="547da-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="547da-105">Requirement</span></span>| <span data-ttu-id="547da-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="547da-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="547da-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="547da-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="547da-108">1.0</span><span class="sxs-lookup"><span data-stu-id="547da-108">1.0</span></span>|
|[<span data-ttu-id="547da-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="547da-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="547da-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="547da-110">ReadItem</span></span>|
|[<span data-ttu-id="547da-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="547da-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="547da-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="547da-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="547da-113">Membres</span><span class="sxs-lookup"><span data-stu-id="547da-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="547da-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="547da-114">displayName :String</span></span>

<span data-ttu-id="547da-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="547da-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="547da-116">Type</span><span class="sxs-lookup"><span data-stu-id="547da-116">Type</span></span>

*   <span data-ttu-id="547da-117">String</span><span class="sxs-lookup"><span data-stu-id="547da-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="547da-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="547da-118">Requirements</span></span>

|<span data-ttu-id="547da-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="547da-119">Requirement</span></span>| <span data-ttu-id="547da-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="547da-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="547da-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="547da-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="547da-122">1.0</span><span class="sxs-lookup"><span data-stu-id="547da-122">1.0</span></span>|
|[<span data-ttu-id="547da-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="547da-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="547da-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="547da-124">ReadItem</span></span>|
|[<span data-ttu-id="547da-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="547da-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="547da-126">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="547da-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="547da-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="547da-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="547da-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="547da-128">emailAddress :String</span></span>

<span data-ttu-id="547da-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="547da-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="547da-130">Type</span><span class="sxs-lookup"><span data-stu-id="547da-130">Type</span></span>

*   <span data-ttu-id="547da-131">String</span><span class="sxs-lookup"><span data-stu-id="547da-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="547da-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="547da-132">Requirements</span></span>

|<span data-ttu-id="547da-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="547da-133">Requirement</span></span>| <span data-ttu-id="547da-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="547da-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="547da-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="547da-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="547da-136">1.0</span><span class="sxs-lookup"><span data-stu-id="547da-136">1.0</span></span>|
|[<span data-ttu-id="547da-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="547da-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="547da-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="547da-138">ReadItem</span></span>|
|[<span data-ttu-id="547da-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="547da-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="547da-140">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="547da-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="547da-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="547da-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="547da-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="547da-142">timeZone :String</span></span>

<span data-ttu-id="547da-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="547da-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="547da-144">Type</span><span class="sxs-lookup"><span data-stu-id="547da-144">Type</span></span>

*   <span data-ttu-id="547da-145">String</span><span class="sxs-lookup"><span data-stu-id="547da-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="547da-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="547da-146">Requirements</span></span>

|<span data-ttu-id="547da-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="547da-147">Requirement</span></span>| <span data-ttu-id="547da-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="547da-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="547da-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="547da-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="547da-150">1.0</span><span class="sxs-lookup"><span data-stu-id="547da-150">1.0</span></span>|
|[<span data-ttu-id="547da-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="547da-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="547da-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="547da-152">ReadItem</span></span>|
|[<span data-ttu-id="547da-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="547da-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="547da-154">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="547da-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="547da-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="547da-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
