---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 06492623e0b9ab16792d6b23dfaeb27d99125ff1
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696399"
---
# <a name="userprofile"></a><span data-ttu-id="e43ba-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e43ba-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e43ba-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e43ba-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e43ba-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e43ba-104">Requirements</span></span>

|<span data-ttu-id="e43ba-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e43ba-105">Requirement</span></span>| <span data-ttu-id="e43ba-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="e43ba-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e43ba-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e43ba-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e43ba-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e43ba-108">1.0</span></span>|
|[<span data-ttu-id="e43ba-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e43ba-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e43ba-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e43ba-110">ReadItem</span></span>|
|[<span data-ttu-id="e43ba-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e43ba-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e43ba-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e43ba-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e43ba-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="e43ba-113">Members and methods</span></span>

| <span data-ttu-id="e43ba-114">Membre</span><span class="sxs-lookup"><span data-stu-id="e43ba-114">Member</span></span> | <span data-ttu-id="e43ba-115">Type</span><span class="sxs-lookup"><span data-stu-id="e43ba-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e43ba-116">displayName</span><span class="sxs-lookup"><span data-stu-id="e43ba-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="e43ba-117">Member</span><span class="sxs-lookup"><span data-stu-id="e43ba-117">Member</span></span> |
| [<span data-ttu-id="e43ba-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e43ba-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e43ba-119">Member</span><span class="sxs-lookup"><span data-stu-id="e43ba-119">Member</span></span> |
| [<span data-ttu-id="e43ba-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="e43ba-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e43ba-121">Membre</span><span class="sxs-lookup"><span data-stu-id="e43ba-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e43ba-122">Membres</span><span class="sxs-lookup"><span data-stu-id="e43ba-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="e43ba-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="e43ba-123">displayName: String</span></span>

<span data-ttu-id="e43ba-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e43ba-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e43ba-125">Type</span><span class="sxs-lookup"><span data-stu-id="e43ba-125">Type</span></span>

*   <span data-ttu-id="e43ba-126">String</span><span class="sxs-lookup"><span data-stu-id="e43ba-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e43ba-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e43ba-127">Requirements</span></span>

|<span data-ttu-id="e43ba-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e43ba-128">Requirement</span></span>| <span data-ttu-id="e43ba-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="e43ba-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="e43ba-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e43ba-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e43ba-131">1.0</span><span class="sxs-lookup"><span data-stu-id="e43ba-131">1.0</span></span>|
|[<span data-ttu-id="e43ba-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e43ba-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e43ba-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e43ba-133">ReadItem</span></span>|
|[<span data-ttu-id="e43ba-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e43ba-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e43ba-135">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e43ba-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e43ba-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="e43ba-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="e43ba-137">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="e43ba-137">emailAddress: String</span></span>

<span data-ttu-id="e43ba-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e43ba-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e43ba-139">Type</span><span class="sxs-lookup"><span data-stu-id="e43ba-139">Type</span></span>

*   <span data-ttu-id="e43ba-140">String</span><span class="sxs-lookup"><span data-stu-id="e43ba-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e43ba-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e43ba-141">Requirements</span></span>

|<span data-ttu-id="e43ba-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e43ba-142">Requirement</span></span>| <span data-ttu-id="e43ba-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="e43ba-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="e43ba-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e43ba-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e43ba-145">1.0</span><span class="sxs-lookup"><span data-stu-id="e43ba-145">1.0</span></span>|
|[<span data-ttu-id="e43ba-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e43ba-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e43ba-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e43ba-147">ReadItem</span></span>|
|[<span data-ttu-id="e43ba-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e43ba-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e43ba-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e43ba-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e43ba-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="e43ba-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="e43ba-151">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="e43ba-151">timeZone: String</span></span>

<span data-ttu-id="e43ba-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e43ba-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e43ba-153">Type</span><span class="sxs-lookup"><span data-stu-id="e43ba-153">Type</span></span>

*   <span data-ttu-id="e43ba-154">String</span><span class="sxs-lookup"><span data-stu-id="e43ba-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e43ba-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e43ba-155">Requirements</span></span>

|<span data-ttu-id="e43ba-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e43ba-156">Requirement</span></span>| <span data-ttu-id="e43ba-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="e43ba-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e43ba-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e43ba-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e43ba-159">1.0</span><span class="sxs-lookup"><span data-stu-id="e43ba-159">1.0</span></span>|
|[<span data-ttu-id="e43ba-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e43ba-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e43ba-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e43ba-161">ReadItem</span></span>|
|[<span data-ttu-id="e43ba-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e43ba-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e43ba-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e43ba-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e43ba-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="e43ba-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
