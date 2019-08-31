---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8ba2a21b16c51c827155d793241b80c5c510dd5a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696336"
---
# <a name="userprofile"></a><span data-ttu-id="60b9c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="60b9c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="60b9c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="60b9c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="60b9c-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="60b9c-104">Requirements</span></span>

|<span data-ttu-id="60b9c-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="60b9c-105">Requirement</span></span>| <span data-ttu-id="60b9c-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="60b9c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="60b9c-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="60b9c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60b9c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="60b9c-108">1.0</span></span>|
|[<span data-ttu-id="60b9c-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="60b9c-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60b9c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60b9c-110">ReadItem</span></span>|
|[<span data-ttu-id="60b9c-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="60b9c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60b9c-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="60b9c-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="60b9c-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="60b9c-113">Members and methods</span></span>

| <span data-ttu-id="60b9c-114">Membre</span><span class="sxs-lookup"><span data-stu-id="60b9c-114">Member</span></span> | <span data-ttu-id="60b9c-115">Type</span><span class="sxs-lookup"><span data-stu-id="60b9c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="60b9c-116">displayName</span><span class="sxs-lookup"><span data-stu-id="60b9c-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="60b9c-117">Member</span><span class="sxs-lookup"><span data-stu-id="60b9c-117">Member</span></span> |
| [<span data-ttu-id="60b9c-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="60b9c-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="60b9c-119">Member</span><span class="sxs-lookup"><span data-stu-id="60b9c-119">Member</span></span> |
| [<span data-ttu-id="60b9c-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="60b9c-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="60b9c-121">Membre</span><span class="sxs-lookup"><span data-stu-id="60b9c-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="60b9c-122">Membres</span><span class="sxs-lookup"><span data-stu-id="60b9c-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="60b9c-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="60b9c-123">displayName: String</span></span>

<span data-ttu-id="60b9c-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="60b9c-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="60b9c-125">Type</span><span class="sxs-lookup"><span data-stu-id="60b9c-125">Type</span></span>

*   <span data-ttu-id="60b9c-126">String</span><span class="sxs-lookup"><span data-stu-id="60b9c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60b9c-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="60b9c-127">Requirements</span></span>

|<span data-ttu-id="60b9c-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="60b9c-128">Requirement</span></span>| <span data-ttu-id="60b9c-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="60b9c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="60b9c-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="60b9c-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60b9c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="60b9c-131">1.0</span></span>|
|[<span data-ttu-id="60b9c-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="60b9c-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60b9c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60b9c-133">ReadItem</span></span>|
|[<span data-ttu-id="60b9c-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="60b9c-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60b9c-135">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="60b9c-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60b9c-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="60b9c-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="60b9c-137">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="60b9c-137">emailAddress: String</span></span>

<span data-ttu-id="60b9c-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="60b9c-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="60b9c-139">Type</span><span class="sxs-lookup"><span data-stu-id="60b9c-139">Type</span></span>

*   <span data-ttu-id="60b9c-140">String</span><span class="sxs-lookup"><span data-stu-id="60b9c-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60b9c-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="60b9c-141">Requirements</span></span>

|<span data-ttu-id="60b9c-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="60b9c-142">Requirement</span></span>| <span data-ttu-id="60b9c-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="60b9c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="60b9c-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="60b9c-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60b9c-145">1.0</span><span class="sxs-lookup"><span data-stu-id="60b9c-145">1.0</span></span>|
|[<span data-ttu-id="60b9c-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="60b9c-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60b9c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60b9c-147">ReadItem</span></span>|
|[<span data-ttu-id="60b9c-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="60b9c-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60b9c-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="60b9c-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60b9c-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="60b9c-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="60b9c-151">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="60b9c-151">timeZone: String</span></span>

<span data-ttu-id="60b9c-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="60b9c-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="60b9c-153">Type</span><span class="sxs-lookup"><span data-stu-id="60b9c-153">Type</span></span>

*   <span data-ttu-id="60b9c-154">String</span><span class="sxs-lookup"><span data-stu-id="60b9c-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60b9c-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="60b9c-155">Requirements</span></span>

|<span data-ttu-id="60b9c-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="60b9c-156">Requirement</span></span>| <span data-ttu-id="60b9c-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="60b9c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="60b9c-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="60b9c-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60b9c-159">1.0</span><span class="sxs-lookup"><span data-stu-id="60b9c-159">1.0</span></span>|
|[<span data-ttu-id="60b9c-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="60b9c-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60b9c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60b9c-161">ReadItem</span></span>|
|[<span data-ttu-id="60b9c-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="60b9c-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60b9c-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="60b9c-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60b9c-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="60b9c-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
