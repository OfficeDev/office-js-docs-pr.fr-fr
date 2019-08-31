---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 993fad674fcc616483ac927619e7ca64d81b7326
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696091"
---
# <a name="userprofile"></a><span data-ttu-id="83699-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="83699-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="83699-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="83699-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="83699-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="83699-104">Requirements</span></span>

|<span data-ttu-id="83699-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="83699-105">Requirement</span></span>| <span data-ttu-id="83699-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="83699-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="83699-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="83699-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83699-108">1.0</span><span class="sxs-lookup"><span data-stu-id="83699-108">1.0</span></span>|
|[<span data-ttu-id="83699-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="83699-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83699-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83699-110">ReadItem</span></span>|
|[<span data-ttu-id="83699-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="83699-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="83699-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="83699-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="83699-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="83699-113">Members and methods</span></span>

| <span data-ttu-id="83699-114">Membre</span><span class="sxs-lookup"><span data-stu-id="83699-114">Member</span></span> | <span data-ttu-id="83699-115">Type</span><span class="sxs-lookup"><span data-stu-id="83699-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="83699-116">displayName</span><span class="sxs-lookup"><span data-stu-id="83699-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="83699-117">Member</span><span class="sxs-lookup"><span data-stu-id="83699-117">Member</span></span> |
| [<span data-ttu-id="83699-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="83699-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="83699-119">Member</span><span class="sxs-lookup"><span data-stu-id="83699-119">Member</span></span> |
| [<span data-ttu-id="83699-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="83699-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="83699-121">Membre</span><span class="sxs-lookup"><span data-stu-id="83699-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="83699-122">Membres</span><span class="sxs-lookup"><span data-stu-id="83699-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="83699-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="83699-123">displayName: String</span></span>

<span data-ttu-id="83699-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="83699-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="83699-125">Type</span><span class="sxs-lookup"><span data-stu-id="83699-125">Type</span></span>

*   <span data-ttu-id="83699-126">String</span><span class="sxs-lookup"><span data-stu-id="83699-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83699-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="83699-127">Requirements</span></span>

|<span data-ttu-id="83699-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="83699-128">Requirement</span></span>| <span data-ttu-id="83699-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="83699-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="83699-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="83699-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83699-131">1.0</span><span class="sxs-lookup"><span data-stu-id="83699-131">1.0</span></span>|
|[<span data-ttu-id="83699-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="83699-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83699-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83699-133">ReadItem</span></span>|
|[<span data-ttu-id="83699-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="83699-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="83699-135">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="83699-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83699-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="83699-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="83699-137">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="83699-137">emailAddress: String</span></span>

<span data-ttu-id="83699-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="83699-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="83699-139">Type</span><span class="sxs-lookup"><span data-stu-id="83699-139">Type</span></span>

*   <span data-ttu-id="83699-140">String</span><span class="sxs-lookup"><span data-stu-id="83699-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83699-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="83699-141">Requirements</span></span>

|<span data-ttu-id="83699-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="83699-142">Requirement</span></span>| <span data-ttu-id="83699-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="83699-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="83699-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="83699-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83699-145">1.0</span><span class="sxs-lookup"><span data-stu-id="83699-145">1.0</span></span>|
|[<span data-ttu-id="83699-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="83699-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83699-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83699-147">ReadItem</span></span>|
|[<span data-ttu-id="83699-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="83699-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="83699-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="83699-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83699-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="83699-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="83699-151">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="83699-151">timeZone: String</span></span>

<span data-ttu-id="83699-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="83699-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="83699-153">Type</span><span class="sxs-lookup"><span data-stu-id="83699-153">Type</span></span>

*   <span data-ttu-id="83699-154">String</span><span class="sxs-lookup"><span data-stu-id="83699-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83699-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="83699-155">Requirements</span></span>

|<span data-ttu-id="83699-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="83699-156">Requirement</span></span>| <span data-ttu-id="83699-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="83699-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="83699-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="83699-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83699-159">1.0</span><span class="sxs-lookup"><span data-stu-id="83699-159">1.0</span></span>|
|[<span data-ttu-id="83699-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="83699-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83699-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83699-161">ReadItem</span></span>|
|[<span data-ttu-id="83699-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="83699-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="83699-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="83699-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83699-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="83699-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
