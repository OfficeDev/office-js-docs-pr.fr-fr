---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7258195e7ec0ef2432723d0f32f3d9ef1a3acf2b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268683"
---
# <a name="userprofile"></a><span data-ttu-id="834c6-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="834c6-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="834c6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="834c6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="834c6-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="834c6-104">Requirements</span></span>

|<span data-ttu-id="834c6-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="834c6-105">Requirement</span></span>| <span data-ttu-id="834c6-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="834c6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="834c6-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="834c6-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="834c6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="834c6-108">1.0</span></span>|
|[<span data-ttu-id="834c6-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="834c6-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="834c6-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="834c6-110">ReadItem</span></span>|
|[<span data-ttu-id="834c6-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="834c6-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="834c6-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="834c6-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="834c6-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="834c6-113">Members and methods</span></span>

| <span data-ttu-id="834c6-114">Membre</span><span class="sxs-lookup"><span data-stu-id="834c6-114">Member</span></span> | <span data-ttu-id="834c6-115">Type</span><span class="sxs-lookup"><span data-stu-id="834c6-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="834c6-116">displayName</span><span class="sxs-lookup"><span data-stu-id="834c6-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="834c6-117">Member</span><span class="sxs-lookup"><span data-stu-id="834c6-117">Member</span></span> |
| [<span data-ttu-id="834c6-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="834c6-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="834c6-119">Member</span><span class="sxs-lookup"><span data-stu-id="834c6-119">Member</span></span> |
| [<span data-ttu-id="834c6-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="834c6-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="834c6-121">Membre</span><span class="sxs-lookup"><span data-stu-id="834c6-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="834c6-122">Membres</span><span class="sxs-lookup"><span data-stu-id="834c6-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="834c6-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="834c6-123">displayName: String</span></span>

<span data-ttu-id="834c6-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="834c6-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="834c6-125">Type</span><span class="sxs-lookup"><span data-stu-id="834c6-125">Type</span></span>

*   <span data-ttu-id="834c6-126">String</span><span class="sxs-lookup"><span data-stu-id="834c6-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="834c6-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="834c6-127">Requirements</span></span>

|<span data-ttu-id="834c6-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="834c6-128">Requirement</span></span>| <span data-ttu-id="834c6-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="834c6-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="834c6-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="834c6-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="834c6-131">1.0</span><span class="sxs-lookup"><span data-stu-id="834c6-131">1.0</span></span>|
|[<span data-ttu-id="834c6-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="834c6-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="834c6-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="834c6-133">ReadItem</span></span>|
|[<span data-ttu-id="834c6-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="834c6-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="834c6-135">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="834c6-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="834c6-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="834c6-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="834c6-137">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="834c6-137">emailAddress: String</span></span>

<span data-ttu-id="834c6-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="834c6-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="834c6-139">Type</span><span class="sxs-lookup"><span data-stu-id="834c6-139">Type</span></span>

*   <span data-ttu-id="834c6-140">String</span><span class="sxs-lookup"><span data-stu-id="834c6-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="834c6-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="834c6-141">Requirements</span></span>

|<span data-ttu-id="834c6-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="834c6-142">Requirement</span></span>| <span data-ttu-id="834c6-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="834c6-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="834c6-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="834c6-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="834c6-145">1.0</span><span class="sxs-lookup"><span data-stu-id="834c6-145">1.0</span></span>|
|[<span data-ttu-id="834c6-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="834c6-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="834c6-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="834c6-147">ReadItem</span></span>|
|[<span data-ttu-id="834c6-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="834c6-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="834c6-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="834c6-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="834c6-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="834c6-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="834c6-151">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="834c6-151">timeZone: String</span></span>

<span data-ttu-id="834c6-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="834c6-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="834c6-153">Type</span><span class="sxs-lookup"><span data-stu-id="834c6-153">Type</span></span>

*   <span data-ttu-id="834c6-154">String</span><span class="sxs-lookup"><span data-stu-id="834c6-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="834c6-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="834c6-155">Requirements</span></span>

|<span data-ttu-id="834c6-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="834c6-156">Requirement</span></span>| <span data-ttu-id="834c6-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="834c6-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="834c6-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="834c6-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="834c6-159">1.0</span><span class="sxs-lookup"><span data-stu-id="834c6-159">1.0</span></span>|
|[<span data-ttu-id="834c6-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="834c6-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="834c6-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="834c6-161">ReadItem</span></span>|
|[<span data-ttu-id="834c6-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="834c6-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="834c6-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="834c6-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="834c6-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="834c6-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
