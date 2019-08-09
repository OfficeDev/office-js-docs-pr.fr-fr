---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8924d8b0dfa5bb43be8867cbd0e83ee01ff788cb
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268389"
---
# <a name="userprofile"></a><span data-ttu-id="4acb4-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="4acb4-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="4acb4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="4acb4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4acb4-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4acb4-104">Requirements</span></span>

|<span data-ttu-id="4acb4-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4acb4-105">Requirement</span></span>| <span data-ttu-id="4acb4-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="4acb4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4acb4-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4acb4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4acb4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4acb4-108">1.0</span></span>|
|[<span data-ttu-id="4acb4-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4acb4-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4acb4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4acb4-110">ReadItem</span></span>|
|[<span data-ttu-id="4acb4-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4acb4-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4acb4-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4acb4-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4acb4-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4acb4-113">Members and methods</span></span>

| <span data-ttu-id="4acb4-114">Membre</span><span class="sxs-lookup"><span data-stu-id="4acb4-114">Member</span></span> | <span data-ttu-id="4acb4-115">Type</span><span class="sxs-lookup"><span data-stu-id="4acb4-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4acb4-116">displayName</span><span class="sxs-lookup"><span data-stu-id="4acb4-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="4acb4-117">Member</span><span class="sxs-lookup"><span data-stu-id="4acb4-117">Member</span></span> |
| [<span data-ttu-id="4acb4-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="4acb4-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="4acb4-119">Member</span><span class="sxs-lookup"><span data-stu-id="4acb4-119">Member</span></span> |
| [<span data-ttu-id="4acb4-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="4acb4-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="4acb4-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4acb4-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="4acb4-122">Membres</span><span class="sxs-lookup"><span data-stu-id="4acb4-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="4acb4-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="4acb4-123">displayName: String</span></span>

<span data-ttu-id="4acb4-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4acb4-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4acb4-125">Type</span><span class="sxs-lookup"><span data-stu-id="4acb4-125">Type</span></span>

*   <span data-ttu-id="4acb4-126">String</span><span class="sxs-lookup"><span data-stu-id="4acb4-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4acb4-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4acb4-127">Requirements</span></span>

|<span data-ttu-id="4acb4-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4acb4-128">Requirement</span></span>| <span data-ttu-id="4acb4-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="4acb4-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="4acb4-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4acb4-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4acb4-131">1.0</span><span class="sxs-lookup"><span data-stu-id="4acb4-131">1.0</span></span>|
|[<span data-ttu-id="4acb4-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4acb4-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4acb4-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4acb4-133">ReadItem</span></span>|
|[<span data-ttu-id="4acb4-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4acb4-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4acb4-135">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4acb4-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4acb4-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="4acb4-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="4acb4-137">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="4acb4-137">emailAddress: String</span></span>

<span data-ttu-id="4acb4-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4acb4-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4acb4-139">Type</span><span class="sxs-lookup"><span data-stu-id="4acb4-139">Type</span></span>

*   <span data-ttu-id="4acb4-140">String</span><span class="sxs-lookup"><span data-stu-id="4acb4-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4acb4-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4acb4-141">Requirements</span></span>

|<span data-ttu-id="4acb4-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4acb4-142">Requirement</span></span>| <span data-ttu-id="4acb4-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="4acb4-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="4acb4-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4acb4-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4acb4-145">1.0</span><span class="sxs-lookup"><span data-stu-id="4acb4-145">1.0</span></span>|
|[<span data-ttu-id="4acb4-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4acb4-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4acb4-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4acb4-147">ReadItem</span></span>|
|[<span data-ttu-id="4acb4-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4acb4-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4acb4-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4acb4-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4acb4-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="4acb4-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="4acb4-151">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="4acb4-151">timeZone: String</span></span>

<span data-ttu-id="4acb4-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4acb4-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4acb4-153">Type</span><span class="sxs-lookup"><span data-stu-id="4acb4-153">Type</span></span>

*   <span data-ttu-id="4acb4-154">String</span><span class="sxs-lookup"><span data-stu-id="4acb4-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4acb4-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4acb4-155">Requirements</span></span>

|<span data-ttu-id="4acb4-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4acb4-156">Requirement</span></span>| <span data-ttu-id="4acb4-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="4acb4-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="4acb4-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4acb4-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4acb4-159">1.0</span><span class="sxs-lookup"><span data-stu-id="4acb4-159">1.0</span></span>|
|[<span data-ttu-id="4acb4-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4acb4-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4acb4-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4acb4-161">ReadItem</span></span>|
|[<span data-ttu-id="4acb4-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4acb4-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4acb4-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4acb4-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4acb4-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="4acb4-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
