---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451821"
---
# <a name="userprofile"></a><span data-ttu-id="df4ac-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="df4ac-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="df4ac-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="df4ac-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="df4ac-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="df4ac-104">Requirements</span></span>

|<span data-ttu-id="df4ac-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="df4ac-105">Requirement</span></span>| <span data-ttu-id="df4ac-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="df4ac-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="df4ac-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="df4ac-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df4ac-108">1.0</span><span class="sxs-lookup"><span data-stu-id="df4ac-108">1.0</span></span>|
|[<span data-ttu-id="df4ac-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="df4ac-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df4ac-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df4ac-110">ReadItem</span></span>|
|[<span data-ttu-id="df4ac-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="df4ac-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df4ac-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="df4ac-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="df4ac-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="df4ac-113">Members and methods</span></span>

| <span data-ttu-id="df4ac-114">Membre</span><span class="sxs-lookup"><span data-stu-id="df4ac-114">Member</span></span> | <span data-ttu-id="df4ac-115">Type</span><span class="sxs-lookup"><span data-stu-id="df4ac-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="df4ac-116">displayName</span><span class="sxs-lookup"><span data-stu-id="df4ac-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="df4ac-117">Member</span><span class="sxs-lookup"><span data-stu-id="df4ac-117">Member</span></span> |
| [<span data-ttu-id="df4ac-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="df4ac-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="df4ac-119">Member</span><span class="sxs-lookup"><span data-stu-id="df4ac-119">Member</span></span> |
| [<span data-ttu-id="df4ac-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="df4ac-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="df4ac-121">Membre</span><span class="sxs-lookup"><span data-stu-id="df4ac-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="df4ac-122">Membres</span><span class="sxs-lookup"><span data-stu-id="df4ac-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="df4ac-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="df4ac-123">displayName :String</span></span>

<span data-ttu-id="df4ac-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="df4ac-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="df4ac-125">Type</span><span class="sxs-lookup"><span data-stu-id="df4ac-125">Type</span></span>

*   <span data-ttu-id="df4ac-126">String</span><span class="sxs-lookup"><span data-stu-id="df4ac-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df4ac-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="df4ac-127">Requirements</span></span>

|<span data-ttu-id="df4ac-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="df4ac-128">Requirement</span></span>| <span data-ttu-id="df4ac-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="df4ac-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="df4ac-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="df4ac-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df4ac-131">1.0</span><span class="sxs-lookup"><span data-stu-id="df4ac-131">1.0</span></span>|
|[<span data-ttu-id="df4ac-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="df4ac-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df4ac-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df4ac-133">ReadItem</span></span>|
|[<span data-ttu-id="df4ac-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="df4ac-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df4ac-135">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="df4ac-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df4ac-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="df4ac-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="df4ac-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="df4ac-137">emailAddress :String</span></span>

<span data-ttu-id="df4ac-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="df4ac-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="df4ac-139">Type</span><span class="sxs-lookup"><span data-stu-id="df4ac-139">Type</span></span>

*   <span data-ttu-id="df4ac-140">String</span><span class="sxs-lookup"><span data-stu-id="df4ac-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df4ac-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="df4ac-141">Requirements</span></span>

|<span data-ttu-id="df4ac-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="df4ac-142">Requirement</span></span>| <span data-ttu-id="df4ac-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="df4ac-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="df4ac-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="df4ac-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df4ac-145">1.0</span><span class="sxs-lookup"><span data-stu-id="df4ac-145">1.0</span></span>|
|[<span data-ttu-id="df4ac-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="df4ac-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df4ac-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df4ac-147">ReadItem</span></span>|
|[<span data-ttu-id="df4ac-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="df4ac-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df4ac-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="df4ac-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df4ac-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="df4ac-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="df4ac-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="df4ac-151">timeZone :String</span></span>

<span data-ttu-id="df4ac-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="df4ac-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="df4ac-153">Type</span><span class="sxs-lookup"><span data-stu-id="df4ac-153">Type</span></span>

*   <span data-ttu-id="df4ac-154">String</span><span class="sxs-lookup"><span data-stu-id="df4ac-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df4ac-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="df4ac-155">Requirements</span></span>

|<span data-ttu-id="df4ac-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="df4ac-156">Requirement</span></span>| <span data-ttu-id="df4ac-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="df4ac-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="df4ac-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="df4ac-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df4ac-159">1.0</span><span class="sxs-lookup"><span data-stu-id="df4ac-159">1.0</span></span>|
|[<span data-ttu-id="df4ac-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="df4ac-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df4ac-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df4ac-161">ReadItem</span></span>|
|[<span data-ttu-id="df4ac-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="df4ac-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df4ac-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="df4ac-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df4ac-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="df4ac-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
