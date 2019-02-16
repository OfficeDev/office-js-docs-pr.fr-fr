---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e98e88cde184db121e69fdd267dff4e39d887b1f
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067828"
---
# <a name="userprofile"></a><span data-ttu-id="98308-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="98308-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="98308-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="98308-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="98308-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="98308-104">Requirements</span></span>

|<span data-ttu-id="98308-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="98308-105">Requirement</span></span>| <span data-ttu-id="98308-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="98308-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="98308-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="98308-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98308-108">1.0</span><span class="sxs-lookup"><span data-stu-id="98308-108">1.0</span></span>|
|[<span data-ttu-id="98308-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="98308-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98308-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98308-110">ReadItem</span></span>|
|[<span data-ttu-id="98308-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="98308-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98308-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="98308-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="98308-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="98308-113">Members and methods</span></span>

| <span data-ttu-id="98308-114">Membre</span><span class="sxs-lookup"><span data-stu-id="98308-114">Member</span></span> | <span data-ttu-id="98308-115">Type</span><span class="sxs-lookup"><span data-stu-id="98308-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="98308-116">displayName</span><span class="sxs-lookup"><span data-stu-id="98308-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="98308-117">Membre</span><span class="sxs-lookup"><span data-stu-id="98308-117">Member</span></span> |
| [<span data-ttu-id="98308-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="98308-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="98308-119">Membre</span><span class="sxs-lookup"><span data-stu-id="98308-119">Member</span></span> |
| [<span data-ttu-id="98308-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="98308-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="98308-121">Membre</span><span class="sxs-lookup"><span data-stu-id="98308-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="98308-122">Membres</span><span class="sxs-lookup"><span data-stu-id="98308-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="98308-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="98308-123">displayName :String</span></span>

<span data-ttu-id="98308-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="98308-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="98308-125">Type</span><span class="sxs-lookup"><span data-stu-id="98308-125">Type</span></span>

*   <span data-ttu-id="98308-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="98308-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98308-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="98308-127">Requirements</span></span>

|<span data-ttu-id="98308-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="98308-128">Requirement</span></span>| <span data-ttu-id="98308-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="98308-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="98308-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="98308-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98308-131">1.0</span><span class="sxs-lookup"><span data-stu-id="98308-131">1.0</span></span>|
|[<span data-ttu-id="98308-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="98308-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98308-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98308-133">ReadItem</span></span>|
|[<span data-ttu-id="98308-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="98308-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98308-135">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="98308-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98308-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="98308-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="98308-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="98308-137">emailAddress :String</span></span>

<span data-ttu-id="98308-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="98308-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="98308-139">Type</span><span class="sxs-lookup"><span data-stu-id="98308-139">Type</span></span>

*   <span data-ttu-id="98308-140">Chaîne</span><span class="sxs-lookup"><span data-stu-id="98308-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98308-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="98308-141">Requirements</span></span>

|<span data-ttu-id="98308-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="98308-142">Requirement</span></span>| <span data-ttu-id="98308-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="98308-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="98308-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="98308-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98308-145">1.0</span><span class="sxs-lookup"><span data-stu-id="98308-145">1.0</span></span>|
|[<span data-ttu-id="98308-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="98308-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98308-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98308-147">ReadItem</span></span>|
|[<span data-ttu-id="98308-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="98308-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98308-149">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="98308-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98308-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="98308-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="98308-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="98308-151">timeZone :String</span></span>

<span data-ttu-id="98308-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="98308-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="98308-153">Type</span><span class="sxs-lookup"><span data-stu-id="98308-153">Type</span></span>

*   <span data-ttu-id="98308-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="98308-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98308-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="98308-155">Requirements</span></span>

|<span data-ttu-id="98308-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="98308-156">Requirement</span></span>| <span data-ttu-id="98308-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="98308-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="98308-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="98308-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98308-159">1.0</span><span class="sxs-lookup"><span data-stu-id="98308-159">1.0</span></span>|
|[<span data-ttu-id="98308-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="98308-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98308-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98308-161">ReadItem</span></span>|
|[<span data-ttu-id="98308-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="98308-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98308-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="98308-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98308-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="98308-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
