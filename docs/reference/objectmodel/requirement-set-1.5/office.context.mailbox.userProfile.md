---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.5
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 748daf4d14aae1d14560d29e1d76eeea09830573
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432717"
---
# <a name="userprofile"></a><span data-ttu-id="94742-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="94742-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="94742-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="94742-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="94742-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="94742-104">Requirements</span></span>

|<span data-ttu-id="94742-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="94742-105">Requirement</span></span>| <span data-ttu-id="94742-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="94742-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="94742-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="94742-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94742-108">1.0</span><span class="sxs-lookup"><span data-stu-id="94742-108">1.0</span></span>|
|[<span data-ttu-id="94742-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="94742-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="94742-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="94742-110">ReadItem</span></span>|
|[<span data-ttu-id="94742-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="94742-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="94742-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="94742-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="94742-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="94742-113">Members and methods</span></span>

| <span data-ttu-id="94742-114">Membre</span><span class="sxs-lookup"><span data-stu-id="94742-114">Member</span></span> | <span data-ttu-id="94742-115">Type</span><span class="sxs-lookup"><span data-stu-id="94742-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="94742-116">displayName</span><span class="sxs-lookup"><span data-stu-id="94742-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="94742-117">Membre</span><span class="sxs-lookup"><span data-stu-id="94742-117">Member</span></span> |
| [<span data-ttu-id="94742-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="94742-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="94742-119">Membre</span><span class="sxs-lookup"><span data-stu-id="94742-119">Member</span></span> |
| [<span data-ttu-id="94742-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="94742-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="94742-121">Membre</span><span class="sxs-lookup"><span data-stu-id="94742-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="94742-122">Membres</span><span class="sxs-lookup"><span data-stu-id="94742-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="94742-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="94742-123">displayName :String</span></span>

<span data-ttu-id="94742-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="94742-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="94742-125">Type :</span><span class="sxs-lookup"><span data-stu-id="94742-125">Type:</span></span>

*   <span data-ttu-id="94742-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="94742-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="94742-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="94742-127">Requirements</span></span>

|<span data-ttu-id="94742-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="94742-128">Requirement</span></span>| <span data-ttu-id="94742-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="94742-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="94742-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="94742-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94742-131">1.0</span><span class="sxs-lookup"><span data-stu-id="94742-131">1.0</span></span>|
|[<span data-ttu-id="94742-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="94742-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="94742-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="94742-133">ReadItem</span></span>|
|[<span data-ttu-id="94742-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="94742-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="94742-135">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="94742-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="94742-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="94742-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="94742-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="94742-137">emailAddress :String</span></span>

<span data-ttu-id="94742-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="94742-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="94742-139">Type :</span><span class="sxs-lookup"><span data-stu-id="94742-139">Type:</span></span>

*   <span data-ttu-id="94742-140">Chaîne</span><span class="sxs-lookup"><span data-stu-id="94742-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="94742-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="94742-141">Requirements</span></span>

|<span data-ttu-id="94742-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="94742-142">Requirement</span></span>| <span data-ttu-id="94742-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="94742-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="94742-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="94742-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94742-145">1.0</span><span class="sxs-lookup"><span data-stu-id="94742-145">1.0</span></span>|
|[<span data-ttu-id="94742-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="94742-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="94742-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="94742-147">ReadItem</span></span>|
|[<span data-ttu-id="94742-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="94742-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="94742-149">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="94742-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="94742-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="94742-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="94742-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="94742-151">timeZone :String</span></span>

<span data-ttu-id="94742-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="94742-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="94742-153">Type :</span><span class="sxs-lookup"><span data-stu-id="94742-153">Type:</span></span>

*   <span data-ttu-id="94742-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="94742-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="94742-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="94742-155">Requirements</span></span>

|<span data-ttu-id="94742-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="94742-156">Requirement</span></span>| <span data-ttu-id="94742-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="94742-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="94742-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="94742-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94742-159">1.0</span><span class="sxs-lookup"><span data-stu-id="94742-159">1.0</span></span>|
|[<span data-ttu-id="94742-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="94742-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="94742-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="94742-161">ReadItem</span></span>|
|[<span data-ttu-id="94742-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="94742-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="94742-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="94742-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="94742-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="94742-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```