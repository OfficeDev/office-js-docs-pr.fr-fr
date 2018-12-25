---
title: Office.context.mailbox.userProfile - requirement set 1.6
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: fe30a390583dc646e9c8792710c580d02c373a1a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432894"
---
# <a name="userprofile"></a><span data-ttu-id="3ad87-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="3ad87-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="3ad87-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="3ad87-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ad87-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3ad87-104">Requirements</span></span>

|<span data-ttu-id="3ad87-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3ad87-105">Requirement</span></span>| <span data-ttu-id="3ad87-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="3ad87-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ad87-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3ad87-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ad87-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3ad87-108">1.0</span></span>|
|[<span data-ttu-id="3ad87-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3ad87-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ad87-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ad87-110">ReadItem</span></span>|
|[<span data-ttu-id="3ad87-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3ad87-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ad87-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="3ad87-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3ad87-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="3ad87-113">Members and methods</span></span>

| <span data-ttu-id="3ad87-114">Membre</span><span class="sxs-lookup"><span data-stu-id="3ad87-114">Member</span></span> | <span data-ttu-id="3ad87-115">Type</span><span class="sxs-lookup"><span data-stu-id="3ad87-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3ad87-116">accountType</span><span class="sxs-lookup"><span data-stu-id="3ad87-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="3ad87-117">Member</span><span class="sxs-lookup"><span data-stu-id="3ad87-117">Member</span></span> |
| [<span data-ttu-id="3ad87-118">displayName</span><span class="sxs-lookup"><span data-stu-id="3ad87-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="3ad87-119">Membre</span><span class="sxs-lookup"><span data-stu-id="3ad87-119">Member</span></span> |
| [<span data-ttu-id="3ad87-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="3ad87-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="3ad87-121">Membre</span><span class="sxs-lookup"><span data-stu-id="3ad87-121">Member</span></span> |
| [<span data-ttu-id="3ad87-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="3ad87-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="3ad87-123">Membre</span><span class="sxs-lookup"><span data-stu-id="3ad87-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="3ad87-124">Members</span><span class="sxs-lookup"><span data-stu-id="3ad87-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="3ad87-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="3ad87-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="3ad87-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou ultérieur).</span><span class="sxs-lookup"><span data-stu-id="3ad87-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="3ad87-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="3ad87-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="3ad87-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="3ad87-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="3ad87-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="3ad87-129">Value</span></span> | <span data-ttu-id="3ad87-130">Description</span><span class="sxs-lookup"><span data-stu-id="3ad87-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="3ad87-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="3ad87-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="3ad87-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="3ad87-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="3ad87-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="3ad87-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="3ad87-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="3ad87-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="3ad87-135">Type :</span><span class="sxs-lookup"><span data-stu-id="3ad87-135">Type:</span></span>

*   <span data-ttu-id="3ad87-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3ad87-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ad87-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3ad87-137">Requirements</span></span>

|<span data-ttu-id="3ad87-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3ad87-138">Requirement</span></span>| <span data-ttu-id="3ad87-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="3ad87-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ad87-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3ad87-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ad87-141">1.6</span><span class="sxs-lookup"><span data-stu-id="3ad87-141">1.6</span></span> |
|[<span data-ttu-id="3ad87-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3ad87-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ad87-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ad87-143">ReadItem</span></span>|
|[<span data-ttu-id="3ad87-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3ad87-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ad87-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="3ad87-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ad87-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="3ad87-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="3ad87-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="3ad87-147">displayName :String</span></span>

<span data-ttu-id="3ad87-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3ad87-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3ad87-149">Type :</span><span class="sxs-lookup"><span data-stu-id="3ad87-149">Type:</span></span>

*   <span data-ttu-id="3ad87-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3ad87-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ad87-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3ad87-151">Requirements</span></span>

|<span data-ttu-id="3ad87-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3ad87-152">Requirement</span></span>| <span data-ttu-id="3ad87-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="3ad87-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ad87-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3ad87-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ad87-155">1.0</span><span class="sxs-lookup"><span data-stu-id="3ad87-155">1.0</span></span>|
|[<span data-ttu-id="3ad87-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3ad87-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ad87-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ad87-157">ReadItem</span></span>|
|[<span data-ttu-id="3ad87-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3ad87-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ad87-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="3ad87-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ad87-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="3ad87-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="3ad87-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="3ad87-161">emailAddress :String</span></span>

<span data-ttu-id="3ad87-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3ad87-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3ad87-163">Type :</span><span class="sxs-lookup"><span data-stu-id="3ad87-163">Type:</span></span>

*   <span data-ttu-id="3ad87-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3ad87-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ad87-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3ad87-165">Requirements</span></span>

|<span data-ttu-id="3ad87-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3ad87-166">Requirement</span></span>| <span data-ttu-id="3ad87-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="3ad87-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ad87-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3ad87-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ad87-169">1.0</span><span class="sxs-lookup"><span data-stu-id="3ad87-169">1.0</span></span>|
|[<span data-ttu-id="3ad87-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3ad87-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ad87-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ad87-171">ReadItem</span></span>|
|[<span data-ttu-id="3ad87-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3ad87-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ad87-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="3ad87-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ad87-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="3ad87-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="3ad87-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="3ad87-175">timeZone :String</span></span>

<span data-ttu-id="3ad87-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3ad87-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3ad87-177">Type :</span><span class="sxs-lookup"><span data-stu-id="3ad87-177">Type:</span></span>

*   <span data-ttu-id="3ad87-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3ad87-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ad87-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3ad87-179">Requirements</span></span>

|<span data-ttu-id="3ad87-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3ad87-180">Requirement</span></span>| <span data-ttu-id="3ad87-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="3ad87-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ad87-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3ad87-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ad87-183">1.0</span><span class="sxs-lookup"><span data-stu-id="3ad87-183">1.0</span></span>|
|[<span data-ttu-id="3ad87-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3ad87-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ad87-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ad87-185">ReadItem</span></span>|
|[<span data-ttu-id="3ad87-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3ad87-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ad87-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="3ad87-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ad87-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="3ad87-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```