---
title: Office.context.mailbox.userProfile- prévisualisations d’ensemble de conditions requises
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 061ee8367005f4af0795c4d9e1236d0b2443521a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432815"
---
# <a name="userprofile"></a><span data-ttu-id="e5814-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e5814-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e5814-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e5814-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5814-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e5814-104">Requirements</span></span>

|<span data-ttu-id="e5814-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e5814-105">Requirement</span></span>| <span data-ttu-id="e5814-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="e5814-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5814-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e5814-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5814-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e5814-108">1.0</span></span>|
|[<span data-ttu-id="e5814-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e5814-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5814-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5814-110">ReadItem</span></span>|
|[<span data-ttu-id="e5814-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e5814-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5814-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e5814-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e5814-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="e5814-113">Members and methods</span></span>

| <span data-ttu-id="e5814-114">Membre</span><span class="sxs-lookup"><span data-stu-id="e5814-114">Member</span></span> | <span data-ttu-id="e5814-115">Type</span><span class="sxs-lookup"><span data-stu-id="e5814-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e5814-116">accountType</span><span class="sxs-lookup"><span data-stu-id="e5814-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="e5814-117">Member</span><span class="sxs-lookup"><span data-stu-id="e5814-117">Member</span></span> |
| [<span data-ttu-id="e5814-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e5814-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e5814-119">Membre</span><span class="sxs-lookup"><span data-stu-id="e5814-119">Member</span></span> |
| [<span data-ttu-id="e5814-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e5814-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e5814-121">Membre</span><span class="sxs-lookup"><span data-stu-id="e5814-121">Member</span></span> |
| [<span data-ttu-id="e5814-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e5814-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e5814-123">Membre</span><span class="sxs-lookup"><span data-stu-id="e5814-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e5814-124">Members</span><span class="sxs-lookup"><span data-stu-id="e5814-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="e5814-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="e5814-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="e5814-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou ultérieur).</span><span class="sxs-lookup"><span data-stu-id="e5814-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="e5814-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e5814-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e5814-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="e5814-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="e5814-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="e5814-129">Value</span></span> | <span data-ttu-id="e5814-130">Description</span><span class="sxs-lookup"><span data-stu-id="e5814-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e5814-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="e5814-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e5814-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="e5814-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e5814-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="e5814-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e5814-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="e5814-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e5814-135">Type :</span><span class="sxs-lookup"><span data-stu-id="e5814-135">Type:</span></span>

*   <span data-ttu-id="e5814-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e5814-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5814-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e5814-137">Requirements</span></span>

|<span data-ttu-id="e5814-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e5814-138">Requirement</span></span>| <span data-ttu-id="e5814-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="e5814-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5814-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e5814-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5814-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e5814-141">1.6</span></span> |
|[<span data-ttu-id="e5814-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e5814-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5814-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5814-143">ReadItem</span></span>|
|[<span data-ttu-id="e5814-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e5814-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5814-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e5814-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5814-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="e5814-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="e5814-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e5814-147">displayName :String</span></span>

<span data-ttu-id="e5814-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e5814-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e5814-149">Type :</span><span class="sxs-lookup"><span data-stu-id="e5814-149">Type:</span></span>

*   <span data-ttu-id="e5814-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e5814-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5814-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e5814-151">Requirements</span></span>

|<span data-ttu-id="e5814-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e5814-152">Requirement</span></span>| <span data-ttu-id="e5814-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="e5814-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5814-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e5814-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5814-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e5814-155">1.0</span></span>|
|[<span data-ttu-id="e5814-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e5814-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5814-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5814-157">ReadItem</span></span>|
|[<span data-ttu-id="e5814-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e5814-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5814-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e5814-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5814-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="e5814-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="e5814-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e5814-161">emailAddress :String</span></span>

<span data-ttu-id="e5814-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e5814-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e5814-163">Type :</span><span class="sxs-lookup"><span data-stu-id="e5814-163">Type:</span></span>

*   <span data-ttu-id="e5814-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e5814-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5814-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e5814-165">Requirements</span></span>

|<span data-ttu-id="e5814-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e5814-166">Requirement</span></span>| <span data-ttu-id="e5814-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="e5814-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5814-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e5814-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5814-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e5814-169">1.0</span></span>|
|[<span data-ttu-id="e5814-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e5814-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5814-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5814-171">ReadItem</span></span>|
|[<span data-ttu-id="e5814-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e5814-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5814-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e5814-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5814-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="e5814-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="e5814-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e5814-175">timeZone :String</span></span>

<span data-ttu-id="e5814-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e5814-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e5814-177">Type :</span><span class="sxs-lookup"><span data-stu-id="e5814-177">Type:</span></span>

*   <span data-ttu-id="e5814-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e5814-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5814-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e5814-179">Requirements</span></span>

|<span data-ttu-id="e5814-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e5814-180">Requirement</span></span>| <span data-ttu-id="e5814-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="e5814-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5814-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e5814-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5814-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e5814-183">1.0</span></span>|
|[<span data-ttu-id="e5814-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e5814-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5814-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5814-185">ReadItem</span></span>|
|[<span data-ttu-id="e5814-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e5814-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5814-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="e5814-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5814-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="e5814-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```