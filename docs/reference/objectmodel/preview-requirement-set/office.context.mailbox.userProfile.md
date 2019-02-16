---
title: Office.context.mailbox.userProfile- prévisualisations d’ensemble de conditions requises
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 214434c988c01ecb1aef93f4067cd95bfe768ae9
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068174"
---
# <a name="userprofile"></a><span data-ttu-id="d7d6a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="d7d6a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="d7d6a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d7d6a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7d6a-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7d6a-104">Requirements</span></span>

|<span data-ttu-id="d7d6a-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7d6a-105">Requirement</span></span>| <span data-ttu-id="d7d6a-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7d6a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7d6a-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7d6a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7d6a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d7d6a-108">1.0</span></span>|
|[<span data-ttu-id="d7d6a-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d7d6a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7d6a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7d6a-110">ReadItem</span></span>|
|[<span data-ttu-id="d7d6a-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7d6a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7d6a-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7d6a-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d7d6a-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="d7d6a-113">Members and methods</span></span>

| <span data-ttu-id="d7d6a-114">Membre</span><span class="sxs-lookup"><span data-stu-id="d7d6a-114">Member</span></span> | <span data-ttu-id="d7d6a-115">Type</span><span class="sxs-lookup"><span data-stu-id="d7d6a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d7d6a-116">accountType</span><span class="sxs-lookup"><span data-stu-id="d7d6a-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="d7d6a-117">Membre</span><span class="sxs-lookup"><span data-stu-id="d7d6a-117">Member</span></span> |
| [<span data-ttu-id="d7d6a-118">displayName</span><span class="sxs-lookup"><span data-stu-id="d7d6a-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="d7d6a-119">Membre</span><span class="sxs-lookup"><span data-stu-id="d7d6a-119">Member</span></span> |
| [<span data-ttu-id="d7d6a-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d7d6a-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="d7d6a-121">Membre</span><span class="sxs-lookup"><span data-stu-id="d7d6a-121">Member</span></span> |
| [<span data-ttu-id="d7d6a-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="d7d6a-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="d7d6a-123">Membre</span><span class="sxs-lookup"><span data-stu-id="d7d6a-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d7d6a-124">Members</span><span class="sxs-lookup"><span data-stu-id="d7d6a-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="d7d6a-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="d7d6a-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="d7d6a-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou ultérieur).</span><span class="sxs-lookup"><span data-stu-id="d7d6a-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="d7d6a-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="d7d6a-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="d7d6a-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7d6a-129">Value</span></span> | <span data-ttu-id="d7d6a-130">Description</span><span class="sxs-lookup"><span data-stu-id="d7d6a-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="d7d6a-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="d7d6a-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="d7d6a-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="d7d6a-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="d7d6a-135">Type</span><span class="sxs-lookup"><span data-stu-id="d7d6a-135">Type</span></span>

*   <span data-ttu-id="d7d6a-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7d6a-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7d6a-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7d6a-137">Requirements</span></span>

|<span data-ttu-id="d7d6a-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7d6a-138">Requirement</span></span>| <span data-ttu-id="d7d6a-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7d6a-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7d6a-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7d6a-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7d6a-141">1.6</span><span class="sxs-lookup"><span data-stu-id="d7d6a-141">1.6</span></span> |
|[<span data-ttu-id="d7d6a-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d7d6a-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7d6a-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7d6a-143">ReadItem</span></span>|
|[<span data-ttu-id="d7d6a-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7d6a-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7d6a-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7d6a-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7d6a-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7d6a-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="d7d6a-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="d7d6a-147">displayName :String</span></span>

<span data-ttu-id="d7d6a-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d7d6a-149">Type</span><span class="sxs-lookup"><span data-stu-id="d7d6a-149">Type</span></span>

*   <span data-ttu-id="d7d6a-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7d6a-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7d6a-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7d6a-151">Requirements</span></span>

|<span data-ttu-id="d7d6a-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7d6a-152">Requirement</span></span>| <span data-ttu-id="d7d6a-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7d6a-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7d6a-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7d6a-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7d6a-155">1.0</span><span class="sxs-lookup"><span data-stu-id="d7d6a-155">1.0</span></span>|
|[<span data-ttu-id="d7d6a-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d7d6a-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7d6a-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7d6a-157">ReadItem</span></span>|
|[<span data-ttu-id="d7d6a-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7d6a-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7d6a-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7d6a-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7d6a-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7d6a-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d7d6a-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="d7d6a-161">emailAddress :String</span></span>

<span data-ttu-id="d7d6a-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d7d6a-163">Type</span><span class="sxs-lookup"><span data-stu-id="d7d6a-163">Type</span></span>

*   <span data-ttu-id="d7d6a-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7d6a-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7d6a-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7d6a-165">Requirements</span></span>

|<span data-ttu-id="d7d6a-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7d6a-166">Requirement</span></span>| <span data-ttu-id="d7d6a-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7d6a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7d6a-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7d6a-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7d6a-169">1.0</span><span class="sxs-lookup"><span data-stu-id="d7d6a-169">1.0</span></span>|
|[<span data-ttu-id="d7d6a-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d7d6a-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7d6a-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7d6a-171">ReadItem</span></span>|
|[<span data-ttu-id="d7d6a-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7d6a-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7d6a-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7d6a-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7d6a-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7d6a-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d7d6a-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="d7d6a-175">timeZone :String</span></span>

<span data-ttu-id="d7d6a-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d7d6a-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d7d6a-177">Type</span><span class="sxs-lookup"><span data-stu-id="d7d6a-177">Type</span></span>

*   <span data-ttu-id="d7d6a-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7d6a-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7d6a-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7d6a-179">Requirements</span></span>

|<span data-ttu-id="d7d6a-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7d6a-180">Requirement</span></span>| <span data-ttu-id="d7d6a-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7d6a-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7d6a-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7d6a-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7d6a-183">1.0</span><span class="sxs-lookup"><span data-stu-id="d7d6a-183">1.0</span></span>|
|[<span data-ttu-id="d7d6a-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d7d6a-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7d6a-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7d6a-185">ReadItem</span></span>|
|[<span data-ttu-id="d7d6a-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7d6a-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7d6a-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7d6a-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7d6a-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7d6a-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
