---
title: Office.context.mailbox.userProfile - requirement set 1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 09457a41fe68ae03e035d3d3f4b80b139be348e0
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067873"
---
# <a name="userprofile"></a><span data-ttu-id="79e08-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="79e08-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="79e08-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="79e08-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="79e08-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="79e08-104">Requirements</span></span>

|<span data-ttu-id="79e08-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="79e08-105">Requirement</span></span>| <span data-ttu-id="79e08-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="79e08-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="79e08-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="79e08-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79e08-108">1.0</span><span class="sxs-lookup"><span data-stu-id="79e08-108">1.0</span></span>|
|[<span data-ttu-id="79e08-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="79e08-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79e08-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79e08-110">ReadItem</span></span>|
|[<span data-ttu-id="79e08-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="79e08-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79e08-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="79e08-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="79e08-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="79e08-113">Members and methods</span></span>

| <span data-ttu-id="79e08-114">Membre</span><span class="sxs-lookup"><span data-stu-id="79e08-114">Member</span></span> | <span data-ttu-id="79e08-115">Type</span><span class="sxs-lookup"><span data-stu-id="79e08-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="79e08-116">accountType</span><span class="sxs-lookup"><span data-stu-id="79e08-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="79e08-117">Membre</span><span class="sxs-lookup"><span data-stu-id="79e08-117">Member</span></span> |
| [<span data-ttu-id="79e08-118">displayName</span><span class="sxs-lookup"><span data-stu-id="79e08-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="79e08-119">Membre</span><span class="sxs-lookup"><span data-stu-id="79e08-119">Member</span></span> |
| [<span data-ttu-id="79e08-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="79e08-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="79e08-121">Membre</span><span class="sxs-lookup"><span data-stu-id="79e08-121">Member</span></span> |
| [<span data-ttu-id="79e08-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="79e08-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="79e08-123">Membre</span><span class="sxs-lookup"><span data-stu-id="79e08-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="79e08-124">Members</span><span class="sxs-lookup"><span data-stu-id="79e08-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="79e08-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="79e08-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="79e08-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou ultérieur).</span><span class="sxs-lookup"><span data-stu-id="79e08-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="79e08-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="79e08-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="79e08-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="79e08-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="79e08-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="79e08-129">Value</span></span> | <span data-ttu-id="79e08-130">Description</span><span class="sxs-lookup"><span data-stu-id="79e08-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="79e08-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="79e08-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="79e08-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="79e08-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="79e08-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="79e08-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="79e08-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="79e08-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="79e08-135">Type</span><span class="sxs-lookup"><span data-stu-id="79e08-135">Type</span></span>

*   <span data-ttu-id="79e08-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="79e08-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79e08-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="79e08-137">Requirements</span></span>

|<span data-ttu-id="79e08-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="79e08-138">Requirement</span></span>| <span data-ttu-id="79e08-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="79e08-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="79e08-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="79e08-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79e08-141">1.6</span><span class="sxs-lookup"><span data-stu-id="79e08-141">1.6</span></span> |
|[<span data-ttu-id="79e08-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="79e08-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79e08-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79e08-143">ReadItem</span></span>|
|[<span data-ttu-id="79e08-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="79e08-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79e08-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="79e08-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79e08-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="79e08-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="79e08-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="79e08-147">displayName :String</span></span>

<span data-ttu-id="79e08-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="79e08-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="79e08-149">Type</span><span class="sxs-lookup"><span data-stu-id="79e08-149">Type</span></span>

*   <span data-ttu-id="79e08-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="79e08-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79e08-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="79e08-151">Requirements</span></span>

|<span data-ttu-id="79e08-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="79e08-152">Requirement</span></span>| <span data-ttu-id="79e08-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="79e08-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="79e08-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="79e08-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79e08-155">1.0</span><span class="sxs-lookup"><span data-stu-id="79e08-155">1.0</span></span>|
|[<span data-ttu-id="79e08-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="79e08-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79e08-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79e08-157">ReadItem</span></span>|
|[<span data-ttu-id="79e08-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="79e08-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79e08-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="79e08-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79e08-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="79e08-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="79e08-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="79e08-161">emailAddress :String</span></span>

<span data-ttu-id="79e08-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="79e08-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="79e08-163">Type</span><span class="sxs-lookup"><span data-stu-id="79e08-163">Type</span></span>

*   <span data-ttu-id="79e08-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="79e08-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79e08-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="79e08-165">Requirements</span></span>

|<span data-ttu-id="79e08-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="79e08-166">Requirement</span></span>| <span data-ttu-id="79e08-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="79e08-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="79e08-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="79e08-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79e08-169">1.0</span><span class="sxs-lookup"><span data-stu-id="79e08-169">1.0</span></span>|
|[<span data-ttu-id="79e08-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="79e08-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79e08-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79e08-171">ReadItem</span></span>|
|[<span data-ttu-id="79e08-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="79e08-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79e08-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="79e08-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79e08-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="79e08-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="79e08-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="79e08-175">timeZone :String</span></span>

<span data-ttu-id="79e08-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="79e08-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="79e08-177">Type</span><span class="sxs-lookup"><span data-stu-id="79e08-177">Type</span></span>

*   <span data-ttu-id="79e08-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="79e08-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79e08-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="79e08-179">Requirements</span></span>

|<span data-ttu-id="79e08-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="79e08-180">Requirement</span></span>| <span data-ttu-id="79e08-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="79e08-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="79e08-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="79e08-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79e08-183">1.0</span><span class="sxs-lookup"><span data-stu-id="79e08-183">1.0</span></span>|
|[<span data-ttu-id="79e08-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="79e08-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79e08-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79e08-185">ReadItem</span></span>|
|[<span data-ttu-id="79e08-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="79e08-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79e08-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="79e08-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79e08-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="79e08-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
