---
title: Office.context.mailbox.userProfile-ensemble de conditions requises 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: fb55d11fd46a9957dab124514ef3bfe5a7c138eb
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067866"
---
# <a name="userprofile"></a><span data-ttu-id="c1b40-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c1b40-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c1b40-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c1b40-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1b40-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1b40-104">Requirements</span></span>

|<span data-ttu-id="c1b40-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1b40-105">Requirement</span></span>| <span data-ttu-id="c1b40-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1b40-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1b40-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1b40-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1b40-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c1b40-108">1.0</span></span>|
|[<span data-ttu-id="c1b40-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1b40-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1b40-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1b40-110">ReadItem</span></span>|
|[<span data-ttu-id="c1b40-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1b40-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1b40-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1b40-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1b40-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c1b40-113">Members and methods</span></span>

| <span data-ttu-id="c1b40-114">Membre</span><span class="sxs-lookup"><span data-stu-id="c1b40-114">Member</span></span> | <span data-ttu-id="c1b40-115">Type</span><span class="sxs-lookup"><span data-stu-id="c1b40-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1b40-116">accountType</span><span class="sxs-lookup"><span data-stu-id="c1b40-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="c1b40-117">Membre</span><span class="sxs-lookup"><span data-stu-id="c1b40-117">Member</span></span> |
| [<span data-ttu-id="c1b40-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c1b40-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c1b40-119">Membre</span><span class="sxs-lookup"><span data-stu-id="c1b40-119">Member</span></span> |
| [<span data-ttu-id="c1b40-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c1b40-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c1b40-121">Membre</span><span class="sxs-lookup"><span data-stu-id="c1b40-121">Member</span></span> |
| [<span data-ttu-id="c1b40-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c1b40-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c1b40-123">Membre</span><span class="sxs-lookup"><span data-stu-id="c1b40-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c1b40-124">Members</span><span class="sxs-lookup"><span data-stu-id="c1b40-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c1b40-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="c1b40-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c1b40-126">Actuellement, ce membre est uniquement pris en charge par Outlook 2016 pour Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="c1b40-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="c1b40-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="c1b40-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c1b40-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="c1b40-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="c1b40-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1b40-129">Value</span></span> | <span data-ttu-id="c1b40-130">Description</span><span class="sxs-lookup"><span data-stu-id="c1b40-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c1b40-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="c1b40-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c1b40-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="c1b40-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c1b40-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="c1b40-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c1b40-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="c1b40-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c1b40-135">Type</span><span class="sxs-lookup"><span data-stu-id="c1b40-135">Type</span></span>

*   <span data-ttu-id="c1b40-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1b40-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1b40-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1b40-137">Requirements</span></span>

|<span data-ttu-id="c1b40-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1b40-138">Requirement</span></span>| <span data-ttu-id="c1b40-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1b40-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1b40-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1b40-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1b40-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c1b40-141">1.6</span></span> |
|[<span data-ttu-id="c1b40-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1b40-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1b40-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1b40-143">ReadItem</span></span>|
|[<span data-ttu-id="c1b40-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1b40-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1b40-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1b40-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1b40-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1b40-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="c1b40-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c1b40-147">displayName :String</span></span>

<span data-ttu-id="c1b40-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c1b40-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c1b40-149">Type</span><span class="sxs-lookup"><span data-stu-id="c1b40-149">Type</span></span>

*   <span data-ttu-id="c1b40-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1b40-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1b40-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1b40-151">Requirements</span></span>

|<span data-ttu-id="c1b40-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1b40-152">Requirement</span></span>| <span data-ttu-id="c1b40-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1b40-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1b40-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1b40-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1b40-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c1b40-155">1.0</span></span>|
|[<span data-ttu-id="c1b40-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1b40-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1b40-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1b40-157">ReadItem</span></span>|
|[<span data-ttu-id="c1b40-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1b40-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1b40-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1b40-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1b40-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1b40-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c1b40-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c1b40-161">emailAddress :String</span></span>

<span data-ttu-id="c1b40-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c1b40-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c1b40-163">Type</span><span class="sxs-lookup"><span data-stu-id="c1b40-163">Type</span></span>

*   <span data-ttu-id="c1b40-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1b40-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1b40-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1b40-165">Requirements</span></span>

|<span data-ttu-id="c1b40-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1b40-166">Requirement</span></span>| <span data-ttu-id="c1b40-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1b40-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1b40-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1b40-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1b40-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c1b40-169">1.0</span></span>|
|[<span data-ttu-id="c1b40-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1b40-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1b40-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1b40-171">ReadItem</span></span>|
|[<span data-ttu-id="c1b40-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1b40-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1b40-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1b40-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1b40-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1b40-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c1b40-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c1b40-175">timeZone :String</span></span>

<span data-ttu-id="c1b40-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c1b40-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c1b40-177">Type</span><span class="sxs-lookup"><span data-stu-id="c1b40-177">Type</span></span>

*   <span data-ttu-id="c1b40-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1b40-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1b40-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1b40-179">Requirements</span></span>

|<span data-ttu-id="c1b40-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1b40-180">Requirement</span></span>| <span data-ttu-id="c1b40-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1b40-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1b40-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1b40-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1b40-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c1b40-183">1.0</span></span>|
|[<span data-ttu-id="c1b40-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1b40-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1b40-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1b40-185">ReadItem</span></span>|
|[<span data-ttu-id="c1b40-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1b40-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1b40-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1b40-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1b40-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1b40-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
