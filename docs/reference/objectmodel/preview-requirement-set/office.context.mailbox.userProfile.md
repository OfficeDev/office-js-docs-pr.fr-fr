---
title: Office. Context. Mailbox. userProfile-aperçu de l'ensemble de conditions requises
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 204097497c958c26a6e67fc01d6dbd5142d8dced
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871870"
---
# <a name="userprofile"></a><span data-ttu-id="185fd-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="185fd-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="185fd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="185fd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="185fd-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185fd-104">Requirements</span></span>

|<span data-ttu-id="185fd-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185fd-105">Requirement</span></span>| <span data-ttu-id="185fd-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="185fd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="185fd-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185fd-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185fd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="185fd-108">1.0</span></span>|
|[<span data-ttu-id="185fd-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="185fd-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="185fd-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="185fd-110">ReadItem</span></span>|
|[<span data-ttu-id="185fd-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185fd-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="185fd-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="185fd-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="185fd-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="185fd-113">Members and methods</span></span>

| <span data-ttu-id="185fd-114">Membre</span><span class="sxs-lookup"><span data-stu-id="185fd-114">Member</span></span> | <span data-ttu-id="185fd-115">Type</span><span class="sxs-lookup"><span data-stu-id="185fd-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="185fd-116">accountType</span><span class="sxs-lookup"><span data-stu-id="185fd-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="185fd-117">Member</span><span class="sxs-lookup"><span data-stu-id="185fd-117">Member</span></span> |
| [<span data-ttu-id="185fd-118">displayName</span><span class="sxs-lookup"><span data-stu-id="185fd-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="185fd-119">Member</span><span class="sxs-lookup"><span data-stu-id="185fd-119">Member</span></span> |
| [<span data-ttu-id="185fd-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="185fd-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="185fd-121">Member</span><span class="sxs-lookup"><span data-stu-id="185fd-121">Member</span></span> |
| [<span data-ttu-id="185fd-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="185fd-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="185fd-123">Membre</span><span class="sxs-lookup"><span data-stu-id="185fd-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="185fd-124">Membres</span><span class="sxs-lookup"><span data-stu-id="185fd-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="185fd-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="185fd-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="185fd-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="185fd-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="185fd-127">Obtient le type de compte de l'utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="185fd-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="185fd-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="185fd-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="185fd-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="185fd-129">Value</span></span> | <span data-ttu-id="185fd-130">Description</span><span class="sxs-lookup"><span data-stu-id="185fd-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="185fd-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="185fd-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="185fd-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="185fd-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="185fd-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="185fd-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="185fd-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="185fd-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="185fd-135">Type</span><span class="sxs-lookup"><span data-stu-id="185fd-135">Type</span></span>

*   <span data-ttu-id="185fd-136">String</span><span class="sxs-lookup"><span data-stu-id="185fd-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="185fd-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185fd-137">Requirements</span></span>

|<span data-ttu-id="185fd-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185fd-138">Requirement</span></span>| <span data-ttu-id="185fd-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="185fd-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="185fd-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185fd-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185fd-141">1.6</span><span class="sxs-lookup"><span data-stu-id="185fd-141">1.6</span></span> |
|[<span data-ttu-id="185fd-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="185fd-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="185fd-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="185fd-143">ReadItem</span></span>|
|[<span data-ttu-id="185fd-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185fd-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="185fd-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="185fd-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="185fd-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="185fd-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="185fd-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="185fd-147">displayName :String</span></span>

<span data-ttu-id="185fd-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="185fd-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="185fd-149">Type</span><span class="sxs-lookup"><span data-stu-id="185fd-149">Type</span></span>

*   <span data-ttu-id="185fd-150">String</span><span class="sxs-lookup"><span data-stu-id="185fd-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="185fd-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185fd-151">Requirements</span></span>

|<span data-ttu-id="185fd-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185fd-152">Requirement</span></span>| <span data-ttu-id="185fd-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="185fd-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="185fd-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185fd-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185fd-155">1.0</span><span class="sxs-lookup"><span data-stu-id="185fd-155">1.0</span></span>|
|[<span data-ttu-id="185fd-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="185fd-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="185fd-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="185fd-157">ReadItem</span></span>|
|[<span data-ttu-id="185fd-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185fd-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="185fd-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="185fd-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="185fd-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="185fd-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="185fd-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="185fd-161">emailAddress :String</span></span>

<span data-ttu-id="185fd-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="185fd-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="185fd-163">Type</span><span class="sxs-lookup"><span data-stu-id="185fd-163">Type</span></span>

*   <span data-ttu-id="185fd-164">String</span><span class="sxs-lookup"><span data-stu-id="185fd-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="185fd-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185fd-165">Requirements</span></span>

|<span data-ttu-id="185fd-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185fd-166">Requirement</span></span>| <span data-ttu-id="185fd-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="185fd-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="185fd-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185fd-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185fd-169">1.0</span><span class="sxs-lookup"><span data-stu-id="185fd-169">1.0</span></span>|
|[<span data-ttu-id="185fd-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="185fd-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="185fd-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="185fd-171">ReadItem</span></span>|
|[<span data-ttu-id="185fd-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185fd-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="185fd-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="185fd-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="185fd-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="185fd-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="185fd-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="185fd-175">timeZone :String</span></span>

<span data-ttu-id="185fd-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="185fd-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="185fd-177">Type</span><span class="sxs-lookup"><span data-stu-id="185fd-177">Type</span></span>

*   <span data-ttu-id="185fd-178">String</span><span class="sxs-lookup"><span data-stu-id="185fd-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="185fd-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185fd-179">Requirements</span></span>

|<span data-ttu-id="185fd-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185fd-180">Requirement</span></span>| <span data-ttu-id="185fd-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="185fd-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="185fd-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185fd-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185fd-183">1.0</span><span class="sxs-lookup"><span data-stu-id="185fd-183">1.0</span></span>|
|[<span data-ttu-id="185fd-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="185fd-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="185fd-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="185fd-185">ReadItem</span></span>|
|[<span data-ttu-id="185fd-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185fd-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="185fd-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="185fd-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="185fd-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="185fd-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
