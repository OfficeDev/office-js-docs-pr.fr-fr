---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,7
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 335c20a242d08031e6f21c48e99c8527dab6d714
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871737"
---
# <a name="userprofile"></a><span data-ttu-id="87dd1-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="87dd1-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="87dd1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="87dd1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="87dd1-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="87dd1-104">Requirements</span></span>

|<span data-ttu-id="87dd1-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="87dd1-105">Requirement</span></span>| <span data-ttu-id="87dd1-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="87dd1-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="87dd1-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="87dd1-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87dd1-108">1.0</span><span class="sxs-lookup"><span data-stu-id="87dd1-108">1.0</span></span>|
|[<span data-ttu-id="87dd1-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="87dd1-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87dd1-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87dd1-110">ReadItem</span></span>|
|[<span data-ttu-id="87dd1-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="87dd1-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87dd1-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="87dd1-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="87dd1-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="87dd1-113">Members and methods</span></span>

| <span data-ttu-id="87dd1-114">Membre</span><span class="sxs-lookup"><span data-stu-id="87dd1-114">Member</span></span> | <span data-ttu-id="87dd1-115">Type</span><span class="sxs-lookup"><span data-stu-id="87dd1-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="87dd1-116">accountType</span><span class="sxs-lookup"><span data-stu-id="87dd1-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="87dd1-117">Member</span><span class="sxs-lookup"><span data-stu-id="87dd1-117">Member</span></span> |
| [<span data-ttu-id="87dd1-118">displayName</span><span class="sxs-lookup"><span data-stu-id="87dd1-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="87dd1-119">Member</span><span class="sxs-lookup"><span data-stu-id="87dd1-119">Member</span></span> |
| [<span data-ttu-id="87dd1-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="87dd1-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="87dd1-121">Member</span><span class="sxs-lookup"><span data-stu-id="87dd1-121">Member</span></span> |
| [<span data-ttu-id="87dd1-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="87dd1-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="87dd1-123">Membre</span><span class="sxs-lookup"><span data-stu-id="87dd1-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="87dd1-124">Membres</span><span class="sxs-lookup"><span data-stu-id="87dd1-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="87dd1-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="87dd1-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="87dd1-126">Actuellement, ce membre est uniquement pris en charge par Outlook 2016 pour Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="87dd1-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="87dd1-127">Obtient le type de compte de l'utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="87dd1-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="87dd1-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="87dd1-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="87dd1-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="87dd1-129">Value</span></span> | <span data-ttu-id="87dd1-130">Description</span><span class="sxs-lookup"><span data-stu-id="87dd1-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="87dd1-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="87dd1-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="87dd1-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="87dd1-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="87dd1-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="87dd1-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="87dd1-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="87dd1-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="87dd1-135">Type</span><span class="sxs-lookup"><span data-stu-id="87dd1-135">Type</span></span>

*   <span data-ttu-id="87dd1-136">String</span><span class="sxs-lookup"><span data-stu-id="87dd1-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87dd1-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="87dd1-137">Requirements</span></span>

|<span data-ttu-id="87dd1-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="87dd1-138">Requirement</span></span>| <span data-ttu-id="87dd1-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="87dd1-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="87dd1-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="87dd1-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87dd1-141">1.6</span><span class="sxs-lookup"><span data-stu-id="87dd1-141">1.6</span></span> |
|[<span data-ttu-id="87dd1-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="87dd1-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87dd1-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87dd1-143">ReadItem</span></span>|
|[<span data-ttu-id="87dd1-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="87dd1-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87dd1-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="87dd1-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87dd1-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="87dd1-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="87dd1-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="87dd1-147">displayName :String</span></span>

<span data-ttu-id="87dd1-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="87dd1-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="87dd1-149">Type</span><span class="sxs-lookup"><span data-stu-id="87dd1-149">Type</span></span>

*   <span data-ttu-id="87dd1-150">String</span><span class="sxs-lookup"><span data-stu-id="87dd1-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87dd1-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="87dd1-151">Requirements</span></span>

|<span data-ttu-id="87dd1-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="87dd1-152">Requirement</span></span>| <span data-ttu-id="87dd1-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="87dd1-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="87dd1-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="87dd1-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87dd1-155">1.0</span><span class="sxs-lookup"><span data-stu-id="87dd1-155">1.0</span></span>|
|[<span data-ttu-id="87dd1-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="87dd1-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87dd1-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87dd1-157">ReadItem</span></span>|
|[<span data-ttu-id="87dd1-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="87dd1-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87dd1-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="87dd1-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87dd1-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="87dd1-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="87dd1-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="87dd1-161">emailAddress :String</span></span>

<span data-ttu-id="87dd1-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="87dd1-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="87dd1-163">Type</span><span class="sxs-lookup"><span data-stu-id="87dd1-163">Type</span></span>

*   <span data-ttu-id="87dd1-164">String</span><span class="sxs-lookup"><span data-stu-id="87dd1-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87dd1-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="87dd1-165">Requirements</span></span>

|<span data-ttu-id="87dd1-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="87dd1-166">Requirement</span></span>| <span data-ttu-id="87dd1-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="87dd1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="87dd1-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="87dd1-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87dd1-169">1.0</span><span class="sxs-lookup"><span data-stu-id="87dd1-169">1.0</span></span>|
|[<span data-ttu-id="87dd1-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="87dd1-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87dd1-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87dd1-171">ReadItem</span></span>|
|[<span data-ttu-id="87dd1-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="87dd1-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87dd1-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="87dd1-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87dd1-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="87dd1-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="87dd1-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="87dd1-175">timeZone :String</span></span>

<span data-ttu-id="87dd1-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="87dd1-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="87dd1-177">Type</span><span class="sxs-lookup"><span data-stu-id="87dd1-177">Type</span></span>

*   <span data-ttu-id="87dd1-178">String</span><span class="sxs-lookup"><span data-stu-id="87dd1-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87dd1-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="87dd1-179">Requirements</span></span>

|<span data-ttu-id="87dd1-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="87dd1-180">Requirement</span></span>| <span data-ttu-id="87dd1-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="87dd1-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="87dd1-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="87dd1-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87dd1-183">1.0</span><span class="sxs-lookup"><span data-stu-id="87dd1-183">1.0</span></span>|
|[<span data-ttu-id="87dd1-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="87dd1-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87dd1-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87dd1-185">ReadItem</span></span>|
|[<span data-ttu-id="87dd1-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="87dd1-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87dd1-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="87dd1-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87dd1-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="87dd1-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
