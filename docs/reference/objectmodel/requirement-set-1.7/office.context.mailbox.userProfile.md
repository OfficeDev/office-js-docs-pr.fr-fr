---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,7
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 036f18e4cb98cfe510a19d85a5a79f393ca8bd17
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353292"
---
# <a name="userprofile"></a><span data-ttu-id="9695f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="9695f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="9695f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="9695f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="9695f-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9695f-104">Requirements</span></span>

|<span data-ttu-id="9695f-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9695f-105">Requirement</span></span>| <span data-ttu-id="9695f-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="9695f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9695f-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9695f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9695f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9695f-108">1.0</span></span>|
|[<span data-ttu-id="9695f-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9695f-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9695f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9695f-110">ReadItem</span></span>|
|[<span data-ttu-id="9695f-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9695f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9695f-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9695f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9695f-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="9695f-113">Members and methods</span></span>

| <span data-ttu-id="9695f-114">Membre</span><span class="sxs-lookup"><span data-stu-id="9695f-114">Member</span></span> | <span data-ttu-id="9695f-115">Type</span><span class="sxs-lookup"><span data-stu-id="9695f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9695f-116">accountType</span><span class="sxs-lookup"><span data-stu-id="9695f-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="9695f-117">Member</span><span class="sxs-lookup"><span data-stu-id="9695f-117">Member</span></span> |
| [<span data-ttu-id="9695f-118">displayName</span><span class="sxs-lookup"><span data-stu-id="9695f-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="9695f-119">Member</span><span class="sxs-lookup"><span data-stu-id="9695f-119">Member</span></span> |
| [<span data-ttu-id="9695f-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="9695f-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="9695f-121">Member</span><span class="sxs-lookup"><span data-stu-id="9695f-121">Member</span></span> |
| [<span data-ttu-id="9695f-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="9695f-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="9695f-123">Membre</span><span class="sxs-lookup"><span data-stu-id="9695f-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="9695f-124">Membres</span><span class="sxs-lookup"><span data-stu-id="9695f-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="9695f-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="9695f-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="9695f-126">Actuellement, ce membre est uniquement pris en charge par Outlook 2016 pour Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="9695f-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="9695f-127">Obtient le type de compte de l'utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="9695f-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="9695f-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="9695f-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="9695f-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="9695f-129">Value</span></span> | <span data-ttu-id="9695f-130">Description</span><span class="sxs-lookup"><span data-stu-id="9695f-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="9695f-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="9695f-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="9695f-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="9695f-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="9695f-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="9695f-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="9695f-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="9695f-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="9695f-135">Type</span><span class="sxs-lookup"><span data-stu-id="9695f-135">Type</span></span>

*   <span data-ttu-id="9695f-136">String</span><span class="sxs-lookup"><span data-stu-id="9695f-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9695f-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9695f-137">Requirements</span></span>

|<span data-ttu-id="9695f-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9695f-138">Requirement</span></span>| <span data-ttu-id="9695f-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="9695f-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="9695f-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9695f-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9695f-141">1.6</span><span class="sxs-lookup"><span data-stu-id="9695f-141">1.6</span></span> |
|[<span data-ttu-id="9695f-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9695f-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9695f-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9695f-143">ReadItem</span></span>|
|[<span data-ttu-id="9695f-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9695f-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9695f-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9695f-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9695f-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="9695f-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="9695f-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="9695f-147">displayName: String</span></span>

<span data-ttu-id="9695f-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9695f-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="9695f-149">Type</span><span class="sxs-lookup"><span data-stu-id="9695f-149">Type</span></span>

*   <span data-ttu-id="9695f-150">String</span><span class="sxs-lookup"><span data-stu-id="9695f-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9695f-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9695f-151">Requirements</span></span>

|<span data-ttu-id="9695f-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9695f-152">Requirement</span></span>| <span data-ttu-id="9695f-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="9695f-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="9695f-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9695f-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9695f-155">1.0</span><span class="sxs-lookup"><span data-stu-id="9695f-155">1.0</span></span>|
|[<span data-ttu-id="9695f-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9695f-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9695f-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9695f-157">ReadItem</span></span>|
|[<span data-ttu-id="9695f-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9695f-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9695f-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9695f-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9695f-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="9695f-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="9695f-161">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="9695f-161">emailAddress: String</span></span>

<span data-ttu-id="9695f-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9695f-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="9695f-163">Type</span><span class="sxs-lookup"><span data-stu-id="9695f-163">Type</span></span>

*   <span data-ttu-id="9695f-164">String</span><span class="sxs-lookup"><span data-stu-id="9695f-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9695f-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9695f-165">Requirements</span></span>

|<span data-ttu-id="9695f-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9695f-166">Requirement</span></span>| <span data-ttu-id="9695f-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="9695f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9695f-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9695f-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9695f-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9695f-169">1.0</span></span>|
|[<span data-ttu-id="9695f-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9695f-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9695f-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9695f-171">ReadItem</span></span>|
|[<span data-ttu-id="9695f-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9695f-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9695f-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9695f-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9695f-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="9695f-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="9695f-175">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="9695f-175">timeZone: String</span></span>

<span data-ttu-id="9695f-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9695f-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="9695f-177">Type</span><span class="sxs-lookup"><span data-stu-id="9695f-177">Type</span></span>

*   <span data-ttu-id="9695f-178">String</span><span class="sxs-lookup"><span data-stu-id="9695f-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9695f-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9695f-179">Requirements</span></span>

|<span data-ttu-id="9695f-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9695f-180">Requirement</span></span>| <span data-ttu-id="9695f-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="9695f-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="9695f-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9695f-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9695f-183">1.0</span><span class="sxs-lookup"><span data-stu-id="9695f-183">1.0</span></span>|
|[<span data-ttu-id="9695f-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9695f-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9695f-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9695f-185">ReadItem</span></span>|
|[<span data-ttu-id="9695f-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9695f-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9695f-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9695f-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9695f-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="9695f-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
