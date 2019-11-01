---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,8
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 39a833a81eab22c70d89cdfc61784555312b23d6
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902164"
---
# <a name="userprofile"></a><span data-ttu-id="a1e4f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a1e4f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a1e4f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a1e4f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1e4f-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a1e4f-104">Requirements</span></span>

|<span data-ttu-id="a1e4f-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a1e4f-105">Requirement</span></span>| <span data-ttu-id="a1e4f-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="a1e4f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1e4f-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a1e4f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1e4f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a1e4f-108">1.0</span></span>|
|[<span data-ttu-id="a1e4f-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a1e4f-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1e4f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1e4f-110">ReadItem</span></span>|
|[<span data-ttu-id="a1e4f-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a1e4f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1e4f-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a1e4f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a1e4f-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a1e4f-113">Members and methods</span></span>

| <span data-ttu-id="a1e4f-114">Membre</span><span class="sxs-lookup"><span data-stu-id="a1e4f-114">Member</span></span> | <span data-ttu-id="a1e4f-115">Type</span><span class="sxs-lookup"><span data-stu-id="a1e4f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a1e4f-116">accountType</span><span class="sxs-lookup"><span data-stu-id="a1e4f-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="a1e4f-117">Membre</span><span class="sxs-lookup"><span data-stu-id="a1e4f-117">Member</span></span> |
| [<span data-ttu-id="a1e4f-118">displayName</span><span class="sxs-lookup"><span data-stu-id="a1e4f-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="a1e4f-119">Membre</span><span class="sxs-lookup"><span data-stu-id="a1e4f-119">Member</span></span> |
| [<span data-ttu-id="a1e4f-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a1e4f-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a1e4f-121">Membre</span><span class="sxs-lookup"><span data-stu-id="a1e4f-121">Member</span></span> |
| [<span data-ttu-id="a1e4f-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="a1e4f-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a1e4f-123">Membre</span><span class="sxs-lookup"><span data-stu-id="a1e4f-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a1e4f-124">Membres</span><span class="sxs-lookup"><span data-stu-id="a1e4f-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="a1e4f-125">accountType : chaîne</span><span class="sxs-lookup"><span data-stu-id="a1e4f-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="a1e4f-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure sur Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="a1e4f-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="a1e4f-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="a1e4f-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="a1e4f-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="a1e4f-129">Value</span></span> | <span data-ttu-id="a1e4f-130">Description</span><span class="sxs-lookup"><span data-stu-id="a1e4f-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="a1e4f-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="a1e4f-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="a1e4f-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="a1e4f-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="a1e4f-135">Type</span><span class="sxs-lookup"><span data-stu-id="a1e4f-135">Type</span></span>

*   <span data-ttu-id="a1e4f-136">String</span><span class="sxs-lookup"><span data-stu-id="a1e4f-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1e4f-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a1e4f-137">Requirements</span></span>

|<span data-ttu-id="a1e4f-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a1e4f-138">Requirement</span></span>| <span data-ttu-id="a1e4f-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="a1e4f-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1e4f-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a1e4f-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1e4f-141">1.6</span><span class="sxs-lookup"><span data-stu-id="a1e4f-141">1.6</span></span> |
|[<span data-ttu-id="a1e4f-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a1e4f-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1e4f-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1e4f-143">ReadItem</span></span>|
|[<span data-ttu-id="a1e4f-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a1e4f-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1e4f-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a1e4f-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1e4f-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="a1e4f-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="a1e4f-147">displayName : String</span><span class="sxs-lookup"><span data-stu-id="a1e4f-147">displayName: String</span></span>

<span data-ttu-id="a1e4f-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a1e4f-149">Type</span><span class="sxs-lookup"><span data-stu-id="a1e4f-149">Type</span></span>

*   <span data-ttu-id="a1e4f-150">String</span><span class="sxs-lookup"><span data-stu-id="a1e4f-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1e4f-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a1e4f-151">Requirements</span></span>

|<span data-ttu-id="a1e4f-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a1e4f-152">Requirement</span></span>| <span data-ttu-id="a1e4f-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="a1e4f-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1e4f-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a1e4f-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1e4f-155">1.0</span><span class="sxs-lookup"><span data-stu-id="a1e4f-155">1.0</span></span>|
|[<span data-ttu-id="a1e4f-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a1e4f-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1e4f-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1e4f-157">ReadItem</span></span>|
|[<span data-ttu-id="a1e4f-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a1e4f-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1e4f-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a1e4f-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1e4f-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="a1e4f-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="a1e4f-161">emailAddress : chaîne</span><span class="sxs-lookup"><span data-stu-id="a1e4f-161">emailAddress: String</span></span>

<span data-ttu-id="a1e4f-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a1e4f-163">Type</span><span class="sxs-lookup"><span data-stu-id="a1e4f-163">Type</span></span>

*   <span data-ttu-id="a1e4f-164">String</span><span class="sxs-lookup"><span data-stu-id="a1e4f-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1e4f-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a1e4f-165">Requirements</span></span>

|<span data-ttu-id="a1e4f-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a1e4f-166">Requirement</span></span>| <span data-ttu-id="a1e4f-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="a1e4f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1e4f-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a1e4f-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1e4f-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a1e4f-169">1.0</span></span>|
|[<span data-ttu-id="a1e4f-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a1e4f-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1e4f-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1e4f-171">ReadItem</span></span>|
|[<span data-ttu-id="a1e4f-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a1e4f-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1e4f-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a1e4f-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1e4f-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="a1e4f-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="a1e4f-175">timeZone : chaîne</span><span class="sxs-lookup"><span data-stu-id="a1e4f-175">timeZone: String</span></span>

<span data-ttu-id="a1e4f-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a1e4f-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a1e4f-177">Type</span><span class="sxs-lookup"><span data-stu-id="a1e4f-177">Type</span></span>

*   <span data-ttu-id="a1e4f-178">String</span><span class="sxs-lookup"><span data-stu-id="a1e4f-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1e4f-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a1e4f-179">Requirements</span></span>

|<span data-ttu-id="a1e4f-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a1e4f-180">Requirement</span></span>| <span data-ttu-id="a1e4f-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="a1e4f-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1e4f-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a1e4f-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1e4f-183">1.0</span><span class="sxs-lookup"><span data-stu-id="a1e4f-183">1.0</span></span>|
|[<span data-ttu-id="a1e4f-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a1e4f-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1e4f-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1e4f-185">ReadItem</span></span>|
|[<span data-ttu-id="a1e4f-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a1e4f-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1e4f-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a1e4f-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1e4f-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="a1e4f-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
