---
title: Office. Context. Mailbox. userProfile-aperçu de l’ensemble de conditions requises
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5941c4e1276535091a3ffcf5b2fb6aa972ed8c4d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696469"
---
# <a name="userprofile"></a><span data-ttu-id="e9e60-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e9e60-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e9e60-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e9e60-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e9e60-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e9e60-104">Requirements</span></span>

|<span data-ttu-id="e9e60-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e9e60-105">Requirement</span></span>| <span data-ttu-id="e9e60-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="e9e60-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e9e60-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e9e60-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e9e60-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e9e60-108">1.0</span></span>|
|[<span data-ttu-id="e9e60-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e9e60-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e9e60-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e9e60-110">ReadItem</span></span>|
|[<span data-ttu-id="e9e60-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e9e60-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e9e60-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e9e60-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e9e60-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="e9e60-113">Members and methods</span></span>

| <span data-ttu-id="e9e60-114">Membre</span><span class="sxs-lookup"><span data-stu-id="e9e60-114">Member</span></span> | <span data-ttu-id="e9e60-115">Type</span><span class="sxs-lookup"><span data-stu-id="e9e60-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e9e60-116">accountType</span><span class="sxs-lookup"><span data-stu-id="e9e60-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="e9e60-117">Member</span><span class="sxs-lookup"><span data-stu-id="e9e60-117">Member</span></span> |
| [<span data-ttu-id="e9e60-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e9e60-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e9e60-119">Member</span><span class="sxs-lookup"><span data-stu-id="e9e60-119">Member</span></span> |
| [<span data-ttu-id="e9e60-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e9e60-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e9e60-121">Member</span><span class="sxs-lookup"><span data-stu-id="e9e60-121">Member</span></span> |
| [<span data-ttu-id="e9e60-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e9e60-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e9e60-123">Membre</span><span class="sxs-lookup"><span data-stu-id="e9e60-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e9e60-124">Membres</span><span class="sxs-lookup"><span data-stu-id="e9e60-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="e9e60-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="e9e60-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="e9e60-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure sur Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="e9e60-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="e9e60-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e9e60-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e9e60-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="e9e60-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="e9e60-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="e9e60-129">Value</span></span> | <span data-ttu-id="e9e60-130">Description</span><span class="sxs-lookup"><span data-stu-id="e9e60-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e9e60-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="e9e60-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e9e60-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="e9e60-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e9e60-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="e9e60-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e9e60-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="e9e60-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e9e60-135">Type</span><span class="sxs-lookup"><span data-stu-id="e9e60-135">Type</span></span>

*   <span data-ttu-id="e9e60-136">String</span><span class="sxs-lookup"><span data-stu-id="e9e60-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e9e60-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e9e60-137">Requirements</span></span>

|<span data-ttu-id="e9e60-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e9e60-138">Requirement</span></span>| <span data-ttu-id="e9e60-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="e9e60-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e9e60-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e9e60-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e9e60-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e9e60-141">1.6</span></span> |
|[<span data-ttu-id="e9e60-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e9e60-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e9e60-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e9e60-143">ReadItem</span></span>|
|[<span data-ttu-id="e9e60-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e9e60-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e9e60-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e9e60-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e9e60-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="e9e60-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="e9e60-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="e9e60-147">displayName: String</span></span>

<span data-ttu-id="e9e60-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e9e60-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e9e60-149">Type</span><span class="sxs-lookup"><span data-stu-id="e9e60-149">Type</span></span>

*   <span data-ttu-id="e9e60-150">String</span><span class="sxs-lookup"><span data-stu-id="e9e60-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e9e60-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e9e60-151">Requirements</span></span>

|<span data-ttu-id="e9e60-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e9e60-152">Requirement</span></span>| <span data-ttu-id="e9e60-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="e9e60-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e9e60-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e9e60-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e9e60-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e9e60-155">1.0</span></span>|
|[<span data-ttu-id="e9e60-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e9e60-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e9e60-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e9e60-157">ReadItem</span></span>|
|[<span data-ttu-id="e9e60-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e9e60-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e9e60-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e9e60-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e9e60-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="e9e60-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="e9e60-161">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="e9e60-161">emailAddress: String</span></span>

<span data-ttu-id="e9e60-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e9e60-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e9e60-163">Type</span><span class="sxs-lookup"><span data-stu-id="e9e60-163">Type</span></span>

*   <span data-ttu-id="e9e60-164">String</span><span class="sxs-lookup"><span data-stu-id="e9e60-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e9e60-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e9e60-165">Requirements</span></span>

|<span data-ttu-id="e9e60-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e9e60-166">Requirement</span></span>| <span data-ttu-id="e9e60-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="e9e60-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e9e60-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e9e60-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e9e60-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e9e60-169">1.0</span></span>|
|[<span data-ttu-id="e9e60-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e9e60-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e9e60-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e9e60-171">ReadItem</span></span>|
|[<span data-ttu-id="e9e60-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e9e60-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e9e60-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e9e60-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e9e60-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="e9e60-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="e9e60-175">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="e9e60-175">timeZone: String</span></span>

<span data-ttu-id="e9e60-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e9e60-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e9e60-177">Type</span><span class="sxs-lookup"><span data-stu-id="e9e60-177">Type</span></span>

*   <span data-ttu-id="e9e60-178">String</span><span class="sxs-lookup"><span data-stu-id="e9e60-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e9e60-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e9e60-179">Requirements</span></span>

|<span data-ttu-id="e9e60-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e9e60-180">Requirement</span></span>| <span data-ttu-id="e9e60-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="e9e60-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e9e60-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e9e60-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e9e60-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e9e60-183">1.0</span></span>|
|[<span data-ttu-id="e9e60-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e9e60-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e9e60-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e9e60-185">ReadItem</span></span>|
|[<span data-ttu-id="e9e60-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e9e60-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e9e60-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e9e60-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e9e60-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="e9e60-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
