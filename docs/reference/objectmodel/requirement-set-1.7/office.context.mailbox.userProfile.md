---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 8cfee874bbb5183d62cc3a9ce8b042a76617ec72
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451807"
---
# <a name="userprofile"></a><span data-ttu-id="c3248-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c3248-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c3248-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c3248-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3248-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3248-104">Requirements</span></span>

|<span data-ttu-id="c3248-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3248-105">Requirement</span></span>| <span data-ttu-id="c3248-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3248-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3248-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3248-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3248-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c3248-108">1.0</span></span>|
|[<span data-ttu-id="c3248-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c3248-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3248-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3248-110">ReadItem</span></span>|
|[<span data-ttu-id="c3248-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3248-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3248-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3248-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c3248-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c3248-113">Members and methods</span></span>

| <span data-ttu-id="c3248-114">Membre</span><span class="sxs-lookup"><span data-stu-id="c3248-114">Member</span></span> | <span data-ttu-id="c3248-115">Type</span><span class="sxs-lookup"><span data-stu-id="c3248-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c3248-116">accountType</span><span class="sxs-lookup"><span data-stu-id="c3248-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="c3248-117">Member</span><span class="sxs-lookup"><span data-stu-id="c3248-117">Member</span></span> |
| [<span data-ttu-id="c3248-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c3248-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c3248-119">Member</span><span class="sxs-lookup"><span data-stu-id="c3248-119">Member</span></span> |
| [<span data-ttu-id="c3248-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c3248-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c3248-121">Member</span><span class="sxs-lookup"><span data-stu-id="c3248-121">Member</span></span> |
| [<span data-ttu-id="c3248-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c3248-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c3248-123">Membre</span><span class="sxs-lookup"><span data-stu-id="c3248-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c3248-124">Membres</span><span class="sxs-lookup"><span data-stu-id="c3248-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c3248-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="c3248-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c3248-126">Actuellement, ce membre est uniquement pris en charge par Outlook 2016 pour Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="c3248-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="c3248-127">Obtient le type de compte de l'utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="c3248-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c3248-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="c3248-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="c3248-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3248-129">Value</span></span> | <span data-ttu-id="c3248-130">Description</span><span class="sxs-lookup"><span data-stu-id="c3248-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c3248-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="c3248-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c3248-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="c3248-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c3248-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="c3248-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c3248-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="c3248-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c3248-135">Type</span><span class="sxs-lookup"><span data-stu-id="c3248-135">Type</span></span>

*   <span data-ttu-id="c3248-136">String</span><span class="sxs-lookup"><span data-stu-id="c3248-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3248-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3248-137">Requirements</span></span>

|<span data-ttu-id="c3248-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3248-138">Requirement</span></span>| <span data-ttu-id="c3248-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3248-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3248-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3248-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3248-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c3248-141">1.6</span></span> |
|[<span data-ttu-id="c3248-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c3248-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3248-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3248-143">ReadItem</span></span>|
|[<span data-ttu-id="c3248-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3248-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3248-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3248-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3248-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3248-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="c3248-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c3248-147">displayName :String</span></span>

<span data-ttu-id="c3248-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c3248-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c3248-149">Type</span><span class="sxs-lookup"><span data-stu-id="c3248-149">Type</span></span>

*   <span data-ttu-id="c3248-150">String</span><span class="sxs-lookup"><span data-stu-id="c3248-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3248-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3248-151">Requirements</span></span>

|<span data-ttu-id="c3248-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3248-152">Requirement</span></span>| <span data-ttu-id="c3248-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3248-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3248-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3248-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3248-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c3248-155">1.0</span></span>|
|[<span data-ttu-id="c3248-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c3248-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3248-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3248-157">ReadItem</span></span>|
|[<span data-ttu-id="c3248-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3248-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3248-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3248-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3248-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3248-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="c3248-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c3248-161">emailAddress :String</span></span>

<span data-ttu-id="c3248-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c3248-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c3248-163">Type</span><span class="sxs-lookup"><span data-stu-id="c3248-163">Type</span></span>

*   <span data-ttu-id="c3248-164">String</span><span class="sxs-lookup"><span data-stu-id="c3248-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3248-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3248-165">Requirements</span></span>

|<span data-ttu-id="c3248-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3248-166">Requirement</span></span>| <span data-ttu-id="c3248-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3248-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3248-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3248-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3248-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c3248-169">1.0</span></span>|
|[<span data-ttu-id="c3248-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c3248-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3248-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3248-171">ReadItem</span></span>|
|[<span data-ttu-id="c3248-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3248-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3248-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3248-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3248-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3248-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="c3248-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c3248-175">timeZone :String</span></span>

<span data-ttu-id="c3248-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c3248-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c3248-177">Type</span><span class="sxs-lookup"><span data-stu-id="c3248-177">Type</span></span>

*   <span data-ttu-id="c3248-178">String</span><span class="sxs-lookup"><span data-stu-id="c3248-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3248-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3248-179">Requirements</span></span>

|<span data-ttu-id="c3248-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3248-180">Requirement</span></span>| <span data-ttu-id="c3248-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3248-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3248-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3248-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3248-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c3248-183">1.0</span></span>|
|[<span data-ttu-id="c3248-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c3248-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3248-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3248-185">ReadItem</span></span>|
|[<span data-ttu-id="c3248-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3248-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3248-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3248-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3248-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3248-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
