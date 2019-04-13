---
title: Office. Context. Mailbox. userProfile-aperçu de l'ensemble de conditions requises
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 29111314f16bb9c6518b350254a3036ffa125796
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838437"
---
# <a name="userprofile"></a><span data-ttu-id="75813-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="75813-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="75813-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="75813-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="75813-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75813-104">Requirements</span></span>

|<span data-ttu-id="75813-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75813-105">Requirement</span></span>| <span data-ttu-id="75813-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="75813-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="75813-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75813-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75813-108">1.0</span><span class="sxs-lookup"><span data-stu-id="75813-108">1.0</span></span>|
|[<span data-ttu-id="75813-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75813-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75813-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75813-110">ReadItem</span></span>|
|[<span data-ttu-id="75813-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75813-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75813-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="75813-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="75813-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="75813-113">Members and methods</span></span>

| <span data-ttu-id="75813-114">Membre</span><span class="sxs-lookup"><span data-stu-id="75813-114">Member</span></span> | <span data-ttu-id="75813-115">Type</span><span class="sxs-lookup"><span data-stu-id="75813-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="75813-116">accountType</span><span class="sxs-lookup"><span data-stu-id="75813-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="75813-117">Member</span><span class="sxs-lookup"><span data-stu-id="75813-117">Member</span></span> |
| [<span data-ttu-id="75813-118">displayName</span><span class="sxs-lookup"><span data-stu-id="75813-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="75813-119">Member</span><span class="sxs-lookup"><span data-stu-id="75813-119">Member</span></span> |
| [<span data-ttu-id="75813-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="75813-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="75813-121">Member</span><span class="sxs-lookup"><span data-stu-id="75813-121">Member</span></span> |
| [<span data-ttu-id="75813-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="75813-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="75813-123">Membre</span><span class="sxs-lookup"><span data-stu-id="75813-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="75813-124">Membres</span><span class="sxs-lookup"><span data-stu-id="75813-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="75813-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="75813-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="75813-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="75813-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="75813-127">Obtient le type de compte de l'utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="75813-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="75813-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="75813-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="75813-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="75813-129">Value</span></span> | <span data-ttu-id="75813-130">Description</span><span class="sxs-lookup"><span data-stu-id="75813-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="75813-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="75813-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="75813-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="75813-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="75813-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="75813-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="75813-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="75813-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="75813-135">Type</span><span class="sxs-lookup"><span data-stu-id="75813-135">Type</span></span>

*   <span data-ttu-id="75813-136">String</span><span class="sxs-lookup"><span data-stu-id="75813-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75813-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75813-137">Requirements</span></span>

|<span data-ttu-id="75813-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75813-138">Requirement</span></span>| <span data-ttu-id="75813-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="75813-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="75813-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75813-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75813-141">1.6</span><span class="sxs-lookup"><span data-stu-id="75813-141">1.6</span></span> |
|[<span data-ttu-id="75813-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75813-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75813-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75813-143">ReadItem</span></span>|
|[<span data-ttu-id="75813-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75813-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75813-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="75813-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75813-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="75813-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="75813-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="75813-147">displayName :String</span></span>

<span data-ttu-id="75813-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="75813-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="75813-149">Type</span><span class="sxs-lookup"><span data-stu-id="75813-149">Type</span></span>

*   <span data-ttu-id="75813-150">String</span><span class="sxs-lookup"><span data-stu-id="75813-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75813-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75813-151">Requirements</span></span>

|<span data-ttu-id="75813-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75813-152">Requirement</span></span>| <span data-ttu-id="75813-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="75813-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="75813-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75813-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75813-155">1.0</span><span class="sxs-lookup"><span data-stu-id="75813-155">1.0</span></span>|
|[<span data-ttu-id="75813-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75813-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75813-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75813-157">ReadItem</span></span>|
|[<span data-ttu-id="75813-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75813-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75813-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="75813-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75813-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="75813-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="75813-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="75813-161">emailAddress :String</span></span>

<span data-ttu-id="75813-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="75813-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="75813-163">Type</span><span class="sxs-lookup"><span data-stu-id="75813-163">Type</span></span>

*   <span data-ttu-id="75813-164">String</span><span class="sxs-lookup"><span data-stu-id="75813-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75813-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75813-165">Requirements</span></span>

|<span data-ttu-id="75813-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75813-166">Requirement</span></span>| <span data-ttu-id="75813-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="75813-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="75813-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75813-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75813-169">1.0</span><span class="sxs-lookup"><span data-stu-id="75813-169">1.0</span></span>|
|[<span data-ttu-id="75813-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75813-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75813-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75813-171">ReadItem</span></span>|
|[<span data-ttu-id="75813-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75813-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75813-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="75813-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75813-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="75813-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="75813-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="75813-175">timeZone :String</span></span>

<span data-ttu-id="75813-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="75813-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="75813-177">Type</span><span class="sxs-lookup"><span data-stu-id="75813-177">Type</span></span>

*   <span data-ttu-id="75813-178">String</span><span class="sxs-lookup"><span data-stu-id="75813-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75813-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75813-179">Requirements</span></span>

|<span data-ttu-id="75813-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75813-180">Requirement</span></span>| <span data-ttu-id="75813-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="75813-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="75813-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75813-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75813-183">1.0</span><span class="sxs-lookup"><span data-stu-id="75813-183">1.0</span></span>|
|[<span data-ttu-id="75813-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75813-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75813-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75813-185">ReadItem</span></span>|
|[<span data-ttu-id="75813-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75813-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75813-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="75813-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75813-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="75813-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
