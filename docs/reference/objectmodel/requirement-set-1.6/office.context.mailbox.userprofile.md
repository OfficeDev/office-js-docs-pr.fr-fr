---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,6
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 3ca06925dcd37d8e68f086daf4705b10fb936623
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127204"
---
# <a name="userprofile"></a><span data-ttu-id="50c33-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="50c33-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="50c33-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="50c33-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="50c33-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="50c33-104">Requirements</span></span>

|<span data-ttu-id="50c33-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="50c33-105">Requirement</span></span>| <span data-ttu-id="50c33-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="50c33-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="50c33-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="50c33-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50c33-108">1.0</span><span class="sxs-lookup"><span data-stu-id="50c33-108">1.0</span></span>|
|[<span data-ttu-id="50c33-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="50c33-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="50c33-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="50c33-110">ReadItem</span></span>|
|[<span data-ttu-id="50c33-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="50c33-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="50c33-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="50c33-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="50c33-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="50c33-113">Members and methods</span></span>

| <span data-ttu-id="50c33-114">Membre</span><span class="sxs-lookup"><span data-stu-id="50c33-114">Member</span></span> | <span data-ttu-id="50c33-115">Type</span><span class="sxs-lookup"><span data-stu-id="50c33-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="50c33-116">accountType</span><span class="sxs-lookup"><span data-stu-id="50c33-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="50c33-117">Member</span><span class="sxs-lookup"><span data-stu-id="50c33-117">Member</span></span> |
| [<span data-ttu-id="50c33-118">displayName</span><span class="sxs-lookup"><span data-stu-id="50c33-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="50c33-119">Member</span><span class="sxs-lookup"><span data-stu-id="50c33-119">Member</span></span> |
| [<span data-ttu-id="50c33-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="50c33-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="50c33-121">Member</span><span class="sxs-lookup"><span data-stu-id="50c33-121">Member</span></span> |
| [<span data-ttu-id="50c33-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="50c33-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="50c33-123">Membre</span><span class="sxs-lookup"><span data-stu-id="50c33-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="50c33-124">Membres</span><span class="sxs-lookup"><span data-stu-id="50c33-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="50c33-125">accountType: chaîne</span><span class="sxs-lookup"><span data-stu-id="50c33-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="50c33-126">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure sur Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="50c33-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="50c33-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="50c33-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="50c33-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="50c33-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="50c33-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="50c33-129">Value</span></span> | <span data-ttu-id="50c33-130">Description</span><span class="sxs-lookup"><span data-stu-id="50c33-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="50c33-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="50c33-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="50c33-132">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="50c33-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="50c33-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="50c33-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="50c33-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="50c33-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="50c33-135">Type</span><span class="sxs-lookup"><span data-stu-id="50c33-135">Type</span></span>

*   <span data-ttu-id="50c33-136">String</span><span class="sxs-lookup"><span data-stu-id="50c33-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="50c33-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="50c33-137">Requirements</span></span>

|<span data-ttu-id="50c33-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="50c33-138">Requirement</span></span>| <span data-ttu-id="50c33-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="50c33-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="50c33-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="50c33-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50c33-141">1.6</span><span class="sxs-lookup"><span data-stu-id="50c33-141">1.6</span></span> |
|[<span data-ttu-id="50c33-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="50c33-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="50c33-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="50c33-143">ReadItem</span></span>|
|[<span data-ttu-id="50c33-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="50c33-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="50c33-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="50c33-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="50c33-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="50c33-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

#### <a name="displayname-string"></a><span data-ttu-id="50c33-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="50c33-147">displayName: String</span></span>

<span data-ttu-id="50c33-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="50c33-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="50c33-149">Type</span><span class="sxs-lookup"><span data-stu-id="50c33-149">Type</span></span>

*   <span data-ttu-id="50c33-150">String</span><span class="sxs-lookup"><span data-stu-id="50c33-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="50c33-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="50c33-151">Requirements</span></span>

|<span data-ttu-id="50c33-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="50c33-152">Requirement</span></span>| <span data-ttu-id="50c33-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="50c33-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="50c33-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="50c33-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50c33-155">1.0</span><span class="sxs-lookup"><span data-stu-id="50c33-155">1.0</span></span>|
|[<span data-ttu-id="50c33-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="50c33-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="50c33-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="50c33-157">ReadItem</span></span>|
|[<span data-ttu-id="50c33-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="50c33-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="50c33-159">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="50c33-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="50c33-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="50c33-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="50c33-161">emailAddress: chaîne</span><span class="sxs-lookup"><span data-stu-id="50c33-161">emailAddress: String</span></span>

<span data-ttu-id="50c33-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="50c33-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="50c33-163">Type</span><span class="sxs-lookup"><span data-stu-id="50c33-163">Type</span></span>

*   <span data-ttu-id="50c33-164">String</span><span class="sxs-lookup"><span data-stu-id="50c33-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="50c33-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="50c33-165">Requirements</span></span>

|<span data-ttu-id="50c33-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="50c33-166">Requirement</span></span>| <span data-ttu-id="50c33-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="50c33-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="50c33-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="50c33-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50c33-169">1.0</span><span class="sxs-lookup"><span data-stu-id="50c33-169">1.0</span></span>|
|[<span data-ttu-id="50c33-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="50c33-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="50c33-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="50c33-171">ReadItem</span></span>|
|[<span data-ttu-id="50c33-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="50c33-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="50c33-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="50c33-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="50c33-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="50c33-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="50c33-175">timeZone: chaîne</span><span class="sxs-lookup"><span data-stu-id="50c33-175">timeZone: String</span></span>

<span data-ttu-id="50c33-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="50c33-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="50c33-177">Type</span><span class="sxs-lookup"><span data-stu-id="50c33-177">Type</span></span>

*   <span data-ttu-id="50c33-178">String</span><span class="sxs-lookup"><span data-stu-id="50c33-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="50c33-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="50c33-179">Requirements</span></span>

|<span data-ttu-id="50c33-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="50c33-180">Requirement</span></span>| <span data-ttu-id="50c33-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="50c33-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="50c33-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="50c33-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50c33-183">1.0</span><span class="sxs-lookup"><span data-stu-id="50c33-183">1.0</span></span>|
|[<span data-ttu-id="50c33-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="50c33-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="50c33-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="50c33-185">ReadItem</span></span>|
|[<span data-ttu-id="50c33-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="50c33-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="50c33-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="50c33-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="50c33-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="50c33-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
