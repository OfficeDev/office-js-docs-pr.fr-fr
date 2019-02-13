---
title: Office.context.mailbox.userProfile-ensemble de conditions requises 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29635901"
---
# <a name="userprofile"></a><span data-ttu-id="8d08e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8d08e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8d08e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8d08e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d08e-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8d08e-104">Requirements</span></span>

|<span data-ttu-id="8d08e-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8d08e-105">Requirement</span></span>| <span data-ttu-id="8d08e-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="8d08e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d08e-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8d08e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d08e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8d08e-108">1.0</span></span>|
|[<span data-ttu-id="8d08e-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8d08e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d08e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d08e-110">ReadItem</span></span>|
|[<span data-ttu-id="8d08e-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8d08e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8d08e-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8d08e-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8d08e-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="8d08e-113">Members and methods</span></span>

| <span data-ttu-id="8d08e-114">Membre</span><span class="sxs-lookup"><span data-stu-id="8d08e-114">Member</span></span> | <span data-ttu-id="8d08e-115">Type</span><span class="sxs-lookup"><span data-stu-id="8d08e-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8d08e-116">accountType</span><span class="sxs-lookup"><span data-stu-id="8d08e-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="8d08e-117">Membre</span><span class="sxs-lookup"><span data-stu-id="8d08e-117">Member</span></span> |
| [<span data-ttu-id="8d08e-118">displayName</span><span class="sxs-lookup"><span data-stu-id="8d08e-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="8d08e-119">Membre</span><span class="sxs-lookup"><span data-stu-id="8d08e-119">Member</span></span> |
| [<span data-ttu-id="8d08e-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="8d08e-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="8d08e-121">Membre</span><span class="sxs-lookup"><span data-stu-id="8d08e-121">Member</span></span> |
| [<span data-ttu-id="8d08e-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="8d08e-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="8d08e-123">Membre</span><span class="sxs-lookup"><span data-stu-id="8d08e-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8d08e-124">Members</span><span class="sxs-lookup"><span data-stu-id="8d08e-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="8d08e-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="8d08e-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="8d08e-126">Ce membre est uniquement pris en charge par 2016 Outlook pour Mac (build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="8d08e-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="8d08e-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="8d08e-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="8d08e-128">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="8d08e-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="8d08e-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="8d08e-129">Value</span></span> | <span data-ttu-id="8d08e-130">Description</span><span class="sxs-lookup"><span data-stu-id="8d08e-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="8d08e-131">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="8d08e-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="8d08e-132">La boîte aux lettres est associée à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="8d08e-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="8d08e-133">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="8d08e-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="8d08e-134">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="8d08e-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="8d08e-135">Type :</span><span class="sxs-lookup"><span data-stu-id="8d08e-135">Type:</span></span>

*   <span data-ttu-id="8d08e-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8d08e-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d08e-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8d08e-137">Requirements</span></span>

|<span data-ttu-id="8d08e-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8d08e-138">Requirement</span></span>| <span data-ttu-id="8d08e-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="8d08e-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d08e-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8d08e-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d08e-141">1.6</span><span class="sxs-lookup"><span data-stu-id="8d08e-141">1.6</span></span> |
|[<span data-ttu-id="8d08e-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8d08e-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d08e-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d08e-143">ReadItem</span></span>|
|[<span data-ttu-id="8d08e-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8d08e-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8d08e-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8d08e-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d08e-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="8d08e-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="8d08e-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="8d08e-147">displayName :String</span></span>

<span data-ttu-id="8d08e-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8d08e-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8d08e-149">Type :</span><span class="sxs-lookup"><span data-stu-id="8d08e-149">Type:</span></span>

*   <span data-ttu-id="8d08e-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8d08e-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d08e-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8d08e-151">Requirements</span></span>

|<span data-ttu-id="8d08e-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8d08e-152">Requirement</span></span>| <span data-ttu-id="8d08e-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="8d08e-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d08e-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8d08e-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d08e-155">1.0</span><span class="sxs-lookup"><span data-stu-id="8d08e-155">1.0</span></span>|
|[<span data-ttu-id="8d08e-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8d08e-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d08e-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d08e-157">ReadItem</span></span>|
|[<span data-ttu-id="8d08e-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8d08e-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8d08e-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8d08e-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d08e-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="8d08e-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="8d08e-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="8d08e-161">emailAddress :String</span></span>

<span data-ttu-id="8d08e-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8d08e-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8d08e-163">Type :</span><span class="sxs-lookup"><span data-stu-id="8d08e-163">Type:</span></span>

*   <span data-ttu-id="8d08e-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8d08e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d08e-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8d08e-165">Requirements</span></span>

|<span data-ttu-id="8d08e-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8d08e-166">Requirement</span></span>| <span data-ttu-id="8d08e-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="8d08e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d08e-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8d08e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d08e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="8d08e-169">1.0</span></span>|
|[<span data-ttu-id="8d08e-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8d08e-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d08e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d08e-171">ReadItem</span></span>|
|[<span data-ttu-id="8d08e-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8d08e-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8d08e-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8d08e-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d08e-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="8d08e-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="8d08e-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="8d08e-175">timeZone :String</span></span>

<span data-ttu-id="8d08e-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8d08e-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8d08e-177">Type :</span><span class="sxs-lookup"><span data-stu-id="8d08e-177">Type:</span></span>

*   <span data-ttu-id="8d08e-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8d08e-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d08e-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8d08e-179">Requirements</span></span>

|<span data-ttu-id="8d08e-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8d08e-180">Requirement</span></span>| <span data-ttu-id="8d08e-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="8d08e-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d08e-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8d08e-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d08e-183">1.0</span><span class="sxs-lookup"><span data-stu-id="8d08e-183">1.0</span></span>|
|[<span data-ttu-id="8d08e-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8d08e-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d08e-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d08e-185">ReadItem</span></span>|
|[<span data-ttu-id="8d08e-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8d08e-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8d08e-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8d08e-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d08e-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="8d08e-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
