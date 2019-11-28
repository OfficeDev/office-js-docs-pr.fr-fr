---
title: Office. Context. Mailbox. userProfile-aperçu de l’ensemble de conditions requises
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 4afc64f247155576ab3f0024d1929a29a0f7dc0c
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629257"
---
# <a name="userprofile"></a><span data-ttu-id="7ac52-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="7ac52-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="7ac52-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="7ac52-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac52-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac52-104">Requirements</span></span>

|<span data-ttu-id="7ac52-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac52-105">Requirement</span></span>| <span data-ttu-id="7ac52-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac52-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac52-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac52-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ac52-108">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-108">1.0</span></span>|
|[<span data-ttu-id="7ac52-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7ac52-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7ac52-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-110">ReadItem</span></span>|
|[<span data-ttu-id="7ac52-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac52-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ac52-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="7ac52-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="7ac52-113">Properties</span></span>

| <span data-ttu-id="7ac52-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="7ac52-114">Property</span></span> | <span data-ttu-id="7ac52-115">Minimale</span><span class="sxs-lookup"><span data-stu-id="7ac52-115">Minimum</span></span><br><span data-ttu-id="7ac52-116">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="7ac52-116">permission level</span></span> | <span data-ttu-id="7ac52-117">Modes</span><span class="sxs-lookup"><span data-stu-id="7ac52-117">Modes</span></span> | <span data-ttu-id="7ac52-118">Type de retour</span><span class="sxs-lookup"><span data-stu-id="7ac52-118">Return type</span></span> | <span data-ttu-id="7ac52-119">Minimale</span><span class="sxs-lookup"><span data-stu-id="7ac52-119">Minimum</span></span><br><span data-ttu-id="7ac52-120">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac52-120">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="7ac52-121">accountType</span><span class="sxs-lookup"><span data-stu-id="7ac52-121">accountType</span></span>](#accounttype-string) | <span data-ttu-id="7ac52-122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-122">ReadItem</span></span> | <span data-ttu-id="7ac52-123">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac52-123">Compose</span></span><br><span data-ttu-id="7ac52-124">Lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-124">Read</span></span> | <span data-ttu-id="7ac52-125">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-125">String</span></span> | <span data-ttu-id="7ac52-126">1.6</span><span class="sxs-lookup"><span data-stu-id="7ac52-126">1.6</span></span> |
| [<span data-ttu-id="7ac52-127">displayName</span><span class="sxs-lookup"><span data-stu-id="7ac52-127">displayName</span></span>](#displayname-string) | <span data-ttu-id="7ac52-128">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-128">ReadItem</span></span> | <span data-ttu-id="7ac52-129">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac52-129">Compose</span></span><br><span data-ttu-id="7ac52-130">Lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-130">Read</span></span> | <span data-ttu-id="7ac52-131">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-131">String</span></span> | <span data-ttu-id="7ac52-132">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-132">1.0</span></span> |
| [<span data-ttu-id="7ac52-133">emailAddress</span><span class="sxs-lookup"><span data-stu-id="7ac52-133">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="7ac52-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-134">ReadItem</span></span> | <span data-ttu-id="7ac52-135">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac52-135">Compose</span></span><br><span data-ttu-id="7ac52-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-136">Read</span></span> | <span data-ttu-id="7ac52-137">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-137">String</span></span> | <span data-ttu-id="7ac52-138">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-138">1.0</span></span> |
| [<span data-ttu-id="7ac52-139">timeZone</span><span class="sxs-lookup"><span data-stu-id="7ac52-139">timeZone</span></span>](#timezone-string) | <span data-ttu-id="7ac52-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-140">ReadItem</span></span> | <span data-ttu-id="7ac52-141">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac52-141">Compose</span></span><br><span data-ttu-id="7ac52-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-142">Read</span></span> | <span data-ttu-id="7ac52-143">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-143">String</span></span> | <span data-ttu-id="7ac52-144">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-144">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="7ac52-145">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="7ac52-145">Property details</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="7ac52-146">accountType : chaîne</span><span class="sxs-lookup"><span data-stu-id="7ac52-146">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="7ac52-147">Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure sur Mac (Build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="7ac52-147">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="7ac52-148">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="7ac52-148">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="7ac52-149">Les valeurs possibles sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="7ac52-149">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="7ac52-150">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac52-150">Value</span></span> | <span data-ttu-id="7ac52-151">Description</span><span class="sxs-lookup"><span data-stu-id="7ac52-151">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="7ac52-152">La boîte aux lettres se trouve sur un serveur Exchange local.</span><span class="sxs-lookup"><span data-stu-id="7ac52-152">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="7ac52-153">La boîte aux lettres est associée à un compte gmail.</span><span class="sxs-lookup"><span data-stu-id="7ac52-153">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="7ac52-154">La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365.</span><span class="sxs-lookup"><span data-stu-id="7ac52-154">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="7ac52-155">La boîte aux lettres est associée à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="7ac52-155">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="7ac52-156">Type</span><span class="sxs-lookup"><span data-stu-id="7ac52-156">Type</span></span>

*   <span data-ttu-id="7ac52-157">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac52-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac52-158">Requirements</span></span>

|<span data-ttu-id="7ac52-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac52-159">Requirement</span></span>| <span data-ttu-id="7ac52-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac52-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac52-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac52-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ac52-162">1.6</span><span class="sxs-lookup"><span data-stu-id="7ac52-162">1.6</span></span> |
|[<span data-ttu-id="7ac52-163">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7ac52-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7ac52-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-164">ReadItem</span></span>|
|[<span data-ttu-id="7ac52-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac52-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ac52-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-166">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac52-167">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac52-167">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="7ac52-168">displayName : String</span><span class="sxs-lookup"><span data-stu-id="7ac52-168">displayName: String</span></span>

<span data-ttu-id="7ac52-169">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7ac52-169">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac52-170">Type</span><span class="sxs-lookup"><span data-stu-id="7ac52-170">Type</span></span>

*   <span data-ttu-id="7ac52-171">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-171">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac52-172">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac52-172">Requirements</span></span>

|<span data-ttu-id="7ac52-173">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac52-173">Requirement</span></span>| <span data-ttu-id="7ac52-174">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac52-174">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac52-175">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac52-175">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ac52-176">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-176">1.0</span></span>|
|[<span data-ttu-id="7ac52-177">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7ac52-177">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7ac52-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-178">ReadItem</span></span>|
|[<span data-ttu-id="7ac52-179">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac52-179">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ac52-180">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-180">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac52-181">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac52-181">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="7ac52-182">emailAddress : chaîne</span><span class="sxs-lookup"><span data-stu-id="7ac52-182">emailAddress: String</span></span>

<span data-ttu-id="7ac52-183">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7ac52-183">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac52-184">Type</span><span class="sxs-lookup"><span data-stu-id="7ac52-184">Type</span></span>

*   <span data-ttu-id="7ac52-185">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac52-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac52-186">Requirements</span></span>

|<span data-ttu-id="7ac52-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac52-187">Requirement</span></span>| <span data-ttu-id="7ac52-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac52-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac52-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac52-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ac52-190">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-190">1.0</span></span>|
|[<span data-ttu-id="7ac52-191">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7ac52-191">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7ac52-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-192">ReadItem</span></span>|
|[<span data-ttu-id="7ac52-193">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac52-193">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ac52-194">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-194">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac52-195">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac52-195">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="7ac52-196">timeZone : chaîne</span><span class="sxs-lookup"><span data-stu-id="7ac52-196">timeZone: String</span></span>

<span data-ttu-id="7ac52-197">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7ac52-197">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac52-198">Type</span><span class="sxs-lookup"><span data-stu-id="7ac52-198">Type</span></span>

*   <span data-ttu-id="7ac52-199">String</span><span class="sxs-lookup"><span data-stu-id="7ac52-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac52-200">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac52-200">Requirements</span></span>

|<span data-ttu-id="7ac52-201">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac52-201">Requirement</span></span>| <span data-ttu-id="7ac52-202">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac52-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac52-203">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac52-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ac52-204">1.0</span><span class="sxs-lookup"><span data-stu-id="7ac52-204">1.0</span></span>|
|[<span data-ttu-id="7ac52-205">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7ac52-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7ac52-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7ac52-206">ReadItem</span></span>|
|[<span data-ttu-id="7ac52-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac52-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ac52-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac52-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac52-209">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac52-209">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
