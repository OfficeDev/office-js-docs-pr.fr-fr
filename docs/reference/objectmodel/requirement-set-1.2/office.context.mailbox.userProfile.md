---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 496a59f4ef02f03cda95fde0bf14634b1db13f77
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870582"
---
# <a name="userprofile"></a><span data-ttu-id="ce97e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ce97e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ce97e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ce97e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce97e-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ce97e-104">Requirements</span></span>

|<span data-ttu-id="ce97e-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ce97e-105">Requirement</span></span>| <span data-ttu-id="ce97e-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="ce97e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce97e-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ce97e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce97e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ce97e-108">1.0</span></span>|
|[<span data-ttu-id="ce97e-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ce97e-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce97e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce97e-110">ReadItem</span></span>|
|[<span data-ttu-id="ce97e-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ce97e-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce97e-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ce97e-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="ce97e-113">Membres</span><span class="sxs-lookup"><span data-stu-id="ce97e-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ce97e-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ce97e-114">displayName :String</span></span>

<span data-ttu-id="ce97e-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce97e-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ce97e-116">Type</span><span class="sxs-lookup"><span data-stu-id="ce97e-116">Type</span></span>

*   <span data-ttu-id="ce97e-117">String</span><span class="sxs-lookup"><span data-stu-id="ce97e-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce97e-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ce97e-118">Requirements</span></span>

|<span data-ttu-id="ce97e-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ce97e-119">Requirement</span></span>| <span data-ttu-id="ce97e-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="ce97e-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce97e-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ce97e-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce97e-122">1.0</span><span class="sxs-lookup"><span data-stu-id="ce97e-122">1.0</span></span>|
|[<span data-ttu-id="ce97e-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ce97e-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce97e-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce97e-124">ReadItem</span></span>|
|[<span data-ttu-id="ce97e-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ce97e-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce97e-126">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ce97e-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce97e-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="ce97e-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ce97e-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ce97e-128">emailAddress :String</span></span>

<span data-ttu-id="ce97e-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce97e-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ce97e-130">Type</span><span class="sxs-lookup"><span data-stu-id="ce97e-130">Type</span></span>

*   <span data-ttu-id="ce97e-131">String</span><span class="sxs-lookup"><span data-stu-id="ce97e-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce97e-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ce97e-132">Requirements</span></span>

|<span data-ttu-id="ce97e-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ce97e-133">Requirement</span></span>| <span data-ttu-id="ce97e-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="ce97e-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce97e-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ce97e-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce97e-136">1.0</span><span class="sxs-lookup"><span data-stu-id="ce97e-136">1.0</span></span>|
|[<span data-ttu-id="ce97e-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ce97e-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce97e-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce97e-138">ReadItem</span></span>|
|[<span data-ttu-id="ce97e-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ce97e-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce97e-140">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ce97e-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce97e-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="ce97e-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ce97e-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ce97e-142">timeZone :String</span></span>

<span data-ttu-id="ce97e-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce97e-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ce97e-144">Type</span><span class="sxs-lookup"><span data-stu-id="ce97e-144">Type</span></span>

*   <span data-ttu-id="ce97e-145">String</span><span class="sxs-lookup"><span data-stu-id="ce97e-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce97e-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ce97e-146">Requirements</span></span>

|<span data-ttu-id="ce97e-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ce97e-147">Requirement</span></span>| <span data-ttu-id="ce97e-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="ce97e-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce97e-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ce97e-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce97e-150">1.0</span><span class="sxs-lookup"><span data-stu-id="ce97e-150">1.0</span></span>|
|[<span data-ttu-id="ce97e-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ce97e-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce97e-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce97e-152">ReadItem</span></span>|
|[<span data-ttu-id="ce97e-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ce97e-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce97e-154">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ce97e-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce97e-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="ce97e-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
