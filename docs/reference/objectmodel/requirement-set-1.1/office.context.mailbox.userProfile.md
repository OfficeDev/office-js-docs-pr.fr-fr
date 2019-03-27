---
title: Office. Context. Mailbox. userProfile-ensemble de conditions requises 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870190"
---
# <a name="userprofile"></a><span data-ttu-id="a54b5-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a54b5-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a54b5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a54b5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a54b5-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a54b5-104">Requirements</span></span>

|<span data-ttu-id="a54b5-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a54b5-105">Requirement</span></span>| <span data-ttu-id="a54b5-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="a54b5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a54b5-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a54b5-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a54b5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a54b5-108">1.0</span></span>|
|[<span data-ttu-id="a54b5-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a54b5-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a54b5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a54b5-110">ReadItem</span></span>|
|[<span data-ttu-id="a54b5-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a54b5-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a54b5-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a54b5-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="a54b5-113">Membres</span><span class="sxs-lookup"><span data-stu-id="a54b5-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a54b5-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a54b5-114">displayName :String</span></span>

<span data-ttu-id="a54b5-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a54b5-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a54b5-116">Type</span><span class="sxs-lookup"><span data-stu-id="a54b5-116">Type</span></span>

*   <span data-ttu-id="a54b5-117">String</span><span class="sxs-lookup"><span data-stu-id="a54b5-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a54b5-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a54b5-118">Requirements</span></span>

|<span data-ttu-id="a54b5-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a54b5-119">Requirement</span></span>| <span data-ttu-id="a54b5-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="a54b5-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="a54b5-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a54b5-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a54b5-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a54b5-122">1.0</span></span>|
|[<span data-ttu-id="a54b5-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a54b5-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a54b5-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a54b5-124">ReadItem</span></span>|
|[<span data-ttu-id="a54b5-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a54b5-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a54b5-126">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a54b5-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a54b5-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="a54b5-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a54b5-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a54b5-128">emailAddress :String</span></span>

<span data-ttu-id="a54b5-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a54b5-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a54b5-130">Type</span><span class="sxs-lookup"><span data-stu-id="a54b5-130">Type</span></span>

*   <span data-ttu-id="a54b5-131">String</span><span class="sxs-lookup"><span data-stu-id="a54b5-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a54b5-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a54b5-132">Requirements</span></span>

|<span data-ttu-id="a54b5-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a54b5-133">Requirement</span></span>| <span data-ttu-id="a54b5-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="a54b5-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a54b5-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a54b5-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a54b5-136">1.0</span><span class="sxs-lookup"><span data-stu-id="a54b5-136">1.0</span></span>|
|[<span data-ttu-id="a54b5-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a54b5-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a54b5-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a54b5-138">ReadItem</span></span>|
|[<span data-ttu-id="a54b5-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a54b5-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a54b5-140">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a54b5-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a54b5-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="a54b5-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a54b5-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a54b5-142">timeZone :String</span></span>

<span data-ttu-id="a54b5-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a54b5-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a54b5-144">Type</span><span class="sxs-lookup"><span data-stu-id="a54b5-144">Type</span></span>

*   <span data-ttu-id="a54b5-145">String</span><span class="sxs-lookup"><span data-stu-id="a54b5-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a54b5-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a54b5-146">Requirements</span></span>

|<span data-ttu-id="a54b5-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a54b5-147">Requirement</span></span>| <span data-ttu-id="a54b5-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="a54b5-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a54b5-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a54b5-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a54b5-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a54b5-150">1.0</span></span>|
|[<span data-ttu-id="a54b5-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a54b5-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a54b5-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a54b5-152">ReadItem</span></span>|
|[<span data-ttu-id="a54b5-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a54b5-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a54b5-154">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a54b5-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a54b5-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="a54b5-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
