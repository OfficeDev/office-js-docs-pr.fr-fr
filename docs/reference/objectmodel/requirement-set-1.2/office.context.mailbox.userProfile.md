---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 4a6739c9b463e49d41e320094a4c9cb1a32655f4
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067826"
---
# <a name="userprofile"></a><span data-ttu-id="5d7f1-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="5d7f1-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="5d7f1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="5d7f1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7f1-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d7f1-104">Requirements</span></span>

|<span data-ttu-id="5d7f1-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d7f1-105">Requirement</span></span>| <span data-ttu-id="5d7f1-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d7f1-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7f1-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d7f1-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7f1-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7f1-108">1.0</span></span>|
|[<span data-ttu-id="5d7f1-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="5d7f1-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7f1-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7f1-110">ReadItem</span></span>|
|[<span data-ttu-id="5d7f1-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d7f1-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7f1-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d7f1-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="5d7f1-113">Membres</span><span class="sxs-lookup"><span data-stu-id="5d7f1-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="5d7f1-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5d7f1-114">displayName :String</span></span>

<span data-ttu-id="5d7f1-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5d7f1-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5d7f1-116">Type</span><span class="sxs-lookup"><span data-stu-id="5d7f1-116">Type</span></span>

*   <span data-ttu-id="5d7f1-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d7f1-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7f1-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d7f1-118">Requirements</span></span>

|<span data-ttu-id="5d7f1-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d7f1-119">Requirement</span></span>| <span data-ttu-id="5d7f1-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d7f1-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7f1-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d7f1-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7f1-122">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7f1-122">1.0</span></span>|
|[<span data-ttu-id="5d7f1-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="5d7f1-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7f1-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7f1-124">ReadItem</span></span>|
|[<span data-ttu-id="5d7f1-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d7f1-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7f1-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d7f1-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7f1-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="5d7f1-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5d7f1-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5d7f1-128">emailAddress :String</span></span>

<span data-ttu-id="5d7f1-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5d7f1-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5d7f1-130">Type</span><span class="sxs-lookup"><span data-stu-id="5d7f1-130">Type</span></span>

*   <span data-ttu-id="5d7f1-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d7f1-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7f1-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d7f1-132">Requirements</span></span>

|<span data-ttu-id="5d7f1-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d7f1-133">Requirement</span></span>| <span data-ttu-id="5d7f1-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d7f1-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7f1-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d7f1-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7f1-136">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7f1-136">1.0</span></span>|
|[<span data-ttu-id="5d7f1-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="5d7f1-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7f1-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7f1-138">ReadItem</span></span>|
|[<span data-ttu-id="5d7f1-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d7f1-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7f1-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d7f1-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7f1-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="5d7f1-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5d7f1-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5d7f1-142">timeZone :String</span></span>

<span data-ttu-id="5d7f1-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5d7f1-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5d7f1-144">Type</span><span class="sxs-lookup"><span data-stu-id="5d7f1-144">Type</span></span>

*   <span data-ttu-id="5d7f1-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d7f1-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7f1-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d7f1-146">Requirements</span></span>

|<span data-ttu-id="5d7f1-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d7f1-147">Requirement</span></span>| <span data-ttu-id="5d7f1-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d7f1-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7f1-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d7f1-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7f1-150">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7f1-150">1.0</span></span>|
|[<span data-ttu-id="5d7f1-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="5d7f1-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7f1-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7f1-152">ReadItem</span></span>|
|[<span data-ttu-id="5d7f1-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d7f1-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7f1-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d7f1-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7f1-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="5d7f1-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
