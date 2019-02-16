---
title: Office.context.mailbox.userProfile-ensemble de conditions requises 1.1
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 265102f42161ffcb326dbeffb7936af78876a41c
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068279"
---
# <a name="userprofile"></a><span data-ttu-id="be2a8-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="be2a8-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="be2a8-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="be2a8-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="be2a8-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be2a8-104">Requirements</span></span>

|<span data-ttu-id="be2a8-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be2a8-105">Requirement</span></span>| <span data-ttu-id="be2a8-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="be2a8-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="be2a8-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be2a8-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be2a8-108">1.0</span><span class="sxs-lookup"><span data-stu-id="be2a8-108">1.0</span></span>|
|[<span data-ttu-id="be2a8-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be2a8-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be2a8-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be2a8-110">ReadItem</span></span>|
|[<span data-ttu-id="be2a8-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be2a8-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be2a8-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be2a8-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="be2a8-113">Membres</span><span class="sxs-lookup"><span data-stu-id="be2a8-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="be2a8-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="be2a8-114">displayName :String</span></span>

<span data-ttu-id="be2a8-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be2a8-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="be2a8-116">Type</span><span class="sxs-lookup"><span data-stu-id="be2a8-116">Type</span></span>

*   <span data-ttu-id="be2a8-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="be2a8-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be2a8-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be2a8-118">Requirements</span></span>

|<span data-ttu-id="be2a8-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be2a8-119">Requirement</span></span>| <span data-ttu-id="be2a8-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="be2a8-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="be2a8-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be2a8-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be2a8-122">1.0</span><span class="sxs-lookup"><span data-stu-id="be2a8-122">1.0</span></span>|
|[<span data-ttu-id="be2a8-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be2a8-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be2a8-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be2a8-124">ReadItem</span></span>|
|[<span data-ttu-id="be2a8-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be2a8-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be2a8-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be2a8-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be2a8-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="be2a8-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="be2a8-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="be2a8-128">emailAddress :String</span></span>

<span data-ttu-id="be2a8-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be2a8-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="be2a8-130">Type</span><span class="sxs-lookup"><span data-stu-id="be2a8-130">Type</span></span>

*   <span data-ttu-id="be2a8-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="be2a8-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be2a8-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be2a8-132">Requirements</span></span>

|<span data-ttu-id="be2a8-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be2a8-133">Requirement</span></span>| <span data-ttu-id="be2a8-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="be2a8-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="be2a8-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be2a8-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be2a8-136">1.0</span><span class="sxs-lookup"><span data-stu-id="be2a8-136">1.0</span></span>|
|[<span data-ttu-id="be2a8-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be2a8-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be2a8-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be2a8-138">ReadItem</span></span>|
|[<span data-ttu-id="be2a8-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be2a8-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be2a8-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be2a8-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be2a8-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="be2a8-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="be2a8-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="be2a8-142">timeZone :String</span></span>

<span data-ttu-id="be2a8-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be2a8-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="be2a8-144">Type</span><span class="sxs-lookup"><span data-stu-id="be2a8-144">Type</span></span>

*   <span data-ttu-id="be2a8-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="be2a8-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be2a8-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be2a8-146">Requirements</span></span>

|<span data-ttu-id="be2a8-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be2a8-147">Requirement</span></span>| <span data-ttu-id="be2a8-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="be2a8-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="be2a8-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be2a8-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be2a8-150">1.0</span><span class="sxs-lookup"><span data-stu-id="be2a8-150">1.0</span></span>|
|[<span data-ttu-id="be2a8-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be2a8-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be2a8-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be2a8-152">ReadItem</span></span>|
|[<span data-ttu-id="be2a8-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be2a8-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be2a8-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be2a8-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be2a8-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="be2a8-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
