---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.2
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: e5548fa514cff9b452c2747324f11e5df8a06def
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432241"
---
# <a name="userprofile"></a><span data-ttu-id="eba3a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="eba3a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="eba3a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="eba3a-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="eba3a-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eba3a-104">Requirements</span></span>

|<span data-ttu-id="eba3a-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eba3a-105">Requirement</span></span>| <span data-ttu-id="eba3a-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="eba3a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba3a-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eba3a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba3a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="eba3a-108">1.0</span></span>|
|[<span data-ttu-id="eba3a-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="eba3a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eba3a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eba3a-110">ReadItem</span></span>|
|[<span data-ttu-id="eba3a-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eba3a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba3a-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eba3a-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="eba3a-113">Membres</span><span class="sxs-lookup"><span data-stu-id="eba3a-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="eba3a-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="eba3a-114">displayName :String</span></span>

<span data-ttu-id="eba3a-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="eba3a-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="eba3a-116">Type :</span><span class="sxs-lookup"><span data-stu-id="eba3a-116">Type:</span></span>

*   <span data-ttu-id="eba3a-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eba3a-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eba3a-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eba3a-118">Requirements</span></span>

|<span data-ttu-id="eba3a-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eba3a-119">Requirement</span></span>| <span data-ttu-id="eba3a-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="eba3a-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba3a-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eba3a-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba3a-122">1.0</span><span class="sxs-lookup"><span data-stu-id="eba3a-122">1.0</span></span>|
|[<span data-ttu-id="eba3a-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="eba3a-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eba3a-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eba3a-124">ReadItem</span></span>|
|[<span data-ttu-id="eba3a-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eba3a-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba3a-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eba3a-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="eba3a-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="eba3a-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="eba3a-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="eba3a-128">emailAddress :String</span></span>

<span data-ttu-id="eba3a-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="eba3a-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="eba3a-130">Type :</span><span class="sxs-lookup"><span data-stu-id="eba3a-130">Type:</span></span>

*   <span data-ttu-id="eba3a-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eba3a-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eba3a-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eba3a-132">Requirements</span></span>

|<span data-ttu-id="eba3a-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eba3a-133">Requirement</span></span>| <span data-ttu-id="eba3a-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="eba3a-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba3a-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eba3a-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba3a-136">1.0</span><span class="sxs-lookup"><span data-stu-id="eba3a-136">1.0</span></span>|
|[<span data-ttu-id="eba3a-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="eba3a-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eba3a-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eba3a-138">ReadItem</span></span>|
|[<span data-ttu-id="eba3a-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eba3a-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba3a-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eba3a-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="eba3a-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="eba3a-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="eba3a-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="eba3a-142">timeZone :String</span></span>

<span data-ttu-id="eba3a-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="eba3a-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="eba3a-144">Type :</span><span class="sxs-lookup"><span data-stu-id="eba3a-144">Type:</span></span>

*   <span data-ttu-id="eba3a-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eba3a-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eba3a-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eba3a-146">Requirements</span></span>

|<span data-ttu-id="eba3a-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eba3a-147">Requirement</span></span>| <span data-ttu-id="eba3a-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="eba3a-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="eba3a-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eba3a-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eba3a-150">1.0</span><span class="sxs-lookup"><span data-stu-id="eba3a-150">1.0</span></span>|
|[<span data-ttu-id="eba3a-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="eba3a-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eba3a-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eba3a-152">ReadItem</span></span>|
|[<span data-ttu-id="eba3a-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eba3a-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eba3a-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eba3a-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="eba3a-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="eba3a-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```