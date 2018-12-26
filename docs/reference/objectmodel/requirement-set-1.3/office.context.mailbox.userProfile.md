---
title: Office.context.mailbox.userProfile-ensemble de conditions requises 1.3
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 9f36b5f1d31ad6709cf2c43ce7dcb3f91a35bd00
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432220"
---
# <a name="userprofile"></a><span data-ttu-id="8c08f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8c08f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8c08f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8c08f-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c08f-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8c08f-104">Requirements</span></span>

|<span data-ttu-id="8c08f-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8c08f-105">Requirement</span></span>| <span data-ttu-id="8c08f-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="8c08f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c08f-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8c08f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c08f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8c08f-108">1.0</span></span>|
|[<span data-ttu-id="8c08f-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8c08f-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c08f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c08f-110">ReadItem</span></span>|
|[<span data-ttu-id="8c08f-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8c08f-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c08f-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8c08f-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="8c08f-113">Membres</span><span class="sxs-lookup"><span data-stu-id="8c08f-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="8c08f-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="8c08f-114">displayName :String</span></span>

<span data-ttu-id="8c08f-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8c08f-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8c08f-116">Type :</span><span class="sxs-lookup"><span data-stu-id="8c08f-116">Type:</span></span>

*   <span data-ttu-id="8c08f-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8c08f-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c08f-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8c08f-118">Requirements</span></span>

|<span data-ttu-id="8c08f-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8c08f-119">Requirement</span></span>| <span data-ttu-id="8c08f-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="8c08f-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c08f-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8c08f-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c08f-122">1.0</span><span class="sxs-lookup"><span data-stu-id="8c08f-122">1.0</span></span>|
|[<span data-ttu-id="8c08f-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8c08f-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c08f-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c08f-124">ReadItem</span></span>|
|[<span data-ttu-id="8c08f-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8c08f-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c08f-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8c08f-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c08f-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="8c08f-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="8c08f-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="8c08f-128">emailAddress :String</span></span>

<span data-ttu-id="8c08f-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8c08f-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8c08f-130">Type :</span><span class="sxs-lookup"><span data-stu-id="8c08f-130">Type:</span></span>

*   <span data-ttu-id="8c08f-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8c08f-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c08f-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8c08f-132">Requirements</span></span>

|<span data-ttu-id="8c08f-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8c08f-133">Requirement</span></span>| <span data-ttu-id="8c08f-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="8c08f-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c08f-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8c08f-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c08f-136">1.0</span><span class="sxs-lookup"><span data-stu-id="8c08f-136">1.0</span></span>|
|[<span data-ttu-id="8c08f-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8c08f-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c08f-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c08f-138">ReadItem</span></span>|
|[<span data-ttu-id="8c08f-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8c08f-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c08f-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8c08f-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c08f-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="8c08f-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="8c08f-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="8c08f-142">timeZone :String</span></span>

<span data-ttu-id="8c08f-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8c08f-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8c08f-144">Type :</span><span class="sxs-lookup"><span data-stu-id="8c08f-144">Type:</span></span>

*   <span data-ttu-id="8c08f-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8c08f-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c08f-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8c08f-146">Requirements</span></span>

|<span data-ttu-id="8c08f-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8c08f-147">Requirement</span></span>| <span data-ttu-id="8c08f-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="8c08f-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c08f-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8c08f-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c08f-150">1.0</span><span class="sxs-lookup"><span data-stu-id="8c08f-150">1.0</span></span>|
|[<span data-ttu-id="8c08f-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8c08f-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c08f-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c08f-152">ReadItem</span></span>|
|[<span data-ttu-id="8c08f-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8c08f-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c08f-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8c08f-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c08f-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="8c08f-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```