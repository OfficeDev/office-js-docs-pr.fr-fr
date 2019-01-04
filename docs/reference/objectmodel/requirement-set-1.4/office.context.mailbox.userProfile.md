---
title: Office.context.mailbox.userProfile – ensemble de conditions requises 1.4
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 55d0a789c8e46fd3f6ee69f39cf33f7e7d94c322
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432647"
---
# <a name="userprofile"></a><span data-ttu-id="ca66b-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ca66b-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ca66b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ca66b-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca66b-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ca66b-104">Requirements</span></span>

|<span data-ttu-id="ca66b-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ca66b-105">Requirement</span></span>| <span data-ttu-id="ca66b-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="ca66b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca66b-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ca66b-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca66b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ca66b-108">1.0</span></span>|
|[<span data-ttu-id="ca66b-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ca66b-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca66b-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca66b-110">ReadItem</span></span>|
|[<span data-ttu-id="ca66b-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ca66b-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca66b-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="ca66b-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="ca66b-113">Membres</span><span class="sxs-lookup"><span data-stu-id="ca66b-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ca66b-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ca66b-114">displayName :String</span></span>

<span data-ttu-id="ca66b-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ca66b-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ca66b-116">Type :</span><span class="sxs-lookup"><span data-stu-id="ca66b-116">Type:</span></span>

*   <span data-ttu-id="ca66b-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ca66b-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca66b-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ca66b-118">Requirements</span></span>

|<span data-ttu-id="ca66b-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ca66b-119">Requirement</span></span>| <span data-ttu-id="ca66b-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="ca66b-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca66b-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ca66b-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca66b-122">1.0</span><span class="sxs-lookup"><span data-stu-id="ca66b-122">1.0</span></span>|
|[<span data-ttu-id="ca66b-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ca66b-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca66b-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca66b-124">ReadItem</span></span>|
|[<span data-ttu-id="ca66b-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ca66b-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca66b-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="ca66b-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca66b-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="ca66b-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ca66b-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ca66b-128">emailAddress :String</span></span>

<span data-ttu-id="ca66b-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ca66b-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ca66b-130">Type :</span><span class="sxs-lookup"><span data-stu-id="ca66b-130">Type:</span></span>

*   <span data-ttu-id="ca66b-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ca66b-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca66b-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ca66b-132">Requirements</span></span>

|<span data-ttu-id="ca66b-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ca66b-133">Requirement</span></span>| <span data-ttu-id="ca66b-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="ca66b-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca66b-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ca66b-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca66b-136">1.0</span><span class="sxs-lookup"><span data-stu-id="ca66b-136">1.0</span></span>|
|[<span data-ttu-id="ca66b-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ca66b-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca66b-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca66b-138">ReadItem</span></span>|
|[<span data-ttu-id="ca66b-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ca66b-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca66b-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="ca66b-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca66b-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="ca66b-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ca66b-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ca66b-142">timeZone :String</span></span>

<span data-ttu-id="ca66b-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ca66b-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ca66b-144">Type :</span><span class="sxs-lookup"><span data-stu-id="ca66b-144">Type:</span></span>

*   <span data-ttu-id="ca66b-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ca66b-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca66b-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ca66b-146">Requirements</span></span>

|<span data-ttu-id="ca66b-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ca66b-147">Requirement</span></span>| <span data-ttu-id="ca66b-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="ca66b-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca66b-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ca66b-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca66b-150">1.0</span><span class="sxs-lookup"><span data-stu-id="ca66b-150">1.0</span></span>|
|[<span data-ttu-id="ca66b-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ca66b-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca66b-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca66b-152">ReadItem</span></span>|
|[<span data-ttu-id="ca66b-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ca66b-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ca66b-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="ca66b-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca66b-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="ca66b-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```