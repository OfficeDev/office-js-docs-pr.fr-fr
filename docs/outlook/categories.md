---
title: Obtenir et définir des catégories
description: Comment gérer les catégories sur la boîte aux lettres et l’élément
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d4589571de47218741308c01caec0166d72919d8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608977"
---
# <a name="get-and-set-categories"></a><span data-ttu-id="1a2b2-103">Obtenir et définir des catégories</span><span class="sxs-lookup"><span data-stu-id="1a2b2-103">Get and set categories</span></span>

<span data-ttu-id="1a2b2-104">Dans Outlook, un utilisateur peut appliquer des catégories à des messages et à des rendez-vous afin d’organiser les données de leurs boîtes aux lettres.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-104">In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data.</span></span> <span data-ttu-id="1a2b2-105">L’utilisateur définit la liste principale des catégories codées en couleur pour sa boîte aux lettres, puis il peut appliquer une ou plusieurs de ces catégories à un message ou à un élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-105">The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item.</span></span> <span data-ttu-id="1a2b2-106">Chaque [catégorie](/javascript/api/outlook/office.categorydetails) de la liste principale est représentée par le nom et la [couleur](/javascript/api/outlook/office.mailboxenums.categorycolor) spécifiés par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-106">Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies.</span></span> <span data-ttu-id="1a2b2-107">Vous pouvez utiliser l’API JavaScript pour Office pour gérer la liste principale des catégories dans la boîte aux lettres et les catégories appliquées à un élément.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-107">You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="1a2b2-108">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="1a2b2-109">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="manage-categories-in-the-master-list"></a><span data-ttu-id="1a2b2-110">Gérer les catégories dans la liste principale</span><span class="sxs-lookup"><span data-stu-id="1a2b2-110">Manage categories in the master list</span></span>

<span data-ttu-id="1a2b2-111">Seules les catégories dans la liste principale de votre boîte aux lettres peuvent être appliquées à un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-111">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="1a2b2-112">Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories principales.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-112">You can use the API to add, get, and remove master categories.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1a2b2-113">Pour que le complément gère la liste principale des catégories, vous devez définir le `Permissions` nœud dans le manifeste sur `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="1a2b2-113">For the add-in to manage the categories master list, you must set the `Permissions` node in the manifest to `ReadWriteMailbox`.</span></span>

### <a name="add-master-categories"></a><span data-ttu-id="1a2b2-114">Ajouter des catégories principales</span><span class="sxs-lookup"><span data-stu-id="1a2b2-114">Add master categories</span></span>

<span data-ttu-id="1a2b2-115">L’exemple suivant montre comment ajouter une catégorie nommée « urgent ! »</span><span class="sxs-lookup"><span data-stu-id="1a2b2-115">The following example shows how to add a category named "Urgent!"</span></span> <span data-ttu-id="1a2b2-116">à la liste principale en appelant [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) sur [Mailbox. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="1a2b2-116">to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-master-categories"></a><span data-ttu-id="1a2b2-117">Obtenir des catégories de formes de base</span><span class="sxs-lookup"><span data-stu-id="1a2b2-117">Get master categories</span></span>

<span data-ttu-id="1a2b2-118">L’exemple suivant montre comment obtenir la liste des catégories en appelant [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) sur [Mailbox. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="1a2b2-118">The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a><span data-ttu-id="1a2b2-119">Supprimer des catégories de formes de base</span><span class="sxs-lookup"><span data-stu-id="1a2b2-119">Remove master categories</span></span>

<span data-ttu-id="1a2b2-120">L’exemple suivant montre comment supprimer la catégorie nommée « urgent ! »</span><span class="sxs-lookup"><span data-stu-id="1a2b2-120">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="1a2b2-121">à partir de la liste principale en appelant [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) sur [Mailbox. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="1a2b2-121">from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a><span data-ttu-id="1a2b2-122">Gérer les catégories d’un message ou d’un rendez-vous</span><span class="sxs-lookup"><span data-stu-id="1a2b2-122">Manage categories on a message or appointment</span></span>

<span data-ttu-id="1a2b2-123">Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories pour un message ou un élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-123">You can use the API to add, get, and remove categories for a message or appointment item.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1a2b2-124">Seules les catégories dans la liste principale de votre boîte aux lettres peuvent être appliquées à un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-124">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="1a2b2-125">Pour plus d’informations, reportez-vous à la section précédente [gérer les catégories dans la liste principale](#manage-categories-in-the-master-list) .</span><span class="sxs-lookup"><span data-stu-id="1a2b2-125">See the earlier section [Manage categories in the master list](#manage-categories-in-the-master-list) for more information.</span></span>
>
> <span data-ttu-id="1a2b2-126">Dans Outlook sur le Web, vous ne pouvez pas utiliser l’API pour gérer les catégories d’un message en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="1a2b2-126">In Outlook on the web, you can't use the API to manage categories on a message in Read mode.</span></span>

### <a name="add-categories-to-an-item"></a><span data-ttu-id="1a2b2-127">Ajouter des catégories à un élément</span><span class="sxs-lookup"><span data-stu-id="1a2b2-127">Add categories to an item</span></span>

<span data-ttu-id="1a2b2-128">L’exemple suivant montre comment appliquer la catégorie nommée « urgent ! »</span><span class="sxs-lookup"><span data-stu-id="1a2b2-128">The following example shows how to apply the category named "Urgent!"</span></span> <span data-ttu-id="1a2b2-129">à l’élément actuel en appelant [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) `item.categories` .</span><span class="sxs-lookup"><span data-stu-id="1a2b2-129">to the current item by calling [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a><span data-ttu-id="1a2b2-130">Obtenir les catégories d’un élément</span><span class="sxs-lookup"><span data-stu-id="1a2b2-130">Get an item's categories</span></span>

<span data-ttu-id="1a2b2-131">L’exemple suivant montre comment obtenir les catégories appliquées à l’élément actuel en appelant [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) `item.categories` .</span><span class="sxs-lookup"><span data-stu-id="1a2b2-131">The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a><span data-ttu-id="1a2b2-132">Supprimer des catégories d’un élément</span><span class="sxs-lookup"><span data-stu-id="1a2b2-132">Remove categories from an item</span></span>

<span data-ttu-id="1a2b2-133">L’exemple suivant montre comment supprimer la catégorie nommée « urgent ! »</span><span class="sxs-lookup"><span data-stu-id="1a2b2-133">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="1a2b2-134">à partir de l’élément actuel en appelant [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) `item.categories` .</span><span class="sxs-lookup"><span data-stu-id="1a2b2-134">from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a><span data-ttu-id="1a2b2-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1a2b2-135">See also</span></span>

- [<span data-ttu-id="1a2b2-136">Autorisations Outlook</span><span class="sxs-lookup"><span data-stu-id="1a2b2-136">Outlook permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="1a2b2-137">Élément permissions dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="1a2b2-137">Permissions element in the manifest</span></span>](../reference/manifest/permissions.md)
