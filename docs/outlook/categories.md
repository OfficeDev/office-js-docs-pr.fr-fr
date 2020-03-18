---
title: Obtenir et définir des catégories
description: Comment gérer les catégories sur la boîte aux lettres et l’élément
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d0bb2e9f51675c263d0a3a130c64e02e7d55b764
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721022"
---
# <a name="get-and-set-categories"></a>Obtenir et définir des catégories

Dans Outlook, un utilisateur peut appliquer des catégories à des messages et à des rendez-vous afin d’organiser les données de leurs boîtes aux lettres. L’utilisateur définit la liste principale des catégories codées en couleur pour sa boîte aux lettres, puis il peut appliquer une ou plusieurs de ces catégories à un message ou à un élément de rendez-vous. Chaque [catégorie](/javascript/api/outlook/office.categorydetails) de la liste principale est représentée par le nom et la [couleur](/javascript/api/outlook/office.mailboxenums.categorycolor) spécifiés par l’utilisateur. Vous pouvez utiliser l’API JavaScript pour Office pour gérer la liste principale des catégories dans la boîte aux lettres et les catégories appliquées à un élément.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,8. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="manage-categories-in-the-master-list"></a>Gérer les catégories dans la liste principale

Seules les catégories dans la liste principale de votre boîte aux lettres peuvent être appliquées à un message ou un rendez-vous. Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories principales.

> [!IMPORTANT]
> Pour que le complément gère la liste principale des catégories, vous devez définir le `Permissions` nœud dans le manifeste sur. `ReadWriteMailbox`

### <a name="add-master-categories"></a>Ajouter des catégories principales

L’exemple suivant montre comment ajouter une catégorie nommée « urgent ! » à la liste principale en appelant [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) sur [Mailbox. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).

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

### <a name="get-master-categories"></a>Obtenir des catégories de formes de base

L’exemple suivant montre comment obtenir la liste des catégories en appelant [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) sur [Mailbox. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).

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

### <a name="remove-master-categories"></a>Supprimer des catégories de formes de base

L’exemple suivant montre comment supprimer la catégorie nommée « urgent ! » à partir de la liste principale en appelant [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) sur [Mailbox. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).

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

## <a name="manage-categories-on-a-message-or-appointment"></a>Gérer les catégories d’un message ou d’un rendez-vous

Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories pour un message ou un élément de rendez-vous.

> [!IMPORTANT]
> Seules les catégories dans la liste principale de votre boîte aux lettres peuvent être appliquées à un message ou un rendez-vous. Pour plus d’informations, reportez-vous à la section précédente [gérer les catégories dans la liste principale](#manage-categories-in-the-master-list) .
>
> Dans Outlook sur le Web, vous ne pouvez pas utiliser l’API pour gérer les catégories d’un message en mode lecture.

### <a name="add-categories-to-an-item"></a>Ajouter des catégories à un élément

L’exemple suivant montre comment appliquer la catégorie nommée « urgent ! » à l’élément actuel en appelant [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) `item.categories`.

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

### <a name="get-an-items-categories"></a>Obtenir les catégories d’un élément

L’exemple suivant montre comment obtenir les catégories appliquées à l’élément actuel en appelant [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) `item.categories`.

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

### <a name="remove-categories-from-an-item"></a>Supprimer des catégories d’un élément

L’exemple suivant montre comment supprimer la catégorie nommée « urgent ! » à partir de l’élément actuel [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) en appelant `item.categories`removeAsync.

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

## <a name="see-also"></a>Voir aussi

- [Autorisations Outlook](understanding-outlook-add-in-permissions.md)
- [Élément permissions dans le manifeste](../reference/manifest/permissions.md)
