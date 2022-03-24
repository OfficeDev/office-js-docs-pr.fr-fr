---
title: Obtenir et définir des catégories
description: Comment gérer les catégories sur la boîte aux lettres et l’élément.
ms.date: 01/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: 93f9167fcc31110543d08019e5428952beab0ccc
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746300"
---
# <a name="get-and-set-categories"></a>Obtenir et définir des catégories

Dans Outlook, un utilisateur peut appliquer des catégories aux messages et aux rendez-vous pour organiser ses données de boîte aux lettres. L’utilisateur définit la liste principale des catégories codées en couleur pour sa boîte aux lettres, puis peut appliquer une ou plusieurs de ces catégories à n’importe quel élément de message ou de rendez-vous. Chaque [catégorie](/javascript/api/outlook/office.categorydetails) de la liste principale est représentée par le nom et la [couleur](/javascript/api/outlook/office.mailboxenums.categorycolor) spécifiés par l’utilisateur. Vous pouvez utiliser l’API JavaScript Office pour gérer la liste principale des catégories sur la boîte aux lettres et les catégories appliquées à un élément.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.8. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="manage-categories-in-the-master-list"></a>Gérer les catégories dans la liste principale

Seules les catégories de la liste principale de votre boîte aux lettres peuvent être appliquées à un message ou à un rendez-vous. Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories maîtres.

> [!IMPORTANT]
> Pour que le add-in gère la liste principale des catégories, `Permissions` vous devez définir le nœud dans le manifeste sur `ReadWriteMailbox`.

### <a name="add-master-categories"></a>Ajouter des catégories principales

L’exemple suivant montre comment ajouter une catégorie nommée « Urgent! » à la liste principale en appelant [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) sur [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

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

### <a name="get-master-categories"></a>Obtenir les catégories principales

L’exemple suivant montre comment obtenir la liste des catégories en appelant [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) sur [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

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

### <a name="remove-master-categories"></a>Supprimer des catégories principales

L’exemple suivant montre comment supprimer la catégorie nommée « Urgent! » à partir de la liste principale en appelant [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) sur [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

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

Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories pour un élément de message ou de rendez-vous.

> [!IMPORTANT]
> Seules les catégories de la liste principale de votre boîte aux lettres peuvent être appliquées à un message ou à un rendez-vous. Pour plus d’informations, voir la section Précédente Gérer les [catégories dans la liste principale](#manage-categories-in-the-master-list) .
>
> Dans Outlook sur le web, vous ne pouvez pas utiliser l’API pour gérer les catégories d’un message en mode lecture.

### <a name="add-categories-to-an-item"></a>Ajouter des catégories à un élément

L’exemple suivant montre comment appliquer la catégorie nommée « Urgent! » à l’élément actuel en appelant [addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) on `item.categories`.

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

L’exemple suivant montre comment obtenir les catégories appliquées à l’élément actuel en appelant [getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) sur `item.categories`.

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

L’exemple suivant montre comment supprimer la catégorie nommée « Urgent! » à partir de l’élément actuel en [appelant removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) on `item.categories`.

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

- [Outlook d’autorisations](understanding-outlook-add-in-permissions.md)
- [Élément Permissions dans le manifeste](../reference/manifest/permissions.md)
