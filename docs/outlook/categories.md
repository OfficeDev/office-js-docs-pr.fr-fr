---
title: Obtenir et définir des catégories
description: Guide pratique pour gérer les catégories sur la boîte aux lettres et l’élément.
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: d31cb8da4cdaf4a88141a1eac927748b1399e0d9
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712824"
---
# <a name="get-and-set-categories"></a>Obtenir et définir des catégories

Dans Outlook, un utilisateur peut appliquer des catégories aux messages et aux rendez-vous pour organiser ses données de boîte aux lettres. L’utilisateur définit la liste principale des catégories codées en couleurs pour sa boîte aux lettres, puis peut appliquer une ou plusieurs de ces catégories à n’importe quel élément de message ou de rendez-vous. Chaque [catégorie](/javascript/api/outlook/office.categorydetails) de la liste maître est représentée par le nom et la [couleur](/javascript/api/outlook/office.mailboxenums.categorycolor) spécifiés par l’utilisateur. Vous pouvez utiliser l’API JavaScript Office pour gérer la liste principale des catégories sur la boîte aux lettres et les catégories appliquées à un élément.

> [!NOTE]
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.8. Voir [les clients et les plateformes](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="manage-categories-in-the-master-list"></a>Gérer les catégories dans la liste maître

Seules les catégories de la liste maître de votre boîte aux lettres peuvent s’appliquer à un message ou à un rendez-vous. Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories principales.

> [!IMPORTANT]
> Pour que le complément gère la liste maître des catégories, vous devez définir le `Permissions` nœud dans le manifeste `ReadWriteMailbox`sur .

### <a name="add-master-categories"></a>Ajouter des catégories de maîtres

L’exemple suivant montre comment ajouter une catégorie nommée « Urgent! » à la liste maître en appelant [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) sur [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
const masterCategoriesToAdd = [
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
        const masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a>Supprimer les catégories principales

L’exemple suivant montre comment supprimer la catégorie nommée « Urgent! » à partir de la liste maître en appelant [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) sur [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
const masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a>Gérer les catégories sur un message ou un rendez-vous

Vous pouvez utiliser l’API pour ajouter, obtenir et supprimer des catégories pour un message ou un élément de rendez-vous.

> [!IMPORTANT]
> Seules les catégories de la liste maître de votre boîte aux lettres peuvent s’appliquer à un message ou à un rendez-vous. Pour plus d’informations, consultez la section précédente [Gérer les catégories dans la liste maître](#manage-categories-in-the-master-list) .
>
> Dans Outlook sur le web, vous ne pouvez pas utiliser l’API pour gérer les catégories d’un message en mode lecture.

### <a name="add-categories-to-an-item"></a>Ajouter des catégories à un élément

L’exemple suivant montre comment appliquer la catégorie nommée « Urgent! » à l’élément actuel en appelant [addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) le `item.categories`.

```js
const categoriesToAdd = ["Urgent!"];

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
        const categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a>Supprimer des catégories d’un élément

L’exemple suivant montre comment supprimer la catégorie nommée « Urgent! » à partir de l’élément actuel en appelant [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) le `item.categories`.

```js
const categoriesToRemove = ["Urgent!"];

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
- [Élément Permissions dans le manifeste](/javascript/api/manifest/permissions)
