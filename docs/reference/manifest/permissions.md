---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 9193651ec0c795cdb55eb3fc6576dbacd59e0fb2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432353"
---
# <a name="permissions-element"></a><span data-ttu-id="b135e-102">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="b135e-102">Permissions element</span></span>

<span data-ttu-id="b135e-103">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="b135e-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="b135e-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="b135e-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b135e-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b135e-105">Syntax</span></span>

<span data-ttu-id="b135e-106">Pour les compléments du volet de tâches et de contenu:</span><span class="sxs-lookup"><span data-stu-id="b135e-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="b135e-107">Pour les compléments de messagerie:</span><span class="sxs-lookup"><span data-stu-id="b135e-107">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="b135e-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b135e-108">Contained in</span></span>

[<span data-ttu-id="b135e-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b135e-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b135e-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="b135e-110">Remarks</span></span>

<span data-ttu-id="b135e-111">Pour plus de détails, consultez l’article relatif à la [demande d’autorisations pour utiliser des API dans des compléments de contenu et de volet Office](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et celui décrivant les [autorisations de complément Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="b135e-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
