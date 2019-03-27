---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872059"
---
# <a name="permissions-element"></a><span data-ttu-id="c2785-102">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="c2785-102">Permissions element</span></span>

<span data-ttu-id="c2785-103">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="c2785-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="c2785-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="c2785-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c2785-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c2785-105">Syntax</span></span>

<span data-ttu-id="c2785-106">Pour les compléments du volet de tâches et de contenu:</span><span class="sxs-lookup"><span data-stu-id="c2785-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="c2785-107">Pour les compléments de messagerie:</span><span class="sxs-lookup"><span data-stu-id="c2785-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="c2785-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c2785-108">Contained in</span></span>

[<span data-ttu-id="c2785-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c2785-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="c2785-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="c2785-110">Remarks</span></span>

<span data-ttu-id="c2785-111">Pour plus de détails, consultez l’article relatif à la [demande d’autorisations pour utiliser des API dans des compléments de contenu et de volet Office](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et celui décrivant les [autorisations de complément Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="c2785-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
