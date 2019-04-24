---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450659"
---
# <a name="permissions-element"></a><span data-ttu-id="db86a-102">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="db86a-102">Permissions element</span></span>

<span data-ttu-id="db86a-103">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="db86a-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="db86a-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="db86a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="db86a-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="db86a-105">Syntax</span></span>

<span data-ttu-id="db86a-106">Pour les compléments du volet de tâches et de contenu:</span><span class="sxs-lookup"><span data-stu-id="db86a-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="db86a-107">Pour les compléments de messagerie:</span><span class="sxs-lookup"><span data-stu-id="db86a-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="db86a-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="db86a-108">Contained in</span></span>

[<span data-ttu-id="db86a-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="db86a-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="db86a-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="db86a-110">Remarks</span></span>

<span data-ttu-id="db86a-111">Pour plus de détails, consultez l’article relatif à la [demande d’autorisations pour utiliser des API dans des compléments de contenu et de volet Office](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et celui décrivant les [autorisations de complément Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="db86a-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
