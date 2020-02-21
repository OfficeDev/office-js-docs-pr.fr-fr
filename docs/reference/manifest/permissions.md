---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 95cb45f89e2a5b92edc29bf32d0b47fcb2dbf8ce
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165544"
---
# <a name="permissions-element"></a><span data-ttu-id="b56a9-102">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="b56a9-102">Permissions element</span></span>

<span data-ttu-id="b56a9-103">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="b56a9-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="b56a9-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="b56a9-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b56a9-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b56a9-105">Syntax</span></span>

<span data-ttu-id="b56a9-106">Pour les compléments du volet de tâches et de contenu:</span><span class="sxs-lookup"><span data-stu-id="b56a9-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="b56a9-107">Pour les compléments de messagerie:</span><span class="sxs-lookup"><span data-stu-id="b56a9-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="b56a9-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b56a9-108">Contained in</span></span>

[<span data-ttu-id="b56a9-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b56a9-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b56a9-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="b56a9-110">Remarks</span></span>

<span data-ttu-id="b56a9-111">Pour plus d’informations, reportez-vous à la rubrique [demande d’autorisations pour l’utilisation des API dans les compléments](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) et [Présentation des autorisations de complément Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="b56a9-111">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
