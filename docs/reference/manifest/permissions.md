---
title: Élément permissions dans le fichier manifest
description: L’élément permissions spécifie le niveau d’accès à l’API pour votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 603494b61ef126b35cb5cdff8c5f5b911bd25840
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611490"
---
# <a name="permissions-element"></a><span data-ttu-id="27092-103">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="27092-103">Permissions element</span></span>

<span data-ttu-id="27092-104">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="27092-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="27092-105">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="27092-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="27092-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="27092-106">Syntax</span></span>

<span data-ttu-id="27092-107">Pour les compléments du volet de tâches et de contenu:</span><span class="sxs-lookup"><span data-stu-id="27092-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="27092-108">Pour les compléments de messagerie:</span><span class="sxs-lookup"><span data-stu-id="27092-108">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="27092-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="27092-109">Contained in</span></span>

[<span data-ttu-id="27092-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="27092-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="27092-111">Remarques</span><span class="sxs-lookup"><span data-stu-id="27092-111">Remarks</span></span>

<span data-ttu-id="27092-112">Pour plus d’informations, reportez-vous à la rubrique [demande d’autorisations pour l’utilisation des API dans les compléments](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) et [Présentation des autorisations de complément Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="27092-112">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
