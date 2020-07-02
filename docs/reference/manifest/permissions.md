---
title: Élément permissions dans le fichier manifest
description: L’élément permissions spécifie le niveau d’accès à l’API pour votre complément Office.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: bc4cc2713d5a781c3407385470acd762910d17fd
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006457"
---
# <a name="permissions-element"></a><span data-ttu-id="8aada-103">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="8aada-103">Permissions element</span></span>

<span data-ttu-id="8aada-104">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="8aada-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="8aada-105">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="8aada-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8aada-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8aada-106">Syntax</span></span>

<span data-ttu-id="8aada-107">Pour les compléments du volet de tâches et de contenu :</span><span class="sxs-lookup"><span data-stu-id="8aada-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="8aada-108">Pour les compléments de messagerie :</span><span class="sxs-lookup"><span data-stu-id="8aada-108">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="8aada-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8aada-109">Contained in</span></span>

[<span data-ttu-id="8aada-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="8aada-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="8aada-111">Remarques</span><span class="sxs-lookup"><span data-stu-id="8aada-111">Remarks</span></span>

<span data-ttu-id="8aada-112">Pour plus d’informations, consultez la rubrique [demande d’autorisations pour l’utilisation d’API dans les compléments de contenu et du volet Office](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) et [Présentation des autorisations de complément Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="8aada-112">For more details, see [Requesting permissions for API use in content and task pane add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
