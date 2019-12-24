---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a70d72e454273873c6a30ffd82c3a2a5194f55e0
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851305"
---
# <a name="permissions-element"></a><span data-ttu-id="798f2-102">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="798f2-102">Permissions element</span></span>

<span data-ttu-id="798f2-103">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="798f2-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="798f2-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="798f2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="798f2-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="798f2-105">Syntax</span></span>

<span data-ttu-id="798f2-106">Pour les compléments du volet de tâches et de contenu:</span><span class="sxs-lookup"><span data-stu-id="798f2-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="798f2-107">Pour les compléments de messagerie:</span><span class="sxs-lookup"><span data-stu-id="798f2-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="798f2-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="798f2-108">Contained in</span></span>

[<span data-ttu-id="798f2-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="798f2-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="798f2-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="798f2-110">Remarks</span></span>

<span data-ttu-id="798f2-111">Pour plus d’informations, reportez-vous à la rubrique [demande d’autorisations pour l’utilisation des API dans les compléments](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et [Présentation des autorisations de complément Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="798f2-111">For more detail, see [Requesting permissions for API use in add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
