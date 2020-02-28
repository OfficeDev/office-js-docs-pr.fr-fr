---
title: Élément Requirements dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3c4cb81ebd6a38ea311e8fcacfa6d5fcd3b26f68
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325247"
---
# <a name="requirements-element"></a><span data-ttu-id="4f2d8-102">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="4f2d8-102">Requirements element</span></span>

<span data-ttu-id="4f2d8-103">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="4f2d8-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="4f2d8-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="4f2d8-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4f2d8-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="4f2d8-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="4f2d8-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="4f2d8-106">Contained in</span></span>

[<span data-ttu-id="4f2d8-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4f2d8-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4f2d8-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="4f2d8-108">Can contain</span></span>

|<span data-ttu-id="4f2d8-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="4f2d8-109">**Element**</span></span>|<span data-ttu-id="4f2d8-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="4f2d8-110">**Content**</span></span>|<span data-ttu-id="4f2d8-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="4f2d8-111">**Mail**</span></span>|<span data-ttu-id="4f2d8-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="4f2d8-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4f2d8-113">Ensembles</span><span class="sxs-lookup"><span data-stu-id="4f2d8-113">Sets</span></span>](sets.md)|<span data-ttu-id="4f2d8-114">x</span><span class="sxs-lookup"><span data-stu-id="4f2d8-114">x</span></span>|<span data-ttu-id="4f2d8-115">x</span><span class="sxs-lookup"><span data-stu-id="4f2d8-115">x</span></span>|<span data-ttu-id="4f2d8-116">x</span><span class="sxs-lookup"><span data-stu-id="4f2d8-116">x</span></span>|
|[<span data-ttu-id="4f2d8-117">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4f2d8-117">Methods</span></span>](methods.md)|<span data-ttu-id="4f2d8-118">x</span><span class="sxs-lookup"><span data-stu-id="4f2d8-118">x</span></span>||<span data-ttu-id="4f2d8-119">x</span><span class="sxs-lookup"><span data-stu-id="4f2d8-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="4f2d8-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="4f2d8-120">Remarks</span></span>

<span data-ttu-id="4f2d8-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="4f2d8-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

