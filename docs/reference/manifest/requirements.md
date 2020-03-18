---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble de conditions requises minimum et les méthodes nécessaires à l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a3f41a763ec820a6c766e6a32b26e55ad34996f7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720448"
---
# <a name="requirements-element"></a><span data-ttu-id="9edb8-103">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="9edb8-103">Requirements element</span></span>

<span data-ttu-id="9edb8-104">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="9edb8-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="9edb8-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="9edb8-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9edb8-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="9edb8-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="9edb8-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="9edb8-107">Contained in</span></span>

[<span data-ttu-id="9edb8-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9edb8-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="9edb8-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="9edb8-109">Can contain</span></span>

|<span data-ttu-id="9edb8-110">**Élément**</span><span class="sxs-lookup"><span data-stu-id="9edb8-110">**Element**</span></span>|<span data-ttu-id="9edb8-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="9edb8-111">**Content**</span></span>|<span data-ttu-id="9edb8-112">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="9edb8-112">**Mail**</span></span>|<span data-ttu-id="9edb8-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="9edb8-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="9edb8-114">Ensembles</span><span class="sxs-lookup"><span data-stu-id="9edb8-114">Sets</span></span>](sets.md)|<span data-ttu-id="9edb8-115">x</span><span class="sxs-lookup"><span data-stu-id="9edb8-115">x</span></span>|<span data-ttu-id="9edb8-116">x</span><span class="sxs-lookup"><span data-stu-id="9edb8-116">x</span></span>|<span data-ttu-id="9edb8-117">x</span><span class="sxs-lookup"><span data-stu-id="9edb8-117">x</span></span>|
|[<span data-ttu-id="9edb8-118">Méthodes</span><span class="sxs-lookup"><span data-stu-id="9edb8-118">Methods</span></span>](methods.md)|<span data-ttu-id="9edb8-119">x</span><span class="sxs-lookup"><span data-stu-id="9edb8-119">x</span></span>||<span data-ttu-id="9edb8-120">x</span><span class="sxs-lookup"><span data-stu-id="9edb8-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="9edb8-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="9edb8-121">Remarks</span></span>

<span data-ttu-id="9edb8-122">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9edb8-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
