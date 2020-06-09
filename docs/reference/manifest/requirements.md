---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble de conditions requises minimum et les méthodes nécessaires à l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611714"
---
# <a name="requirements-element"></a><span data-ttu-id="772df-103">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="772df-103">Requirements element</span></span>

<span data-ttu-id="772df-104">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="772df-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="772df-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="772df-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="772df-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="772df-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="772df-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="772df-107">Contained in</span></span>

[<span data-ttu-id="772df-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="772df-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="772df-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="772df-109">Can contain</span></span>

|<span data-ttu-id="772df-110">**Élément**</span><span class="sxs-lookup"><span data-stu-id="772df-110">**Element**</span></span>|<span data-ttu-id="772df-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="772df-111">**Content**</span></span>|<span data-ttu-id="772df-112">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="772df-112">**Mail**</span></span>|<span data-ttu-id="772df-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="772df-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="772df-114">Ensembles</span><span class="sxs-lookup"><span data-stu-id="772df-114">Sets</span></span>](sets.md)|<span data-ttu-id="772df-115">x</span><span class="sxs-lookup"><span data-stu-id="772df-115">x</span></span>|<span data-ttu-id="772df-116">x</span><span class="sxs-lookup"><span data-stu-id="772df-116">x</span></span>|<span data-ttu-id="772df-117">x</span><span class="sxs-lookup"><span data-stu-id="772df-117">x</span></span>|
|[<span data-ttu-id="772df-118">Méthodes</span><span class="sxs-lookup"><span data-stu-id="772df-118">Methods</span></span>](methods.md)|<span data-ttu-id="772df-119">x</span><span class="sxs-lookup"><span data-stu-id="772df-119">x</span></span>||<span data-ttu-id="772df-120">x</span><span class="sxs-lookup"><span data-stu-id="772df-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="772df-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="772df-121">Remarks</span></span>

<span data-ttu-id="772df-122">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="772df-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
