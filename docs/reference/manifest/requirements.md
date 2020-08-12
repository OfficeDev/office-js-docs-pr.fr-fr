---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble de conditions requises minimum et les méthodes nécessaires à l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c6a9a7b5923401fc2551f239b2c6cbc0d1e90755
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641318"
---
# <a name="requirements-element"></a><span data-ttu-id="e0094-103">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="e0094-103">Requirements element</span></span>

<span data-ttu-id="e0094-104">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="e0094-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="e0094-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="e0094-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e0094-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e0094-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="e0094-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="e0094-107">Contained in</span></span>

[<span data-ttu-id="e0094-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e0094-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e0094-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="e0094-109">Can contain</span></span>

|<span data-ttu-id="e0094-110">Élément</span><span class="sxs-lookup"><span data-stu-id="e0094-110">Element</span></span>|<span data-ttu-id="e0094-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="e0094-111">Content</span></span>|<span data-ttu-id="e0094-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="e0094-112">Mail</span></span>|<span data-ttu-id="e0094-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="e0094-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e0094-114">Ensembles</span><span class="sxs-lookup"><span data-stu-id="e0094-114">Sets</span></span>](sets.md)|<span data-ttu-id="e0094-115">x</span><span class="sxs-lookup"><span data-stu-id="e0094-115">x</span></span>|<span data-ttu-id="e0094-116">x</span><span class="sxs-lookup"><span data-stu-id="e0094-116">x</span></span>|<span data-ttu-id="e0094-117">x</span><span class="sxs-lookup"><span data-stu-id="e0094-117">x</span></span>|
|[<span data-ttu-id="e0094-118">Méthodes</span><span class="sxs-lookup"><span data-stu-id="e0094-118">Methods</span></span>](methods.md)|<span data-ttu-id="e0094-119">x</span><span class="sxs-lookup"><span data-stu-id="e0094-119">x</span></span>||<span data-ttu-id="e0094-120">x</span><span class="sxs-lookup"><span data-stu-id="e0094-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="e0094-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="e0094-121">Remarks</span></span>

<span data-ttu-id="e0094-122">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e0094-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
