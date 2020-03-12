---
title: Élément Requirements dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 43c66118b9129c4c8ae395254ea82ef1cbcbaab1
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596458"
---
# <a name="requirements-element"></a><span data-ttu-id="bb832-102">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="bb832-102">Requirements element</span></span>

<span data-ttu-id="bb832-103">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="bb832-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="bb832-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="bb832-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bb832-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="bb832-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="bb832-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="bb832-106">Contained in</span></span>

[<span data-ttu-id="bb832-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="bb832-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="bb832-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="bb832-108">Can contain</span></span>

|<span data-ttu-id="bb832-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="bb832-109">**Element**</span></span>|<span data-ttu-id="bb832-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="bb832-110">**Content**</span></span>|<span data-ttu-id="bb832-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="bb832-111">**Mail**</span></span>|<span data-ttu-id="bb832-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="bb832-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="bb832-113">Ensembles</span><span class="sxs-lookup"><span data-stu-id="bb832-113">Sets</span></span>](sets.md)|<span data-ttu-id="bb832-114">x</span><span class="sxs-lookup"><span data-stu-id="bb832-114">x</span></span>|<span data-ttu-id="bb832-115">x</span><span class="sxs-lookup"><span data-stu-id="bb832-115">x</span></span>|<span data-ttu-id="bb832-116">x</span><span class="sxs-lookup"><span data-stu-id="bb832-116">x</span></span>|
|[<span data-ttu-id="bb832-117">Méthodes</span><span class="sxs-lookup"><span data-stu-id="bb832-117">Methods</span></span>](methods.md)|<span data-ttu-id="bb832-118">x</span><span class="sxs-lookup"><span data-stu-id="bb832-118">x</span></span>||<span data-ttu-id="bb832-119">x</span><span class="sxs-lookup"><span data-stu-id="bb832-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="bb832-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="bb832-120">Remarks</span></span>

<span data-ttu-id="bb832-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="bb832-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
