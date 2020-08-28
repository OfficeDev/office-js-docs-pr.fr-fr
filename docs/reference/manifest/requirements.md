---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble de conditions requises minimum et les méthodes nécessaires à l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292271"
---
# <a name="requirements-element"></a><span data-ttu-id="7a6ab-103">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="7a6ab-103">Requirements element</span></span>

<span data-ttu-id="7a6ab-104">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="7a6ab-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="7a6ab-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="7a6ab-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7a6ab-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7a6ab-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="7a6ab-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="7a6ab-107">Contained in</span></span>

[<span data-ttu-id="7a6ab-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7a6ab-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="7a6ab-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="7a6ab-109">Can contain</span></span>

|<span data-ttu-id="7a6ab-110">Élément</span><span class="sxs-lookup"><span data-stu-id="7a6ab-110">Element</span></span>|<span data-ttu-id="7a6ab-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="7a6ab-111">Content</span></span>|<span data-ttu-id="7a6ab-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="7a6ab-112">Mail</span></span>|<span data-ttu-id="7a6ab-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="7a6ab-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="7a6ab-114">Ensembles</span><span class="sxs-lookup"><span data-stu-id="7a6ab-114">Sets</span></span>](sets.md)|<span data-ttu-id="7a6ab-115">x</span><span class="sxs-lookup"><span data-stu-id="7a6ab-115">x</span></span>|<span data-ttu-id="7a6ab-116">x</span><span class="sxs-lookup"><span data-stu-id="7a6ab-116">x</span></span>|<span data-ttu-id="7a6ab-117">x</span><span class="sxs-lookup"><span data-stu-id="7a6ab-117">x</span></span>|
|[<span data-ttu-id="7a6ab-118">Méthodes</span><span class="sxs-lookup"><span data-stu-id="7a6ab-118">Methods</span></span>](methods.md)|<span data-ttu-id="7a6ab-119">x</span><span class="sxs-lookup"><span data-stu-id="7a6ab-119">x</span></span>||<span data-ttu-id="7a6ab-120">x</span><span class="sxs-lookup"><span data-stu-id="7a6ab-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="7a6ab-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="7a6ab-121">Remarks</span></span>

<span data-ttu-id="7a6ab-122">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="7a6ab-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
