---
title: Élément Requirements dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432703"
---
# <a name="requirements-element"></a><span data-ttu-id="72247-102">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="72247-102">Requirements element</span></span>

<span data-ttu-id="72247-103">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles des conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="72247-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="72247-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="72247-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="72247-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="72247-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="72247-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="72247-106">Contained in</span></span>

[<span data-ttu-id="72247-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="72247-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="72247-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="72247-108">Can contain</span></span>

|<span data-ttu-id="72247-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="72247-109">**Element**</span></span>|<span data-ttu-id="72247-110">**Contenu**</span><span class="sxs-lookup"><span data-stu-id="72247-110">**Content**</span></span>|<span data-ttu-id="72247-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="72247-111">**Mail**</span></span>|<span data-ttu-id="72247-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="72247-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="72247-113">Ensembles</span><span class="sxs-lookup"><span data-stu-id="72247-113">Sets</span></span>](sets.md)|<span data-ttu-id="72247-114">x</span><span class="sxs-lookup"><span data-stu-id="72247-114">x</span></span>|<span data-ttu-id="72247-115">x</span><span class="sxs-lookup"><span data-stu-id="72247-115">x</span></span>|<span data-ttu-id="72247-116">x</span><span class="sxs-lookup"><span data-stu-id="72247-116">x</span></span>|
|[<span data-ttu-id="72247-117">Méthodes</span><span class="sxs-lookup"><span data-stu-id="72247-117">Methods</span></span>](methods.md)|<span data-ttu-id="72247-118">x</span><span class="sxs-lookup"><span data-stu-id="72247-118">x</span></span>||<span data-ttu-id="72247-119">x</span><span class="sxs-lookup"><span data-stu-id="72247-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="72247-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="72247-120">Remarks</span></span>

<span data-ttu-id="72247-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="72247-121">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

