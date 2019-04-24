---
title: Élément Requirements dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450561"
---
# <a name="requirements-element"></a><span data-ttu-id="09d82-102">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="09d82-102">Requirements element</span></span>

<span data-ttu-id="09d82-103">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles des conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="09d82-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="09d82-104">**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="09d82-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="09d82-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="09d82-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="09d82-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="09d82-106">Contained in</span></span>

[<span data-ttu-id="09d82-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="09d82-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="09d82-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="09d82-108">Can contain</span></span>

|<span data-ttu-id="09d82-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="09d82-109">**Element**</span></span>|<span data-ttu-id="09d82-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="09d82-110">**Content**</span></span>|<span data-ttu-id="09d82-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="09d82-111">**Mail**</span></span>|<span data-ttu-id="09d82-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="09d82-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="09d82-113">Ensembles</span><span class="sxs-lookup"><span data-stu-id="09d82-113">Sets</span></span>](sets.md)|<span data-ttu-id="09d82-114">x</span><span class="sxs-lookup"><span data-stu-id="09d82-114">x</span></span>|<span data-ttu-id="09d82-115">x</span><span class="sxs-lookup"><span data-stu-id="09d82-115">x</span></span>|<span data-ttu-id="09d82-116">x</span><span class="sxs-lookup"><span data-stu-id="09d82-116">x</span></span>|
|[<span data-ttu-id="09d82-117">Méthodes</span><span class="sxs-lookup"><span data-stu-id="09d82-117">Methods</span></span>](methods.md)|<span data-ttu-id="09d82-118">x</span><span class="sxs-lookup"><span data-stu-id="09d82-118">x</span></span>||<span data-ttu-id="09d82-119">x</span><span class="sxs-lookup"><span data-stu-id="09d82-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="09d82-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="09d82-120">Remarks</span></span>

<span data-ttu-id="09d82-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="09d82-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

