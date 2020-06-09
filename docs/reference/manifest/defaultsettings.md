---
title: Élément defaultSettings dans le fichier manifeste
description: Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ace4f971d342f98d0aca5c21a7a48ceaf2563a2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611581"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="62c76-103">Élément DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="62c76-103">DefaultSettings element</span></span>

<span data-ttu-id="62c76-104">Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="62c76-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="62c76-105">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="62c76-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="62c76-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="62c76-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="62c76-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="62c76-107">Contained in</span></span>

[<span data-ttu-id="62c76-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="62c76-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="62c76-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="62c76-109">Can contain</span></span>

|<span data-ttu-id="62c76-110">**Élément**</span><span class="sxs-lookup"><span data-stu-id="62c76-110">**Element**</span></span>|<span data-ttu-id="62c76-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="62c76-111">**Content**</span></span>|<span data-ttu-id="62c76-112">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="62c76-112">**Mail**</span></span>|<span data-ttu-id="62c76-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="62c76-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="62c76-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="62c76-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="62c76-115">x</span><span class="sxs-lookup"><span data-stu-id="62c76-115">x</span></span>||<span data-ttu-id="62c76-116">x</span><span class="sxs-lookup"><span data-stu-id="62c76-116">x</span></span>|
|[<span data-ttu-id="62c76-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="62c76-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="62c76-118">x</span><span class="sxs-lookup"><span data-stu-id="62c76-118">x</span></span>|||
|[<span data-ttu-id="62c76-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="62c76-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="62c76-120">x</span><span class="sxs-lookup"><span data-stu-id="62c76-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="62c76-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="62c76-121">Remarks</span></span>

<span data-ttu-id="62c76-122">L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de contenu et du volet Office. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut des fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="62c76-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

