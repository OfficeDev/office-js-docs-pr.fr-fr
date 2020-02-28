---
title: Élément defaultSettings dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 824c575b39a99c6028ffd603390d2b41ee0ad7dd
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324883"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="bd431-102">Élément DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="bd431-102">DefaultSettings element</span></span>

<span data-ttu-id="bd431-103">Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="bd431-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="bd431-104">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="bd431-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="bd431-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="bd431-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="bd431-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="bd431-106">Contained in</span></span>

[<span data-ttu-id="bd431-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="bd431-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="bd431-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="bd431-108">Can contain</span></span>

|<span data-ttu-id="bd431-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="bd431-109">**Element**</span></span>|<span data-ttu-id="bd431-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="bd431-110">**Content**</span></span>|<span data-ttu-id="bd431-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="bd431-111">**Mail**</span></span>|<span data-ttu-id="bd431-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="bd431-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="bd431-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bd431-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="bd431-114">x</span><span class="sxs-lookup"><span data-stu-id="bd431-114">x</span></span>||<span data-ttu-id="bd431-115">x</span><span class="sxs-lookup"><span data-stu-id="bd431-115">x</span></span>|
|[<span data-ttu-id="bd431-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="bd431-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="bd431-117">x</span><span class="sxs-lookup"><span data-stu-id="bd431-117">x</span></span>|||
|[<span data-ttu-id="bd431-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="bd431-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="bd431-119">x</span><span class="sxs-lookup"><span data-stu-id="bd431-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="bd431-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="bd431-120">Remarks</span></span>

<span data-ttu-id="bd431-121">L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de contenu et du volet Office. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut des fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="bd431-121">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

