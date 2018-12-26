---
title: Élément defaultSettings dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433753"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="686ea-102">Élément DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="686ea-102">DefaultSettings element</span></span>

<span data-ttu-id="686ea-103">Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="686ea-103">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="686ea-104">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="686ea-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="686ea-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="686ea-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="686ea-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="686ea-106">Contained in</span></span>

[<span data-ttu-id="686ea-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="686ea-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="686ea-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="686ea-108">Can contain</span></span>

|<span data-ttu-id="686ea-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="686ea-109">**Element**</span></span>|<span data-ttu-id="686ea-110">**Contenu**</span><span class="sxs-lookup"><span data-stu-id="686ea-110">**Content**</span></span>|<span data-ttu-id="686ea-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="686ea-111">**Mail**</span></span>|<span data-ttu-id="686ea-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="686ea-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="686ea-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="686ea-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="686ea-114">x</span><span class="sxs-lookup"><span data-stu-id="686ea-114">x</span></span>||<span data-ttu-id="686ea-115">x</span><span class="sxs-lookup"><span data-stu-id="686ea-115">x</span></span>|
|[<span data-ttu-id="686ea-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="686ea-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="686ea-117">x</span><span class="sxs-lookup"><span data-stu-id="686ea-117">x</span></span>|||
|[<span data-ttu-id="686ea-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="686ea-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="686ea-119">x</span><span class="sxs-lookup"><span data-stu-id="686ea-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="686ea-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="686ea-120">Remarks</span></span>

<span data-ttu-id="686ea-121">L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de volet Office et de contenu. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut pour les fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="686ea-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

