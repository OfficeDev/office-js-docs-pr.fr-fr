---
title: Élément defaultSettings dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 199acf8be888ba51fda83d159937a74685ca48e0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450624"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="2c962-102">Élément DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="2c962-102">DefaultSettings element</span></span>

<span data-ttu-id="2c962-103">Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="2c962-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="2c962-104">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="2c962-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="2c962-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="2c962-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="2c962-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="2c962-106">Contained in</span></span>

[<span data-ttu-id="2c962-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2c962-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="2c962-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="2c962-108">Can contain</span></span>

|<span data-ttu-id="2c962-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="2c962-109">**Element**</span></span>|<span data-ttu-id="2c962-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="2c962-110">**Content**</span></span>|<span data-ttu-id="2c962-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="2c962-111">**Mail**</span></span>|<span data-ttu-id="2c962-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="2c962-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2c962-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="2c962-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="2c962-114">x</span><span class="sxs-lookup"><span data-stu-id="2c962-114">x</span></span>||<span data-ttu-id="2c962-115">x</span><span class="sxs-lookup"><span data-stu-id="2c962-115">x</span></span>|
|[<span data-ttu-id="2c962-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="2c962-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="2c962-117">x</span><span class="sxs-lookup"><span data-stu-id="2c962-117">x</span></span>|||
|[<span data-ttu-id="2c962-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="2c962-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="2c962-119">x</span><span class="sxs-lookup"><span data-stu-id="2c962-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="2c962-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="2c962-120">Remarks</span></span>

<span data-ttu-id="2c962-121">L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de volet Office et de contenu. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut pour les fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="2c962-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

