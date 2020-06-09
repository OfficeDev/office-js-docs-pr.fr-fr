---
title: Élément RequestedHeight dans le fichier manifeste
description: L’élément RequestedHeight spécifie la hauteur initiale (en pixels) d’un complément de contenu ou de messagerie.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 44675918a4208683f442fe8a6e8f4f906f484571
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611728"
---
# <a name="requestedheight-element"></a><span data-ttu-id="6e4ec-103">Élément RequestedHeight.</span><span class="sxs-lookup"><span data-stu-id="6e4ec-103">RequestedHeight element</span></span>

<span data-ttu-id="6e4ec-104">Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou d’un complément de messagerie.</span><span class="sxs-lookup"><span data-stu-id="6e4ec-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="6e4ec-105">**Type de complément :** contenu, messagerie</span><span class="sxs-lookup"><span data-stu-id="6e4ec-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6e4ec-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6e4ec-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="6e4ec-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="6e4ec-107">Contained in</span></span>

- <span data-ttu-id="6e4ec-108">[DefaultSettings](defaultsettings.md) (compléments de contenu) avec une valeur qui peut être comprise entre 32 et 1000</span><span class="sxs-lookup"><span data-stu-id="6e4ec-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="6e4ec-109">[DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450</span><span class="sxs-lookup"><span data-stu-id="6e4ec-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="6e4ec-110">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être comprise entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le [point d’extension **CustomPane** (déconseillé)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span><span class="sxs-lookup"><span data-stu-id="6e4ec-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>
