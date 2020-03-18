---
title: Élément RequestedHeight dans le fichier manifeste
description: L’élément RequestedHeight spécifie la hauteur initiale (en pixels) d’un complément de contenu ou de messagerie.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 853d12baf290167f3e6a635201e8b5d1d0e35a51
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720455"
---
# <a name="requestedheight-element"></a><span data-ttu-id="fd34b-103">Élément RequestedHeight.</span><span class="sxs-lookup"><span data-stu-id="fd34b-103">RequestedHeight element</span></span>

<span data-ttu-id="fd34b-104">Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou d’un complément de messagerie.</span><span class="sxs-lookup"><span data-stu-id="fd34b-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="fd34b-105">**Type de complément :** contenu, messagerie</span><span class="sxs-lookup"><span data-stu-id="fd34b-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="fd34b-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="fd34b-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="fd34b-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="fd34b-107">Contained in</span></span>

- <span data-ttu-id="fd34b-108">[DefaultSettings](defaultsettings.md) (compléments de contenu) avec une valeur qui peut être comprise entre 32 et 1000</span><span class="sxs-lookup"><span data-stu-id="fd34b-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="fd34b-109">[DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450</span><span class="sxs-lookup"><span data-stu-id="fd34b-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="fd34b-110">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le point d’extension **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="fd34b-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
