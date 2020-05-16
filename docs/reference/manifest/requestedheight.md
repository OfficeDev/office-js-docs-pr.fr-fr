---
title: Élément RequestedHeight dans le fichier manifeste
description: L’élément RequestedHeight spécifie la hauteur initiale (en pixels) d’un complément de contenu ou de messagerie.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: fa40043e6192e1304e67f1f96f770898b230036c
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253613"
---
# <a name="requestedheight-element"></a><span data-ttu-id="07232-103">Élément RequestedHeight.</span><span class="sxs-lookup"><span data-stu-id="07232-103">RequestedHeight element</span></span>

<span data-ttu-id="07232-104">Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou d’un complément de messagerie.</span><span class="sxs-lookup"><span data-stu-id="07232-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="07232-105">**Type de complément :** contenu, messagerie</span><span class="sxs-lookup"><span data-stu-id="07232-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="07232-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="07232-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="07232-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="07232-107">Contained in</span></span>

- <span data-ttu-id="07232-108">[DefaultSettings](defaultsettings.md) (compléments de contenu) avec une valeur qui peut être comprise entre 32 et 1000</span><span class="sxs-lookup"><span data-stu-id="07232-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="07232-109">[DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450</span><span class="sxs-lookup"><span data-stu-id="07232-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="07232-110">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être comprise entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le [point d’extension **CustomPane** (déconseillé)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span><span class="sxs-lookup"><span data-stu-id="07232-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>
