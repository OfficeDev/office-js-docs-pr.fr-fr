---
title: Élément RequestedHeight dans le fichier manifeste
description: L’élément RequestedHeight spécifie la hauteur initiale (en pixels) d’un complément de contenu ou de messagerie.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f4c3ca1ff39cc3150249fbc824b0db76f6b8a85
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215039"
---
# <a name="requestedheight-element"></a><span data-ttu-id="8ef5d-103">Élément RequestedHeight.</span><span class="sxs-lookup"><span data-stu-id="8ef5d-103">RequestedHeight element</span></span>

<span data-ttu-id="8ef5d-104">Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou d’un complément de messagerie.</span><span class="sxs-lookup"><span data-stu-id="8ef5d-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="8ef5d-105">**Type de complément :** contenu, messagerie</span><span class="sxs-lookup"><span data-stu-id="8ef5d-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8ef5d-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8ef5d-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="8ef5d-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8ef5d-107">Contained in</span></span>

- <span data-ttu-id="8ef5d-108">[DefaultSettings](defaultsettings.md) (compléments de contenu) avec une valeur qui peut être comprise entre 32 et 1000</span><span class="sxs-lookup"><span data-stu-id="8ef5d-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="8ef5d-109">[DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450</span><span class="sxs-lookup"><span data-stu-id="8ef5d-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="8ef5d-110">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le point d’extension **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="8ef5d-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
