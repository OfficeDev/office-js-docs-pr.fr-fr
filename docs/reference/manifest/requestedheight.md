---
title: Élément RequestedHeight dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e175d9012bb2f2a42fd466c35e5e28ade967d6f2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450526"
---
# <a name="requestedheight-element"></a><span data-ttu-id="089cc-102">Élément RequestedHeight.</span><span class="sxs-lookup"><span data-stu-id="089cc-102">RequestedHeight element</span></span>

<span data-ttu-id="089cc-103">Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou d’un complément de messagerie.</span><span class="sxs-lookup"><span data-stu-id="089cc-103">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="089cc-104">**Type de complément :** contenu, messagerie</span><span class="sxs-lookup"><span data-stu-id="089cc-104">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="089cc-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="089cc-105">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="089cc-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="089cc-106">Contained in</span></span>

- <span data-ttu-id="089cc-107">[DefaultSettings](defaultsettings.md) (compléments de contenu) avec une valeur qui peut être comprise entre 32 et 1000</span><span class="sxs-lookup"><span data-stu-id="089cc-107">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="089cc-108">[DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450</span><span class="sxs-lookup"><span data-stu-id="089cc-108">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="089cc-109">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le point d’extension **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="089cc-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
