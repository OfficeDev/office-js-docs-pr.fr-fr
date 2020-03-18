---
title: Élément AllFormFactors dans le fichier manifeste
description: Spécifie les paramètres d’un complément pour tous les facteurs de forme.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f1285f92b5eb89993e7fcfe79aab2325b86aca3d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720714"
---
# <a name="allformfactors-element"></a><span data-ttu-id="40bb6-103">AllFormFactors, élément</span><span class="sxs-lookup"><span data-stu-id="40bb6-103">AllFormFactors element</span></span>

<span data-ttu-id="40bb6-104">Spécifie les paramètres d’un complément pour tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="40bb6-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="40bb6-105">Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="40bb6-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="40bb6-106">**AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="40bb6-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="40bb6-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="40bb6-107">Child elements</span></span>

|  <span data-ttu-id="40bb6-108">Élément</span><span class="sxs-lookup"><span data-stu-id="40bb6-108">Element</span></span> |  <span data-ttu-id="40bb6-109">Requis</span><span class="sxs-lookup"><span data-stu-id="40bb6-109">Required</span></span>  |  <span data-ttu-id="40bb6-110">Description</span><span class="sxs-lookup"><span data-stu-id="40bb6-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="40bb6-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="40bb6-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="40bb6-112">Oui</span><span class="sxs-lookup"><span data-stu-id="40bb6-112">Yes</span></span> |  <span data-ttu-id="40bb6-113">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="40bb6-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="40bb6-114">Exemple AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="40bb6-114">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
