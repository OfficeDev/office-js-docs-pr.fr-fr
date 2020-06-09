---
title: Élément AllFormFactors dans le fichier manifeste
description: Spécifie les paramètres d’un complément pour tous les facteurs de forme.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608795"
---
# <a name="allformfactors-element"></a><span data-ttu-id="0d1d5-103">AllFormFactors, élément</span><span class="sxs-lookup"><span data-stu-id="0d1d5-103">AllFormFactors element</span></span>

<span data-ttu-id="0d1d5-104">Spécifie les paramètres d’un complément pour tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="0d1d5-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="0d1d5-105">Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0d1d5-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="0d1d5-106">**AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0d1d5-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0d1d5-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="0d1d5-107">Child elements</span></span>

|  <span data-ttu-id="0d1d5-108">Élément</span><span class="sxs-lookup"><span data-stu-id="0d1d5-108">Element</span></span> |  <span data-ttu-id="0d1d5-109">Requis</span><span class="sxs-lookup"><span data-stu-id="0d1d5-109">Required</span></span>  |  <span data-ttu-id="0d1d5-110">Description</span><span class="sxs-lookup"><span data-stu-id="0d1d5-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0d1d5-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0d1d5-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="0d1d5-112">Oui</span><span class="sxs-lookup"><span data-stu-id="0d1d5-112">Yes</span></span> |  <span data-ttu-id="0d1d5-113">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="0d1d5-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="0d1d5-114">Exemple AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="0d1d5-114">AllFormFactors example</span></span>

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
