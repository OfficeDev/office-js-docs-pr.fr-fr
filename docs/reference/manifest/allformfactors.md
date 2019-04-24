---
title: Élément AllFormFactors dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8059501f88f966b285398ac7cf243e6b0e4e44ea
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450736"
---
# <a name="allformfactors-element"></a><span data-ttu-id="0cf08-102">AllFormFactors, élément</span><span class="sxs-lookup"><span data-stu-id="0cf08-102">AllFormFactors element</span></span>

<span data-ttu-id="0cf08-103">Spécifie les paramètres d’un complément pour tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="0cf08-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="0cf08-104">Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0cf08-104">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="0cf08-105">**AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0cf08-105">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0cf08-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="0cf08-106">Child elements</span></span>

|  <span data-ttu-id="0cf08-107">Élément</span><span class="sxs-lookup"><span data-stu-id="0cf08-107">Element</span></span> |  <span data-ttu-id="0cf08-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="0cf08-108">Required</span></span>  |  <span data-ttu-id="0cf08-109">Description</span><span class="sxs-lookup"><span data-stu-id="0cf08-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0cf08-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0cf08-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="0cf08-111">Oui</span><span class="sxs-lookup"><span data-stu-id="0cf08-111">Yes</span></span> |  <span data-ttu-id="0cf08-112">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="0cf08-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="0cf08-113">Exemple AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="0cf08-113">AllFormFactors example</span></span>

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
