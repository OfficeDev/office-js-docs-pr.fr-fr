---
title: Élément AllFormFactors dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: de7fcdce48e175d15ca6268f24082e37b2085b05
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433277"
---
# <a name="allformfactors-element"></a><span data-ttu-id="0ecc8-102">AllFormFactors, élément</span><span class="sxs-lookup"><span data-stu-id="0ecc8-102">AllFormFactors element</span></span>

<span data-ttu-id="0ecc8-103">Spécifie les paramètres d’un complément pour tous les facteurs de forme.</span><span class="sxs-lookup"><span data-stu-id="0ecc8-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="0ecc8-104">Actuellement, la seule fonctionnalité qui utilise **AllFormFactors** est celle des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0ecc8-104">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="0ecc8-105">**AllFormFactors** est un élément obligatoire lorsque vous utilisez des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0ecc8-105">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0ecc8-106">Child, élément</span><span class="sxs-lookup"><span data-stu-id="0ecc8-106">Child elements</span></span>

|  <span data-ttu-id="0ecc8-107">Élément</span><span class="sxs-lookup"><span data-stu-id="0ecc8-107">Element</span></span> |  <span data-ttu-id="0ecc8-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="0ecc8-108">Required</span></span>  |  <span data-ttu-id="0ecc8-109">Description</span><span class="sxs-lookup"><span data-stu-id="0ecc8-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0ecc8-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0ecc8-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="0ecc8-111">Oui</span><span class="sxs-lookup"><span data-stu-id="0ecc8-111">Yes</span></span> |  <span data-ttu-id="0ecc8-112">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="0ecc8-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="0ecc8-113">Exemple AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="0ecc8-113">AllFormFactors example</span></span>

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
