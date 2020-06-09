---
title: Élément DesktopFormFactor dans le fichier manifeste
description: Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 46de234f2d97a9e6c7645c17a0f0a61d0c3e1a80
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612282"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="cfc8c-103">DesktopFormFactor, élément</span><span class="sxs-lookup"><span data-stu-id="cfc8c-103">DesktopFormFactor element</span></span>

<span data-ttu-id="cfc8c-104">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="cfc8c-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="cfc8c-105">Le format de bureau inclut Office sur le Web, Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="cfc8c-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="cfc8c-106">Elle contient toutes les informations de complément pour le facteur de forme de bureau, à l’exception du nœud **ressources** .</span><span class="sxs-lookup"><span data-stu-id="cfc8c-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="cfc8c-107">Chaque définition DesktopFormFactor contient l’élément **FunctionFile** et un ou plusieurs éléments **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="cfc8c-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="cfc8c-108">Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="cfc8c-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="cfc8c-109">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="cfc8c-109">Child elements</span></span>

| <span data-ttu-id="cfc8c-110">Élément</span><span class="sxs-lookup"><span data-stu-id="cfc8c-110">Element</span></span>                               | <span data-ttu-id="cfc8c-111">Requis</span><span class="sxs-lookup"><span data-stu-id="cfc8c-111">Required</span></span> | <span data-ttu-id="cfc8c-112">Description</span><span class="sxs-lookup"><span data-stu-id="cfc8c-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="cfc8c-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="cfc8c-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="cfc8c-114">Oui</span><span class="sxs-lookup"><span data-stu-id="cfc8c-114">Yes</span></span>      | <span data-ttu-id="cfc8c-115">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="cfc8c-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="cfc8c-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="cfc8c-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="cfc8c-117">Oui</span><span class="sxs-lookup"><span data-stu-id="cfc8c-117">Yes</span></span>      | <span data-ttu-id="cfc8c-118">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cfc8c-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="cfc8c-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="cfc8c-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="cfc8c-120">Non</span><span class="sxs-lookup"><span data-stu-id="cfc8c-120">No</span></span>       | <span data-ttu-id="cfc8c-121">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="cfc8c-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="cfc8c-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="cfc8c-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="cfc8c-123">Non</span><span class="sxs-lookup"><span data-stu-id="cfc8c-123">No</span></span> | <span data-ttu-id="cfc8c-124">Définit si le complément Outlook est disponible dans les scénarios de délégation et est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="cfc8c-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="cfc8c-125">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cfc8c-125">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
