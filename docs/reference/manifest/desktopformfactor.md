---
title: Élément DesktopFormFactor dans le fichier manifeste
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bada3cd4cff7973517aedb83235a224ef6c273eb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901961"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="d9a95-102">DesktopFormFactor, élément</span><span class="sxs-lookup"><span data-stu-id="d9a95-102">DesktopFormFactor element</span></span>

<span data-ttu-id="d9a95-103">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="d9a95-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="d9a95-104">Le format de bureau inclut Office sur le Web, Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="d9a95-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="d9a95-105">Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="d9a95-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="d9a95-p102">Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="d9a95-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="d9a95-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d9a95-108">Child elements</span></span>

| <span data-ttu-id="d9a95-109">Élément</span><span class="sxs-lookup"><span data-stu-id="d9a95-109">Element</span></span>                               | <span data-ttu-id="d9a95-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d9a95-110">Required</span></span> | <span data-ttu-id="d9a95-111">Description</span><span class="sxs-lookup"><span data-stu-id="d9a95-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="d9a95-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="d9a95-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="d9a95-113">Oui</span><span class="sxs-lookup"><span data-stu-id="d9a95-113">Yes</span></span>      | <span data-ttu-id="d9a95-114">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="d9a95-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="d9a95-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="d9a95-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="d9a95-116">Oui</span><span class="sxs-lookup"><span data-stu-id="d9a95-116">Yes</span></span>      | <span data-ttu-id="d9a95-117">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d9a95-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="d9a95-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="d9a95-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="d9a95-119">Non</span><span class="sxs-lookup"><span data-stu-id="d9a95-119">No</span></span>       | <span data-ttu-id="d9a95-120">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d9a95-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="d9a95-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="d9a95-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="d9a95-122">Non</span><span class="sxs-lookup"><span data-stu-id="d9a95-122">No</span></span> | <span data-ttu-id="d9a95-123">Définit si le complément Outlook est disponible dans les scénarios de délégation et est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="d9a95-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="d9a95-124">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="d9a95-124">DesktopFormFactor example</span></span>

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
