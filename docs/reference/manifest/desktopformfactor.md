---
title: Élément DesktopFormFactor dans le fichier manifeste
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d1f09203518a38f1568b13e6c1a9c70752697152
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128516"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="74d4d-102">DesktopFormFactor, élément</span><span class="sxs-lookup"><span data-stu-id="74d4d-102">DesktopFormFactor element</span></span>

<span data-ttu-id="74d4d-103">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="74d4d-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="74d4d-104">Le format de bureau inclut Office sur le Web, Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="74d4d-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="74d4d-105">Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="74d4d-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="74d4d-p102">Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="74d4d-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="74d4d-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="74d4d-108">Child elements</span></span>

| <span data-ttu-id="74d4d-109">Élément</span><span class="sxs-lookup"><span data-stu-id="74d4d-109">Element</span></span>                               | <span data-ttu-id="74d4d-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="74d4d-110">Required</span></span> | <span data-ttu-id="74d4d-111">Description</span><span class="sxs-lookup"><span data-stu-id="74d4d-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="74d4d-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="74d4d-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="74d4d-113">Oui</span><span class="sxs-lookup"><span data-stu-id="74d4d-113">Yes</span></span>      | <span data-ttu-id="74d4d-114">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="74d4d-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="74d4d-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="74d4d-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="74d4d-116">Oui</span><span class="sxs-lookup"><span data-stu-id="74d4d-116">Yes</span></span>      | <span data-ttu-id="74d4d-117">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="74d4d-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="74d4d-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="74d4d-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="74d4d-119">Non</span><span class="sxs-lookup"><span data-stu-id="74d4d-119">No</span></span>       | <span data-ttu-id="74d4d-120">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="74d4d-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="74d4d-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="74d4d-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="74d4d-122">Non</span><span class="sxs-lookup"><span data-stu-id="74d4d-122">No</span></span> | <span data-ttu-id="74d4d-123">Définit si le complément Outlook est disponible dans les scénarios de délégation et est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="74d4d-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="74d4d-124">**Important**: étant donné que l’accès délégué pour les compléments Outlook est actuellement en préversion, les `SupportSharedFolders` compléments qui utilisent l’élément ne peuvent pas être publiés dans AppSource ou déployés via un déploiement centralisé.</span><span class="sxs-lookup"><span data-stu-id="74d4d-124">**Important**: Because delegate access for Outlook add-ins is currently in preview, add-ins that use the `SupportSharedFolders` element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="74d4d-125">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="74d4d-125">DesktopFormFactor example</span></span>

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
