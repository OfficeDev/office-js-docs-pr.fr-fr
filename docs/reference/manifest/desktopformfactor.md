---
title: Élément DesktopFormFactor dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433739"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="d7872-102">DesktopFormFactor, élément</span><span class="sxs-lookup"><span data-stu-id="d7872-102">DesktopFormFactor element</span></span>

<span data-ttu-id="d7872-p101">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme pour bureau inclut Office pour Windows, Office pour Mac et Office Online. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="d7872-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="d7872-p102">Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="d7872-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="d7872-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="d7872-108">Child elements</span></span>

| <span data-ttu-id="d7872-109">Élément</span><span class="sxs-lookup"><span data-stu-id="d7872-109">Element</span></span>                               | <span data-ttu-id="d7872-110">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="d7872-110">Required</span></span> | <span data-ttu-id="d7872-111">Description</span><span class="sxs-lookup"><span data-stu-id="d7872-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="d7872-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="d7872-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="d7872-113">Oui</span><span class="sxs-lookup"><span data-stu-id="d7872-113">Yes</span></span>      | <span data-ttu-id="d7872-114">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="d7872-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="d7872-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="d7872-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="d7872-116">Oui</span><span class="sxs-lookup"><span data-stu-id="d7872-116">Yes</span></span>      | <span data-ttu-id="d7872-117">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d7872-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="d7872-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="d7872-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="d7872-119">Non</span><span class="sxs-lookup"><span data-stu-id="d7872-119">No</span></span>       | <span data-ttu-id="d7872-120">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d7872-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="d7872-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="d7872-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="d7872-122">Non</span><span class="sxs-lookup"><span data-stu-id="d7872-122">No</span></span> | <span data-ttu-id="d7872-123">Définit si le complément Outlook est disponible dans les scénarios de délégation et est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="d7872-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="d7872-124">**Important **: cet élément est disponible uniquement dans l’ensemble d’exigences d’aperçu des compléments Outlook par rapport à Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="d7872-124">**Important**: This element is only available in the Outlook add-ins Preview requirement set against Exchange Online.</span></span> <span data-ttu-id="d7872-125">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="d7872-125">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="d7872-126">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="d7872-126">DesktopFormFactor example</span></span>

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
