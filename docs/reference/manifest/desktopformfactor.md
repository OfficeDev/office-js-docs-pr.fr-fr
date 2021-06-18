---
title: Élément DesktopFormFactor dans le fichier manifeste
description: Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 66673d83fd8608a1ec10492d7a944b0515de61c0
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007788"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="2a561-103">DesktopFormFactor, élément</span><span class="sxs-lookup"><span data-stu-id="2a561-103">DesktopFormFactor element</span></span>

<span data-ttu-id="2a561-104">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.</span><span class="sxs-lookup"><span data-stu-id="2a561-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="2a561-105">Le facteur de forme de bureau inclut Office sur le Web, Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="2a561-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="2a561-106">Il contient toutes les informations de l’application pour le facteur de forme de bureau, à l’exception **du** nœud Resources.</span><span class="sxs-lookup"><span data-stu-id="2a561-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="2a561-107">Chaque définition DesktopFormFactor contient **l’élément FunctionFile** et un ou plusieurs **éléments ExtensionPoint.**</span><span class="sxs-lookup"><span data-stu-id="2a561-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="2a561-108">Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="2a561-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="2a561-109">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="2a561-109">Child elements</span></span>

| <span data-ttu-id="2a561-110">Élément</span><span class="sxs-lookup"><span data-stu-id="2a561-110">Element</span></span>                               | <span data-ttu-id="2a561-111">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2a561-111">Required</span></span> | <span data-ttu-id="2a561-112">Description</span><span class="sxs-lookup"><span data-stu-id="2a561-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="2a561-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="2a561-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="2a561-114">Oui</span><span class="sxs-lookup"><span data-stu-id="2a561-114">Yes</span></span>      | <span data-ttu-id="2a561-115">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="2a561-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="2a561-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="2a561-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="2a561-117">Oui</span><span class="sxs-lookup"><span data-stu-id="2a561-117">Yes</span></span>      | <span data-ttu-id="2a561-118">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2a561-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="2a561-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="2a561-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="2a561-120">Non</span><span class="sxs-lookup"><span data-stu-id="2a561-120">No</span></span>       | <span data-ttu-id="2a561-121">Définit la callout qui s’affiche lors de l’installation du module dans Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2a561-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="2a561-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="2a561-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="2a561-123">Non</span><span class="sxs-lookup"><span data-stu-id="2a561-123">No</span></span> | <span data-ttu-id="2a561-124">Définit si le Outlook est disponible dans les scénarios de boîte aux lettres partagée (désormais en prévisualisation) et de dossiers partagés (autrement dit, accès délégué).</span><span class="sxs-lookup"><span data-stu-id="2a561-124">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="2a561-125">Valeur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="2a561-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="2a561-126">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="2a561-126">DesktopFormFactor example</span></span>

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
