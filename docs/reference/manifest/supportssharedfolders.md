---
title: Élément SupportsSharedFolders dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 776d44ec66c4e27a72e5487051bed1edf4b3dcaf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432682"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="5a9b7-102">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="5a9b7-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="5a9b7-103">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="5a9b7-103">It defines whether the add-in is available in delegate scenarios.</span></span> <span data-ttu-id="5a9b7-104">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="5a9b7-104">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="5a9b7-105">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="5a9b7-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5a9b7-106">Cet élément est disponible uniquement dans l’[ensemble d’exigence d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) par rapport à Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="5a9b7-106">This element is only available in the [Outlook add-ins Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="5a9b7-107">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="5a9b7-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="5a9b7-108">L’exemple suivant présente l’élément**SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="5a9b7-108">The following is an example of how the **Rows** element should look.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
