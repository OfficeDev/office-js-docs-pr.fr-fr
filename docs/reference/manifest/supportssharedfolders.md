---
title: Élément SupportsSharedFolders dans le fichier manifest
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: bfbce42c7d1aa5eefab40b528c5b622aa7d2d54f
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413614"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="ac8c9-102">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="ac8c9-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="ac8c9-103">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="ac8c9-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="ac8c9-104">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="ac8c9-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="ac8c9-105">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="ac8c9-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ac8c9-106">L'accès délégué pour les compléments Outlook est actuellement en préversion et uniquement pris en charge dans les clients qui s'exécutent sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="ac8c9-106">Delegate access for Outlook add-ins is currently in preview and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="ac8c9-107">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="ac8c9-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="ac8c9-108">L’exemple suivant présente l’élément**SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="ac8c9-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
