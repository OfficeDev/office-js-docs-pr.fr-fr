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
# <a name="supportssharedfolders-element"></a>Élément SupportsSharedFolders

Définit si le complément Outlook est disponible dans les scénarios de délégué. L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md). Ce paramètre est défini sur *false* par défaut.

> [!IMPORTANT]
> L'accès délégué pour les compléments Outlook est actuellement en préversion et uniquement pris en charge dans les clients qui s'exécutent sur Exchange Online. Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité déploiement centralisée.

L’exemple suivant présente l’élément**SupportsSharedFolders**.

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
