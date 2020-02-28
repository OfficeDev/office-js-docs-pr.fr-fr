---
title: Élément DesktopFormFactor dans le fichier manifeste
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 2fe97d99ff5bdc9f23a5760824e241ee4dfb800f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325275"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor, élément

Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le format de bureau inclut Office sur le Web, Windows et Mac. Elle contient toutes les informations de complément pour le facteur de forme de bureau, à l’exception du nœud **ressources** .

Chaque définition DesktopFormFactor contient l’élément **FunctionFile** et un ou plusieurs éléments **ExtensionPoint** . Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Éléments enfants

| Élément                               | Obligatoire | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](functionfile.md)       | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [GetStarted](getstarted.md)           | Non       | Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Non | Définit si le complément Outlook est disponible dans les scénarios de délégation et est défini sur *false* par défaut. |

## <a name="desktopformfactor-example"></a>Exemple pour DesktopFormFactor

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
