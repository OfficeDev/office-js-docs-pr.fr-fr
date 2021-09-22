---
title: Élément DesktopFormFactor dans le fichier manifeste
description: Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3f15840a7b6716cd8acabe9e061effa566d48930
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474328"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor, élément

Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme de bureau inclut Office sur le Web, Windows et Mac. Il contient toutes les informations de l’application pour le facteur de forme de bureau, à l’exception **du** nœud Resources.

Chaque définition DesktopFormFactor contient **l’élément FunctionFile** et un ou plusieurs **éléments ExtensionPoint.** Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Éléments enfants

| Élément                               | Obligatoire | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](functionfile.md)       | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [GetStarted](getstarted.md)           | Non       | Définit la callout qui s’affiche lors de l’installation du module dans Word, Excel ou PowerPoint. Si elle est omise, la légende utilise les valeurs des éléments [DisplayName](displayname.md) et [Description](description.md) à la place. |
| [SupportsSharedFolders](supportssharedfolders.md) | Non | Définit si le Outlook est disponible dans les scénarios de boîte aux lettres partagée (désormais en prévisualisation) et de dossiers partagés (autrement dit, accès délégué). Valeur *false* par défaut. |

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
