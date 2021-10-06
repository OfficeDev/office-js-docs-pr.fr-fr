---
title: Élément DesktopFormFactor dans le fichier manifeste
description: Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 52c9a029e3f43e9b7d5416455eb99ef3de4dae7a
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138729"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor, élément

Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme de bureau inclut Office sur le Web, Windows et Mac. Il contient toutes les informations de l’application pour le facteur de forme de bureau, à l’exception **du** nœud Resources.

Chaque définition DesktopFormFactor contient **l’élément FunctionFile** et un ou plusieurs **éléments ExtensionPoint.** Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

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
