---
title: Élément DesktopFormFactor dans le fichier manifeste
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: b46536886d59692d03976083412a8b8d2e6ae859
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952389"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor, élément

Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le format de bureau inclut Office sur Windows, Office pour Mac et Office Online. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.

Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Éléments enfants

| Élément                               | Obligatoire | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](functionfile.md)       | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [GetStarted](getstarted.md)           | Non       | Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Non | Définit si le complément Outlook est disponible dans les scénarios de délégation et est défini sur *false* par défaut.<br><br>**Important**: étant donné que l’accès délégué pour les compléments Outlook est actuellement en préversion, les `SupportSharedFolders` compléments qui utilisent l’élément ne peuvent pas être publiés dans AppSource ou déployés via un déploiement centralisé. |

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
