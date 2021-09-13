---
title: Élément MobileFormFactor dans le fichier manifest
description: L’élément MobileFormFactor spécifie les paramètres de facteur de forme mobile d’un module de recherche.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 619e0465ccf0c4b327956ca166aaa6195744ebee
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153664"
---
# <a name="mobileformfactor-element"></a>Élément MobileFormFactor

Spécifie les paramètres d’un complément pour le facteur de forme pour environnement mobile. Il contient toutes les informations de complément pour ce facteur de forme pour environnement mobile pour le nœud **Resources**.

Chaque **définition MobileFormFactor** contient **l’élément FunctionFile** et un ou plusieurs **éléments ExtensionPoint.** Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).

L’élément **MobileFormFactor** est défini dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.

## <a name="child-elements"></a>Éléments enfants

| Élément                             | Obligatoire | Description  |
|:------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément |
| [FunctionFile](functionfile.md)     | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|

## <a name="mobileformfactor-example"></a>Exemple MobileFormFactor

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
