---
title: Élément Icon dans le fichier manifeste
description: Définit les éléments Image pour les contrôles de bouton ou de menu.
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: f47f35f18995b3d9e0af1115668b43a506e830d8
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152267"
---
# <a name="icon-element"></a>Icon, élément

Définit les éléments **Image** pour les contrôles de [bouton](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls).

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Non  | Type d’icône en cours de définition. Uniquement applicable aux icônes dans des facteurs de forme pour environnement mobile. Pour les éléments **Icon** contenus dans un élément [MobileFormFactor](mobileformfactor.md), cet attribut doit être défini sur `bt:MobileIconList`. |

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Image](#image)        | Oui |   Attribut resid d’une image à utiliser         |

### <a name="image"></a>Image

Image du bouton. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de l’attribut **id** d’un élément **Image** dans l’élément **Images** dans l’élément [Resources.](resources.md) L’attribut **size** indique la taille de l’image en pixels. Trois tailles d’image sont requises (16, 32 et 80 pixels) et cinq autres tailles sont prises en charge (20, 24, 40, 48 et 64 pixels).

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> Si cette image est l’icône représentative de votre application, voir Créer des listes efficaces dans [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) et dans Office pour la taille et d’autres exigences.

## <a name="additional-requirements-for-mobile-form-factors"></a>Configuration requise supplémentaire pour les facteurs de forme pour environnement mobile

Lorsque l’élément **Icon** parent est un descendant de l’élément [MobileFormFactor](mobileformfactor.md), la taille minimale requise est légèrement différente. Le manifeste doit fournir au minimum les tailles 25, 32 et 48 pixels. Chaque taille fournie doit apparaître trois fois, avec un ensemble d’attributs `scale` défini sur `1`, `2` ou `3`.

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
