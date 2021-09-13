---
title: Élément RequestedHeight dans le fichier manifeste
description: L’élément RequestedHeight spécifie la hauteur initiale (en pixels) d’un contenu ou d’un module de messagerie.
ms.date: 05/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: e0589e81e8905c4fc8c7a8e50ec7c14038035677
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150380"
---
# <a name="requestedheight-element"></a>Élément RequestedHeight.

Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou d’un complément de messagerie.

**Type de complément :** contenu, messagerie

## <a name="syntax"></a>Syntaxe

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (compléments de contenu) avec une valeur qui peut être comprise entre 32 et 1000
- [DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450
- [ExtensionPoint](extensionpoint.md) (modules de messagerie contextuels) avec une valeur qui peut être entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le point d’extension [ **CustomPane** (supprimé)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)
