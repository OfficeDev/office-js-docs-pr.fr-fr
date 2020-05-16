---
title: Élément RequestedHeight dans le fichier manifeste
description: L’élément RequestedHeight spécifie la hauteur initiale (en pixels) d’un complément de contenu ou de messagerie.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: fa40043e6192e1304e67f1f96f770898b230036c
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253613"
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
- [ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être comprise entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le [point d’extension **CustomPane** (déconseillé)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)
