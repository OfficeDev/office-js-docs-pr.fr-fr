---
title: Élément RequestedHeight dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ea8c0403146f526b28eb20b8364fd210ac357baf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433473"
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
- [ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels) avec une valeur qui peut être entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le point d’extension **CustomPane**