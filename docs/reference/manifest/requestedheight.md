---
title: Élément RequestedHeight dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e175d9012bb2f2a42fd466c35e5e28ade967d6f2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450526"
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
