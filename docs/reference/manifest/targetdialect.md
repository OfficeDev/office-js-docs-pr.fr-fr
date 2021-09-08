---
title: Élément TargetDialect dans le fichier manifest
description: L’élément TargetDialect définit une langue régionale prise en charge par ce dictionnaire, représentée par une chaîne de nom de culture.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0f60989ee5375f356343a8b3495f9c84120d467
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937238"
---
# <a name="targetdialect-element"></a>Élément TargetDialect

Définit une langue régionale prise en charge par ce dictionnaire, représentée sous forme de chaîne de nom de culture.

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<TargetDialect>
   string 
</TargetDialect>
```

## <a name="contained-in"></a>Contenu dans

[TargetDialects](targetdialects.md)

## <a name="remarks"></a>Remarques

Indiquez la valeur au format de balise de langue BCP 47, comme `en-US`.

## <a name="see-also"></a>Voir aussi

- [Créer un complément dictionnaire du volet Office](../../word/dictionary-task-pane-add-ins.md)
