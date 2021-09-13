---
title: Élément TargetDialect dans le fichier manifest
description: L’élément TargetDialect définit une langue régionale prise en charge par ce dictionnaire, représentée par une chaîne de nom de culture.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: a208b80f1a715c5ee3626f632fb757f347bdcc0a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150303"
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
