---
title: Valeurs Null dans les add-ins Word
description: Découvrez comment travailler avec des valeurs null dans votre add-in Word.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e21677dafcaaaa7e9e9164ef18c82f49820298d6
ms.sourcegitcommit: 9d930b4c77c342246607aef30479e31fdbdd47f0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63353856"
---
# <a name="null-values-in-word-add-ins"></a>Valeurs Null dans les add-ins Word

`null` a des implications spéciales dans les API JavaScript pour Word. Il est utilisé pour représenter les valeurs par défaut ou aucune mise en forme.

## <a name="null-property-values-in-the-response"></a>valeurs de la propriété Null dans la réponse

Les propriétés de mise en forme telles que [la couleur](/javascript/api/word/word.font#word-word-font-color-member) contiennent `null` des valeurs dans la réponse lorsque différentes valeurs existent dans la plage [spécifiée](/javascript/api/word/word.range). Par exemple, si vous récupérez une plage et chargez sa propriété `range.font.color`:

- Si tout le texte de la plage a la même couleur de police, `range.font.color` spécifie cette couleur.
- Si plusieurs couleurs de police sont présentes dans la plage, `range.font.color` est `null`.
