---
title: Élément type dans le fichier manifeste
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628227"
---
# <a name="type-element"></a>Élément Type

Indique si le complément équivalent est un complément COM ou un XLL.

**Type de complément :** Volet Office, fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Contenu dans

[EquivalentAdd-in](equivalentaddin.md)

## <a name="add-in-type-values"></a>Valeurs de type de complément

Vous devez spécifier l’une des valeurs suivantes pour l' `Type` élément.

- COM : spécifie que le complément équivalent est un complément COM.
- XLL : spécifie que le complément équivalent est une XLL Excel.

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)