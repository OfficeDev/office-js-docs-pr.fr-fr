---
title: Élément type dans le fichier manifeste
description: L’élément type spécifie si le complément équivalent est un complément COM ou un XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604558"
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