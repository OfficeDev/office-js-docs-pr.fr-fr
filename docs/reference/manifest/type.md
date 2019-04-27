---
title: Élément type dans le fichier manifeste
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356873"
---
# <a name="type-element"></a>Élément Type

Indique si le complément équivalent est un complément COM ou un XLL.

**Type de complément:** Volet Office, fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Contenu dans

[EquivalentAdd-in](equivalentaddin.md)

## <a name="add-in-type-values"></a>Valeurs de type de complément

Vous devez spécifier l'une des valeurs suivantes pour l' `Type` élément.

- COM: spécifie que le complément équivalent est un complément COM.
- XLL: spécifie que le complément équivalent est une XLL Excel.

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Faire en sorte que votre complément Office soit compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)