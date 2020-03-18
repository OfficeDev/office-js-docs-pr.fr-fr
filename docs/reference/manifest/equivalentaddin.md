---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 425b926901b7325665eeede04263f74e4b854d50
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718285"
---
# <a name="equivalentaddin-element"></a>Élément EquivalentAddin

Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.

**Type de complément :** Volet Office, fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contenu dans

[EquivalentAdd-ins](equivalentaddins.md)

## <a name="must-contain"></a>Doit contenir

[Type](type.md)

## <a name="can-contain"></a>Peut contenir

[ProgID](progid.md)
[nom de fichier](filename.md)

## <a name="remarks"></a>Remarques

Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgId` éléments et. `Type` Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)