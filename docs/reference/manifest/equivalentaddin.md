---
title: Élément EquivalentAddin dans le fichier manifeste
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356864"
---
# <a name="equivalentaddin-element"></a>Élément EquivalentAddin

Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.

**Type de complément:** Volet Office, fonction personnalisée

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

Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgID` éléments et. `Type` Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Faire en sorte que votre complément Office soit compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)