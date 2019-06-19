---
title: Élément EquivalentAddin dans le fichier manifeste
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059922"
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

Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgId` éléments et. `Type` Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)