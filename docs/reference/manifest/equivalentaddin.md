---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611560"
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
 [Nom de fichier](filename.md)

## <a name="remarks"></a>Remarques

Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgId` `Type` éléments et. Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Faire en sorte que votre complément Excel soit compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)