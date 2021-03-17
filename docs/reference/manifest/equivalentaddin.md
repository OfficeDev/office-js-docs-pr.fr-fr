---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836836"
---
# <a name="equivalentaddin-element"></a>Élément EquivalentAddin

Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.

**Type de add-in :** Volet Des tâches, Fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contenu dans

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Doit contenir

[Type (Type)](type.md)

## <a name="can-contain"></a>Peut contenir

[ProgId](progid.md) 
 [FileName](filename.md)

## <a name="remarks"></a>Remarques

Pour spécifier un compl?ment COM en tant que compl?ment équivalent, fournissez les deux `ProgId` `Type` éléments. Pour spécifier un XLL en tant que module équivalent, fournissez à la fois les `FileName` éléments et les `Type` éléments.

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Rendre votre complément Office compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)