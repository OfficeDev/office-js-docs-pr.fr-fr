---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: f77a70681c8a12674d9e22022276e511552861ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990690"
---
# <a name="equivalentaddin-element"></a>Élément EquivalentAddin

Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Type de add-in :** Volet Des tâches, Courrier, Fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Contenu dans

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Doit contenir

[Type](type.md)

## <a name="can-contain"></a>Peut contenir

[ProgId](progid.md) 
 [FileName](filename.md)

## <a name="remarks"></a>Remarques

Pour spécifier un compl?ment COM en tant que compl?ment équivalent, fournissez les deux `ProgId` `Type` éléments. Pour spécifier un XLL en tant que module équivalent, fournissez à la fois les `FileName` éléments et les `Type` éléments.

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Rendre votre complément Office compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)