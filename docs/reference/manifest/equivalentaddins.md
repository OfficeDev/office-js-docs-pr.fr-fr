---
title: Élément EquivalentAddins dans le fichier manifeste
description: Spécifie la compatibilité ascendante avec un add-in COM équivalent, XLL ou les deux.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: d32f67f49d334a75433aec2d079b45a44a04121a
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990809"
---
# <a name="equivalentaddins-element"></a>Élément EquivalentAddins

Spécifie la compatibilité ascendante avec un add-in COM équivalent, XLL ou les deux.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Type de add-in :** Volet Des tâches, Courrier, Fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>Contenu dans

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>Doit contenir

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Rendre votre complément Office compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)