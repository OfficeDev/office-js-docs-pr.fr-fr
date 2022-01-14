---
title: Élément EquivalentAddins dans le fichier manifeste
description: Spécifie la compatibilité ascendante avec un compl?ment COM, une XLL ou les deux.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48f3ef86f71ad3d4f0c759df4583af4cd95e5c5a
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042153"
---
# <a name="equivalentaddins-element"></a>Élément EquivalentAddins

Spécifie la compatibilité ascendante avec un compl?ment COM, une XLL ou les deux.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Type de add-in :** Volet Des tâches, Courrier, Fonction personnalisée

**Valide uniquement dans ces schémas VersionOverrides**:

- Volet De tâches 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

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