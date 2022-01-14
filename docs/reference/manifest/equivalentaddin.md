---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: e318a9028ebefdeca9aaf5baac465a1ec1af0a73
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042132"
---
# <a name="equivalentaddin-element"></a>Élément EquivalentAddin

Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Type de add-in :** Volet Des tâches, Courrier, Fonction personnalisée

**Valide uniquement dans ces schémas VersionOverrides**:

- Volet De tâches 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

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