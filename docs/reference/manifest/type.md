---
title: Élément Type dans le fichier manifeste
description: L’élément Type spécifie si le add-in équivalent est un compl?ment COM ou un XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937580"
---
# <a name="type-element"></a>Élément Type

Spécifie si le compl?ment équivalent est un compl?ment COM ou un XLL.

**Type de add-in :** Volet Des tâches, Fonction personnalisée

## <a name="syntax"></a>Syntaxe

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Contenu dans

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>Valeurs de type de add-in

Vous devez spécifier l’une des valeurs suivantes pour `Type` l’élément.

- COM : spécifie que le add-in équivalent est un compl?ment COM.
- XLL : spécifie que le add-in équivalent est Excel XLL.

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Rendre votre complément Office compatible avec un complément COM existant](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)