---
title: Élément version dans le fichier manifest
description: L’élément Version spécifie votre Office version du add-in.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937228"
---
# <a name="version-element"></a>Version, élément

Spécifie la version de votre complément Office. Le numéro de version peut être 1, 2, 3 ou 4 parties (par exemple, n, n.n, n.n.n ou n.n.n.n).

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Remarques

Chaque partie du numéro de version peut être un maximum de 5 chiffres.
