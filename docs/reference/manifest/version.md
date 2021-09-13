---
title: Élément version dans le fichier manifest
description: L’élément Version spécifie votre Office version du add-in.
ms.date: 02/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 34cefa22123ed4ee723d51a669e01e042efc2934
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153504"
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
