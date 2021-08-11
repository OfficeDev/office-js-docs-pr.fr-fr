---
title: Élément version dans le fichier manifest
description: L’élément Version spécifie votre Office version du add-in.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 9641153cbe6fa0284986b8dd286ba2114b32a82894bd5f8d33516e2a56c90be9
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096328"
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
