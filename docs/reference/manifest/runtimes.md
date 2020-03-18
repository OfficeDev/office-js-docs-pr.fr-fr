---
title: Runtimes dans le fichier manifeste (aperçu)
description: L’élément runtimes spécifie le runtime de votre complément.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720420"
---
# <a name="runtimes-element-preview"></a>Runtimes, élément (aperçu)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Spécifie le runtime de votre complément et active des fonctions personnalisées, des boutons du ruban et le volet des tâches pour utiliser le même Runtime JavaScript. Enfant de l' `<Host>` élément dans votre fichier manifeste. Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Type de complément :** volet Office

> [!IMPORTANT]
> Le runtime partagé est actuellement en préversion et n’est disponible que sur Excel sur Windows. Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans 
[Host](./host.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  **Runtime**     | Oui |  Le runtime de votre complément.

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
