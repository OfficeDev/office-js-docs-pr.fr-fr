---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour son ruban, son volet de tâches et ses fonctions personnalisées.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217759"
---
# <a name="runtime-element"></a>Élément Runtime

Élément enfant de l' [`<Runtimes>`](runtimes.md) élément. Cet élément configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé de sorte que votre ruban, votre volet de tâches et vos fonctions personnalisées s’exécutent dans le même Runtime. Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans

- [Services d’exécution](runtimes.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Lifetime = "long"**  |  Oui  | Doit toujours être `long` utilisé pour utiliser un runtime partagé pour le complément Excel. |
|  **resid**  |  Oui  | Spécifie l’URL de la page HTML de votre complément. L' `resid` doit correspondre à un `id` attribut d’un `Url` élément dans l' `Resources` élément. |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
