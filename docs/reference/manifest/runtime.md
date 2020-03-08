---
title: Runtime dans le fichier manifeste (aperçu)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561827"
---
# <a name="runtime-element-preview"></a>Élément Runtime (aperçu)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Élément enfant de l' [`<Runtimes>`](runtimes.md) élément. Cet élément configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé de sorte que votre ruban, votre volet de tâches et vos fonctions personnalisées s’exécutent dans le même Runtime. Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

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

- [Services d’exécution](runtimes.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Lifetime = "long"**  |  Oui  | Doit toujours être `long` utilisé pour utiliser un runtime partagé pour le complément Excel. |
|  **resid**  |  Oui  | Spécifie l’URL de la page HTML de votre complément. L `resid` 'doit correspondre `id` à un attribut `Url` d’un élément `Resources` dans l’élément. |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
