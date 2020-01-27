---
title: Runtimes dans le fichier manifeste
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554005"
---
# <a name="runtimes-element"></a>Élément runtimes

Cette fonctionnalité est en aperçu. Spécifie le runtime de votre complément et permet aux fonctions personnalisées et au volet Office de partager des données globales et d’effectuer des appels de fonction. Doit suivre l' `<Host>` élément dans votre fichier manifeste.

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Runtime**     | Oui |  Le runtime de votre complément, souvent utilisé avec des fonctions personnalisées Excel.

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
