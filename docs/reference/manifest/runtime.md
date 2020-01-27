---
title: Runtime dans le fichier manifeste
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41553998"
---
# <a name="runtime-element"></a>Élément Runtime

Cette fonctionnalité est en aperçu. Élément enfant de l' [`<Runtimes>`](runtimes.md) élément. Cet élément facilite le partage des données globales et des appels de fonction entre des fonctions personnalisées Excel et le volet Office de votre complément.

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans

- [Services d’exécution](runtimes.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Lifetime = "long"**  |  Oui  | Doit toujours être mentionné si vous souhaitez que les fonctions personnalisées Excel fonctionnent pendant la fermeture du volet Office de votre complément. |
|  **resid**  |  Oui  | S’il est utilisé pour les fonctions personnalisées Excel `resid` , `TaskPaneAndCustomFunction.Url`le doit pointer vers. |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
