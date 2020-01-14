---
title: Runtime dans le fichier manifeste
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111169"
---
# <a name="runtime-element"></a>Élément Runtime

Cette fonctionnalité est en aperçu. Élément enfant de l' [`<Runtimes>`](runtime.md) élément. Cet élément facilite le partage des données globales et des appels de fonction entre des fonctions personnalisées Excel et le volet Office de votre complément. 

## <a name="contained-in"></a>Contenu dans

-[Runtimes](runtimes.md)

**Type de complément :** volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Lifetime = "long"**  |  Oui  | Doit toujours être mentionné si vous souhaitez que les fonctions personnalisées Excel fonctionnent pendant la fermeture du volet Office de votre complément. |
|  **resid**  |  Oui  | S’il est utilisé pour les fonctions personnalisées Excel `resid` , `TaskPaneAndCustomFunction.Url`le doit pointer vers. |

## <a name="see-also"></a>Voir aussi

-[Runtime](runtime.md)
