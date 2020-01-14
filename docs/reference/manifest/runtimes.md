---
title: Runtimes dans le fichier manifeste
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111176"
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

-[Runtimes](runtimes.md)
