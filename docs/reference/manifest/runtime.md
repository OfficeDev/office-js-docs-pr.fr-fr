---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet des tâches, fonctions personnalisées.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652243"
---
# <a name="runtime-element"></a>Élément Runtime

Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime. Enfant de [`<Runtimes>`](runtimes.md) l’élément.

**Type de add-in :** Volet De tâches, Courrier

[!include[Runtimes support](../../includes/runtimes-note.md)]

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
|  **resid**  |  Oui  | Spécifie l’emplacement URL de la page HTML de votre application. Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément. |
|  **lifetime**  |  Non  | La valeur par `lifetime` défaut est `short` et n’a pas besoin d’être spécifiée. Les add-ins Outlook utilisent uniquement la `short` valeur. Si vous souhaitez utiliser un runtime partagé dans un add-in Excel, définissez explicitement la valeur sur `long` . |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md)
