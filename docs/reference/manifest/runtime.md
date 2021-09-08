---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet des tâches, fonctions personnalisées.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936915"
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

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Override](override.md) | Non | **Outlook**: spécifie l’emplacement d’URL du fichier JavaScript dont Outlook Desktop a besoin pour les handleurs de [point d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent) **Important**: Pour le moment, vous ne pouvez définir qu’un seul élément et `<Override>` il doit être de type `javascript` .|

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid**  |  Oui  | Spécifie l’emplacement URL de la page HTML de votre application. Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément. |
|  **lifetime**  |  Non  | La valeur par `lifetime` défaut est `short` et n’a pas besoin d’être spécifiée. Outlook’utilisent que la `short` valeur. Si vous souhaitez utiliser un runtime partagé dans un Excel, définissez explicitement la valeur sur `long` . |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md)
