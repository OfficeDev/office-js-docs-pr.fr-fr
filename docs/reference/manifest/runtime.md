---
title: Temps d’exécution dans le fichier manifeste
description: L’élément Runtime configure votre module d’ajout pour utiliser un temps d’exécution JavaScript partagé pour ses différents composants, par exemple, ruban, volet de tâches, fonctions personnalisées.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555303"
---
# <a name="runtime-element"></a>Élément runtime

Configure votre module d’ajout pour utiliser un temps d’exécution JavaScript partagé afin que les différents composants s’exécutent tous dans le même temps d’exécution. Enfant de [`<Runtimes>`](runtimes.md) l’élément.

**Type d’add-in :** Volet de tâche, Courrier

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
| [Override](override.md) (aperçu) | Non | **Outlook**: Spécifie l’emplacement de l’URL du fichier JavaScript Outlook Desktop nécessite pour [les gestionnaires de points d’extension LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview) **Important**: À l’heure actuelle, vous ne pouvez définir `<Override>` qu’un seul élément et il doit être de type `javascript` .|

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid**  |  Oui  | Spécifie l’emplacement de l’URL de la page HTML pour votre module d’ajout. Le `resid` ne peut pas être plus de 32 caractères et doit correspondre à un attribut `id` d’un `Url` élément dans `Resources` l’élément. |
|  **vie**  |  Non  | La valeur par défaut `lifetime` pour est `short` et n’a pas besoin d’être spécifiée. Outlook add-ins n’utilisent que la `short` valeur. Si vous souhaitez utiliser un temps d’exécution partagé dans un Excel add-in, définissez explicitement la valeur à `long` . |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurez votre Outlook add-in pour l’activation basée sur l’événement](../../outlook/autolaunch.md)
