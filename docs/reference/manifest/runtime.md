---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet des tâches, fonctions personnalisées.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: acdff8f7ffb1e9392c1671eadc36a79348ece5fa
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138442"
---
# <a name="runtime-element"></a>Élément Runtime

Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime. Enfant de [`<Runtimes>`](runtimes.md) l’élément.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

 - Volet De tâches 1.0
 - Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (uniquement lorsqu’il est utilisé dans un add-in de volet de tâches.)

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
