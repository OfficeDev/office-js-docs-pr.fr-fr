---
title: Runtimes dans le fichier manifeste
description: L’élément Runtimes spécifie le runtime de votre add-in.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 758bb7b830009d6691190a0279440a52da724624
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138603"
---
# <a name="runtimes-element"></a>Élément Runtimes

Spécifie le runtime de votre add-in. Enfant de [`<Host>`](host.md) l’élément.

> [!NOTE]
> Lors de l’exécution dans Office sur Windows, un add-in qui possède un élément dans son manifeste ne s’exécute pas nécessairement dans le même contrôle webview que dans le cas `<Runtimes>` contraire. Pour plus d’informations sur la façon dont les versions de Windows et de Office déterminent quel contrôle webview est normalement utilisé, voir Navigateurs utilisés par les Office des [applications.](../../concepts/browsers-used-by-office-web-add-ins.md) Si les conditions décrites ici pour l’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies, le add-in utilise ce navigateur, qu’il ait ou non un `<Runtimes>` élément. Toutefois, lorsque ces conditions ne sont pas remplies, un Microsoft 365 avec un élément utilise toujours `<Runtimes>` Internet Explorer 1 Windows 1.

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

[Host](host.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Oui |  Runtime de votre add-in. **Important**: pour le moment, vous ne pouvez définir qu’un `<Runtime>` seul élément. |

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md)
