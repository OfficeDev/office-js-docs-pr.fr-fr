---
title: Runtimes dans le fichier manifeste
description: L’élément Runtimes spécifie le runtime de votre add-in.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: e5ec70449d3984671d507131ac8d4fc0f7617cdcda1ad8f99b4f4bf52773aded
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091627"
---
# <a name="runtimes-element"></a>Élément Runtimes

Spécifie le runtime de votre add-in. Enfant de [`<Host>`](host.md) l’élément.

> [!NOTE]
> Lors de l’exécution dans Office sur Windows, un add-in qui possède un élément dans son manifeste ne s’exécute pas nécessairement dans le même contrôle webview que dans le cas `<Runtimes>` contraire. Pour plus d’informations sur la façon dont les versions de Windows et de Office déterminent quel contrôle webview est normalement utilisé, voir Navigateurs utilisés par les Office des [applications.](../../concepts/browsers-used-by-office-web-add-ins.md) Si les conditions décrites ici pour l’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies, le add-in utilise ce navigateur, qu’il ait ou non un `<Runtimes>` élément. Toutefois, lorsque ces conditions ne sont pas remplies, un Microsoft 365 avec un élément utilise toujours `<Runtimes>` Internet Explorer 1 Windows 1.

**Type de add-in :** Volet De tâches, Courrier

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
