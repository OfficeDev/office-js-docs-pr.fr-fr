---
title: Temps d’exécution dans le fichier manifeste
description: L’élément Runtimes spécifie le temps d’exécution de votre module.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555296"
---
# <a name="runtimes-element"></a>Élément runtimes

Spécifie le temps d’exécution de votre module d’exécution. Enfant de [`<Host>`](host.md) l’élément.

> [!NOTE]
> Lors de l’exécution Office sur Windows, un add-in qui a un élément dans son manifeste ne fonctionne `<Runtimes>` pas nécessairement dans le même contrôle webview qu’il le ferait autrement. Pour plus d’informations sur la façon dont les versions de Windows et Office déterminent quel contrôle webview est normalement utilisé, [voir Navigateurs utilisés par Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Si les conditions décrites pour l’utilisation de Microsoft Edge avec WebView2 (basé sur Chromium) sont remplies, alors l’add-in utilise ce navigateur, qu’il ait ou non un `<Runtimes>` élément. Toutefois, lorsque ces conditions ne sont pas remplies, un module avec un `<Runtimes>` élément utilise toujours Internet Explorer 11 indépendamment de la version Windows ou Microsoft 365 version.

**Type d’add-in :** Volet de tâche, Courrier

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
| [Runtime](runtime.md) | Oui |  Le temps d’exécution de votre add-in. **Important :** À l’heure actuelle, vous ne pouvez définir qu’un `<Runtime>` seul élément. |

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurez votre Outlook add-in pour l’activation basée sur l’événement](../../outlook/autolaunch.md)
