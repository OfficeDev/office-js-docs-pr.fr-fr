---
title: Runtimes dans le fichier manifeste
description: L'élément Runtimes spécifie le runtime de votre add-in.
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917085"
---
# <a name="runtimes-element"></a>Élément Runtimes

Spécifie le runtime de votre add-in. Enfant de [`<Host>`](host.md) l'élément.

> [!NOTE]
> Lors de l'exécution dans Office sur Windows, un add-in qui possède un élément dans son manifeste ne s'exécute pas nécessairement dans le même contrôle webview que dans le `<Runtimes>` cas contraire. Pour plus d'informations sur la façon dont les versions de Windows et d'Office déterminent le contrôle webview utilisé normalement, voir Navigateurs utilisés par les [applications Office.](../../concepts/browsers-used-by-office-web-add-ins.md) Si les conditions décrites ici pour l'utilisation de Microsoft Edge avec WebView2 (basé sur Chromium) sont remplies, le add-in utilise ce navigateur, qu'il ait ou non un `<Runtimes>` élément. Toutefois, lorsque ces conditions ne sont pas remplies, un add-in avec un élément utilise toujours Internet Explorer 11, quelle que soit la version de Windows ou `<Runtimes>` de Microsoft 365.

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
| [Runtime](runtime.md) | Oui |  Runtime de votre add-in. |

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurer votre complément Outlook pour l'activation basée sur des événements](../../outlook/autolaunch.md)
