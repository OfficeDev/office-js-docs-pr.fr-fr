---
title: LaunchEvents dans le fichier manifeste
description: L’élément LaunchEvents configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590915"
---
# <a name="launchevents-element"></a>Élément LaunchEvents

Configure votre add-in pour qu’il s’active en fonction des événements pris en charge. Enfant de [`<ExtensionPoint>`](extensionpoint.md) l’élément. Pour plus d’informations, [voir Configurer Outlook complément pour l’activation basée sur des événements.](../../outlook/autolaunch.md)

**Type de complément :** messagerie

## <a name="syntax"></a>Syntaxe

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a>Contenu dans

[ExtensionPoint](extensionpoint.md) (**launchEvent** mail add-in)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Oui |  Mapz l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation du complément. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvent](launchevent.md)
