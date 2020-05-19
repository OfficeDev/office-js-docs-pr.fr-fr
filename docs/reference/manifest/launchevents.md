---
title: LaunchEvents dans le fichier manifeste (aperçu)
description: L’élément LaunchEvents configure votre complément de sorte qu’il s’active en fonction des événements pris en charge.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278541"
---
# <a name="launchevents-element-preview"></a>Élément LaunchEvents (aperçu)

Configure votre complément pour qu’il s’active en fonction des événements pris en charge. Enfant de l' [`<ExtensionPoint>`](extensionpoint.md) élément. Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).

**Type de complément :** messagerie

> [!IMPORTANT]
> L’activation basée sur les événements est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web. Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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

[ExtensionPoint](extensionpoint.md) (complément de messagerie**LaunchEvent** )

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Oui |  Mappez l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation des compléments. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvent](launchevent.md)
