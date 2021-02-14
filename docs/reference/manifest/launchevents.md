---
title: LaunchEvents dans le fichier manifeste (aperçu)
description: L’élément LaunchEvents configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 59c52aa3f60e69e2bdda84718c6123f02942fedc
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237979"
---
# <a name="launchevents-element-preview"></a>Élément LaunchEvents (aperçu)

Configure votre add-in pour qu’il s’active en fonction des événements pris en charge. Enfant de [`<ExtensionPoint>`](extensionpoint.md) l’élément. Pour plus d’informations, voir Configurer votre complément [Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md)

**Type de complément :** messagerie

> [!IMPORTANT]
> L’activation basée sur des événements est actuellement [en prévisualisation](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) et disponible uniquement dans Outlook sur le web et sur Windows. Pour plus d’informations, [voir Comment afficher un aperçu de la fonctionnalité d’activation basée sur des événements.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

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

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Oui |  Mapz l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation du complément. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvent](launchevent.md)
