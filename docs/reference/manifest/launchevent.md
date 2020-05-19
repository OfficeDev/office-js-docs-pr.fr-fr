---
title: LaunchEvent dans le fichier manifeste (aperçu)
description: L’élément LaunchEvent configure votre complément de sorte qu’il s’active en fonction des événements pris en charge.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278542"
---
# <a name="launchevent-element-preview"></a>Élément LaunchEvent (aperçu)

Configure votre complément pour qu’il s’active en fonction des événements pris en charge. Enfant de l' [`<LaunchEvents>`](launchevents.md) élément. Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Type**  |  Oui  | Spécifie un type d’événement pris en charge. Les types disponibles sont `OnNewMessageCompose` et `OnNewAppointmentOrganizer` . |
|  **FunctionName**  |  Oui  | Spécifie le nom de la fonction JavaScript permettant de gérer l’événement spécifié dans l' `Type` attribut. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvents](launchevents.md)
