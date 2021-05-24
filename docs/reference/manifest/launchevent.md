---
title: LaunchEvent dans le fichier manifeste
description: L’élément LaunchEvent configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c866a085ed6b7a33c8d7bf02d25e6ec748629e07
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591078"
---
# <a name="launchevent-element"></a>Élément LaunchEvent

Configure votre add-in pour qu’il s’active en fonction des événements pris en charge. Enfant de [`<LaunchEvents>`](launchevents.md) l’élément. Pour plus d’informations, [voir Configurer Outlook complément pour l’activation basée sur des événements.](../../outlook/autolaunch.md)

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Type**  |  Oui  | Spécifie un type d’événement pris en charge. Pour obtenir l’ensemble des types pris en charge, voir [Configurer Outlook complément pour l’activation basée sur des événements.](../../outlook/autolaunch.md#supported-events) |
|  **FunctionName**  |  Oui  | Spécifie le nom de la fonction JavaScript pour gérer l’événement spécifié dans `Type` l’attribut. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvents](launchevents.md)
