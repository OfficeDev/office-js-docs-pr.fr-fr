---
title: LaunchEvents dans le fichier manifeste
description: L’élément LaunchEvents configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="launchevents-element"></a>Élément LaunchEvents

Configure votre add-in pour qu’il s’active en fonction des événements pris en charge. Enfant de l’élément [`<ExtensionPoint>`](extensionpoint.md) . Pour plus d’informations, [voir Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Type de complément :** messagerie

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

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

[ExtensionPoint](extensionpoint.md) (add-in de messagerie **LaunchEvent** )

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Oui |  Masez l’événement pris en charge à sa fonction dans le fichier JavaScript pour l’activation du complément. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvent](launchevent.md)
