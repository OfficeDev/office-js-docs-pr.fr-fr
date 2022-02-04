---
title: LaunchEvent dans le fichier manifeste
description: L’élément LaunchEvent configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="launchevent-element"></a>Élément LaunchEvent

Configure votre add-in pour qu’il s’active en fonction des événements pris en charge. Enfant de l’élément [`<LaunchEvents>`](launchevents.md) . Pour plus d’informations, [voir Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Type**  |  Oui  | Spécifie un type d’événement pris en charge. Pour l’ensemble des types pris en charge, voir [Configurer Outlook complément pour l’activation basée sur des événements](../../outlook/autolaunch.md#supported-events). |
|  **FunctionName**  |  Oui  | Spécifie le nom de la fonction JavaScript pour gérer l’événement spécifié dans l’attribut `Type` . |
|  **SendMode** (aperçu) |  Non  | Obligatoire pour et `OnMessageSend` les événements `OnAppointmentSend` . Spécifie les options disponibles pour l’utilisateur si votre add-in arrête l’envoi de l’élément. Pour les options disponibles, reportez-vous [aux options SendMode disponibles](#available-sendmode-options-preview). |

## <a name="available-sendmode-options-preview"></a>Options SendMode disponibles (aperçu)

Lorsque vous incluez l’événement `OnMessageSend` `OnAppointmentSend` ou l’événement dans le manifeste, vous devez également définir la **propriété SendMode** . Les options disponibles sont les suivantes. En fonction des conditions que recherche votre add-in, l’utilisateur est alerté si votre add-in trouve un problème dans l’élément envoyé.

| Option SendMode | Description |
|---|---|
|`PromptUser`|Dans l’alerte, l’utilisateur peut choisir d’envoyer malgré **tout**, ou de résoudre le problème, puis essayer d’envoyer à nouveau l’élément.|
|`SoftBlock`|L’utilisateur doit résoudre le problème avant d’essayer de renvoyer l’élément.|

## <a name="see-also"></a>Voir aussi

- [LaunchEvents](launchevents.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md#supported-events)
- [Utiliser les alertes intelligentes et l’événement OnMessageSend dans votre Outlook de gestion](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
