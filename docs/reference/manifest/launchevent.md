---
title: LaunchEvent dans le fichier manifeste
description: L’élément LaunchEvent configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: a8ab75633d87284e02e9db9b1a71f7a8436f7daf
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681708"
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
|  **Type**  |  Oui  | Spécifie un type d’événement pris en charge. Pour l’ensemble des types pris en charge, voir [Configurer votre complément Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md#supported-events) |
|  **FunctionName**  |  Oui  | Spécifie le nom de la fonction JavaScript pour gérer l’événement spécifié dans `Type` l’attribut. |
|  **SendMode** (aperçu) |  Non  | Obligatoire pour `OnMessageSend` et les `OnAppointmentSend` événements. Spécifie les options disponibles pour l’utilisateur si votre add-in arrête l’envoi de l’élément. Pour les options disponibles, reportez-vous [aux options SendMode disponibles.](#available-sendmode-options-preview) |

## <a name="available-sendmode-options-preview"></a>Options SendMode disponibles (aperçu)

Lorsque vous incluez l’événement ou l’événement dans le manifeste, vous devez également définir `OnMessageSend` `OnAppointmentSend` la propriété **SendMode.** Les options disponibles sont les suivantes. En fonction des conditions que recherche votre add-in, l’utilisateur est alerté si votre add-in trouve un problème dans l’élément envoyé.

| Option SendMode | Description |
|---|---|
|`PromptUser`|Dans l’alerte, l’utilisateur peut choisir d’envoyer malgré **tout** ou de résoudre le problème, puis essayer de renvoyer l’élément.|
|`SoftBlock`|L’utilisateur doit résoudre le problème avant d’essayer de renvoyer l’élément.|

## <a name="see-also"></a>Voir aussi

- [LaunchEvents](launchevents.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md#supported-events)
- [Utiliser les alertes intelligentes et l’événement OnMessageSend dans votre Outlook de gestion](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
