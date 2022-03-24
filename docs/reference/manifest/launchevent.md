---
title: LaunchEvent dans le fichier manifeste
description: L’élément LaunchEvent configure votre add-in pour qu’il s’active en fonction des événements pris en charge.
ms.date: 03/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 71469693bff7213455582a3247778cabf92c2aa3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745806"
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
|  **SendMode** (aperçu) |  Non  | Utilisé par et `OnMessageSend` événements `OnAppointmentSend` . Spécifie les options disponibles pour l’utilisateur si votre add-in arrête l’envoi d’un élément ou s’il n’est pas disponible. Si la **propriété SendMode** n’est pas incluse, l’option `SoftBlock` est définie par défaut. Pour les options disponibles, reportez-vous [aux options SendMode disponibles](#available-sendmode-options-preview). |

## <a name="available-sendmode-options-preview"></a>Options SendMode disponibles (aperçu)

Lorsque vous incluez l’événement `OnMessageSend` `OnAppointmentSend` ou l’événement dans le manifeste, vous devez également définir la **propriété SendMode** . Si la **propriété SendMode** n’est pas incluse, l’option `SoftBlock` est définie par défaut. Les options disponibles sont les suivantes. En fonction des conditions que recherche votre add-in, l’utilisateur est alerté si votre add-in trouve un problème dans l’élément envoyé.

| Option SendMode | Description |
|---|---|
|`PromptUser`|Si l’élément ne répond pas aux conditions du module, l’utilisateur peut choisir Envoyer quand  même dans l’alerte, ou résoudre le problème, puis essayer de renvoyer l’élément. Si le traitement de l’élément par le module prend beaucoup de temps, l’utilisateur est invité à arrêter l’exécution du module et à choisir Envoyer quand **même**. En cas d’indisponibilité du module (par exemple, une erreur de chargement du module), l’élément est envoyé.|
|`SoftBlock`|Option par défaut si **la propriété SendMode** n’est pas incluse. L’utilisateur est averti que l’élément qu’il envoie ne répond pas aux conditions du module et qu’il doit résoudre le problème avant d’essayer de renvoyer l’élément. Toutefois, si le add-in n’est pas disponible (par exemple, une erreur de chargement du module), l’élément est envoyé.|
|`Block`|L’élément n’est pas envoyé si l’une des situations suivantes se produit.<br>- L’élément ne répond pas aux conditions du module.<br>- Le add-in ne parvient pas à se connecter au serveur.<br>- Une erreur s’est produite lors du chargement du module.|

## <a name="see-also"></a>Voir aussi

- [LaunchEvents](launchevents.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md#supported-events)
- [Utiliser les alertes intelligentes et l’événement OnMessageSend dans votre Outlook de gestion](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
