---
title: LaunchEvent dans le fichier manifeste (aperçu)
description: L’élément LaunchEvent configure votre module d’activation en fonction des événements pris en charge.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555310"
---
# <a name="launchevent-element-preview"></a>LaunchEvent élément (aperçu)

Configure votre module d’activation en fonction des événements pris en charge. Enfant de [`<LaunchEvents>`](launchevents.md) l’élément. Pour plus d’informations, [consultez Configurez votre Outlook pour l’activation basée sur l’événement.](../../outlook/autolaunch.md)

**Type de complément :** messagerie

> [!IMPORTANT]
> L’activation basée sur [l’événement est actuellement en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) avant-première et n’est disponible Outlook sur le Web et sur Windows. Pour plus d’informations, voir [Comment prévisualiser la fonction d’activation basée sur l’événement](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
|  **Type**  |  Oui  | Spécifie un type d’événement pris en charge. Pour l’ensemble des types pris en charge, voir [Comment prévisualiser la fonction d’activation basée sur l’événement](../../outlook/autolaunch.md#supported-events). |
|  **FunctionName**  |  Oui  | Spécifie le nom de la fonction JavaScript pour gérer l’événement spécifié dans `Type` l’attribut. |

## <a name="see-also"></a>Voir aussi

- [LaunchEvents](launchevents.md)
