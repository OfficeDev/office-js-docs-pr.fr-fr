---
title: Élément Event dans le fichier manifeste
description: Définit un gestionnaire d’événements dans un complément.
ms.date: 01/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: fac920fc91abd908d3d159877c0c414bd7fae244
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/12/2022
ms.locfileid: "61765891"
---
# <a name="event-element"></a>Élément Event

Définit un gestionnaire d’événements dans un complément.

> [!NOTE]
> Pour plus d’informations sur la prise en charge et l’utilisation, voir La fonctionnalité [d’envoi pour Outlook des applications.](../../outlook/outlook-on-send-addins.md)

**Type de complément :** messagerie

**Valide uniquement dans ces schémas VersionOverrides**:

- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Type](#type-attribute)  |  Oui  | Indique l’événement à gérer. |
|  [FunctionExecution](#functionexecution-attribute)  |  Oui  | Indique le style d’exécution du gestionnaire d’événements, asynchrone ou synchrone. Actuellement, seuls les gestionnaires d’événement synchrones sont pris en charge. |
|  [FunctionName](#functionname-attribute)  |  Oui  | Indique le nom de la fonction du gestionnaire d’événements. |

### <a name="type-attribute"></a>Attribut Type

Obligatoire. Indique l’événement qui appelle le gestionnaire d’événements. Les valeurs possibles pour cet attribut sont répertoriées dans le tableau suivant.

|  Type d’événement  |  Description  |
|:-----|:-----|
|  `ItemSend`  |  Le gestionnaire d’événements est appelé quand l’utilisateur envoie un message ou une convocation.  |

### <a name="functionexecution-attribute"></a>Attribut FunctionExecution

Obligatoire. DOIT être défini sur `synchronous`.

### <a name="functionname-attribute"></a>Attribut FunctionName

Obligatoire. Indique le nom de la fonction du gestionnaire d’événements. Cette valeur doit correspondre au nom d’une fonction dans le [fichier de fonction](functionfile.md) du complément.

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
