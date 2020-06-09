---
title: Élément Event dans le fichier manifeste
description: Définit un gestionnaire d’événements dans un complément.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 3d8e94c10bed214dd976b3048e11328f10f99325
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611546"
---
# <a name="event-element"></a>Élément Event

Définit un gestionnaire d’événements dans un complément.

> [!NOTE]
> Pour plus d’informations sur la prise en charge et l’utilisation, consultez la rubrique relative à la [fonctionnalité d’envoi pour les compléments Outlook](../../outlook/outlook-on-send-addins.md).

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
