---
title: Élément Event dans le fichier manifeste
description: Définit un gestionnaire d’événements dans un complément.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 486236f2c2dc19f835e06bad027b4fca33809fb257ba6f6d455add66ab5b5ce0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093295"
---
# <a name="event-element"></a>Élément Event

Définit un gestionnaire d’événements dans un complément.

> [!NOTE]
> Pour plus d’informations sur la prise en charge et l’utilisation, voir La fonctionnalité [d’envoi pour Outlook des applications.](../../outlook/outlook-on-send-addins.md)

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
