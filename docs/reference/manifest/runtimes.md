---
title: Runtimes dans le fichier manifeste
description: L’élément Runtimes spécifie le runtime de votre add-in.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789162"
---
# <a name="runtimes-element"></a>Élément Runtimes

Spécifie le runtime de votre add-in. Enfant de [`<Host>`](host.md) l’élément.

> [!NOTE]
> Lorsque vous exécutez Office sur Windows, votre application utilise le navigateur Internet Explorer 11.

Dans Excel, cet élément permet au ruban, au volet Des tâches et aux fonctions personnalisées d’utiliser le même runtime. Pour plus d’informations, voir Configurer votre add-in Excel pour utiliser [un runtime JavaScript partagé.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

Dans Outlook, cet élément active l’activation des compléments basés sur des événements. Pour plus d’informations, voir Configurer votre complément [Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md)

**Type de add-in :** Volet De tâches, Courrier

> [!IMPORTANT]
> **Outlook**: la fonctionnalité d’activation basée sur des événements est actuellement en [prévisualisation](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) et disponible uniquement dans Outlook sur le web. Pour plus d’informations, [voir Comment afficher un aperçu de la fonctionnalité d’activation basée sur des événements.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans

[Host](host.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Oui |  Runtime de votre add-in. |

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
