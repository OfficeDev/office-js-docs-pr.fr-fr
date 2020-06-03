---
title: Runtimes dans le fichier manifeste
description: L’élément runtimes spécifie le runtime de votre complément.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 95549d88df24a7d7c54cf27c92c15693491bdf29
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/03/2020
ms.locfileid: "44520338"
---
# <a name="runtimes-element"></a>Élément runtimes

Spécifie le runtime de votre complément. Enfant de l' [`<Host>`](host.md) élément.

> [!NOTE]
> Lors de l’exécution dans Office sur Windows, votre complément utilise le navigateur Internet Explorer 11.

Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime. Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

Dans Outlook, cet élément active l’activation de complément basée sur les événements. Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).

**Type de complément :** Volet Office, messagerie

> [!IMPORTANT]
> **Excel**: le runtime partagé est actuellement disponible uniquement dans Excel sur Windows.
>
> **Outlook**: la fonctionnalité d’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web. Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans

[Hôte](host.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Oui |  Le runtime de votre complément. |

## <a name="see-also"></a>Voir aussi

- [Runtime](runtime.md)
