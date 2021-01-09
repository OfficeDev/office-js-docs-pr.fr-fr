---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet Des tâches, fonctions personnalisées.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789183"
---
# <a name="runtime-element-preview"></a>Élément Runtime (aperçu)

Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime. Enfant de [`<Runtimes>`](runtimes.md) l’élément.

Dans Excel, cet élément permet au ruban, au volet Des tâches et aux fonctions personnalisées d’utiliser le même runtime. Pour plus d’informations, voir Configurer votre add-in Excel pour utiliser [un runtime JavaScript partagé.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

Dans Outlook, cet élément active l’activation des compléments basés sur des événements. Pour plus d’informations, voir Configurer votre complément [Outlook pour l’activation basée sur des événements.](../../outlook/autolaunch.md)

**Type de add-in :** Volet De tâches, Courrier

> [!IMPORTANT]
> **Outlook**: l’activation basée sur des événements est actuellement [en prévisualisation](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) et disponible uniquement dans Outlook sur le web. Pour plus d’informations, [voir Comment afficher un aperçu de la fonctionnalité d’activation basée sur des événements.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans

- [Services d’exécution](runtimes.md)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid**  |  Oui  | Spécifie l’emplacement URL de la page HTML de votre application. Il ne peut pas y avoir plus de 32 caractères et doit correspondre à un `resid` `id` attribut `Url` d’un élément dans `Resources` l’élément. |
|  **lifetime**  |  Non  | La valeur par `lifetime` défaut est et n’a pas besoin `short` d’être spécifiée. Les add-ins Outlook utilisent uniquement la `short` valeur. Si vous souhaitez utiliser un runtime partagé dans un add-in Excel, définissez explicitement la valeur sur `long` . |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
