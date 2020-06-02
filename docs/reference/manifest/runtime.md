---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre complément de sorte qu’il utilise un Runtime JavaScript partagé pour ses différents composants, par exemple le ruban, le volet Office, les fonctions personnalisées.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: a463b72f22b41f74e2fe98acca467762bb00cf39
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471337"
---
# <a name="runtime-element-preview"></a>Élément Runtime (aperçu)

Configure votre complément pour qu’il utilise un Runtime JavaScript partagé afin que les différents composants s’exécutent tous dans le même Runtime. Enfant de l' [`<Runtimes>`](runtimes.md) élément.

Dans Excel, cet élément active le ruban, le volet des tâches et les fonctions personnalisées pour utiliser le même Runtime. Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

Dans Outlook, cet élément active l’activation de complément basée sur les événements. Pour plus d’informations, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../outlook/autolaunch.md).

**Type de complément :** Volet Office, messagerie

> [!IMPORTANT]
> **Excel**: le runtime partagé est actuellement disponible uniquement dans Excel sur Windows.
>
> **Outlook**: l’activation basée sur un événement est actuellement [en](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) préversion et disponible uniquement dans Outlook sur le Web. Pour plus d’informations, voir [comment afficher un aperçu de la fonctionnalité activation basée sur les événements](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
|  **resid**  |  Oui  | Spécifie l’URL de la page HTML de votre complément. L' `resid` doit correspondre à un `id` attribut d’un `Url` élément dans l' `Resources` élément. |
|  **vie**  |  Non  | La valeur par défaut de `lifetime` est `short` et n’a pas besoin d’être spécifiée. Les compléments Outlook utilisent uniquement la `short` valeur. Si vous souhaitez utiliser un runtime partagé dans un complément Excel, définissez explicitement la valeur sur `long` . |

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
