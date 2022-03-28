---
title: Runtime dans le fichier manifeste
description: L’élément Runtime configure votre add-in pour utiliser un runtime JavaScript partagé pour ses différents composants, par exemple, ruban, volet Des tâches, fonctions personnalisées.
ms.date: 03/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38920dc43349be8da629785167d03252578f2a42
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404673"
---
# <a name="runtime-element"></a>Élément Runtime

Configure votre add-in pour utiliser un runtime JavaScript partagé afin que différents composants s’exécutent tous dans le même runtime. Enfant de l’élément [`<Runtimes>`](runtimes.md) .

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

 - Volet De tâches 1.0
 - Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (uniquement lorsqu’il est utilisé dans un add-in de volet de tâches.)

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>Syntaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contenu dans

- [Services d’exécution](runtimes.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
| [Override](override.md) | Non | **Outlook** : spécifie l’emplacement URL du fichier JavaScript dont Outlook Desktop a besoin pour les handleurs de [point d’extension LaunchEvent](../../reference/manifest/extensionpoint.md#launchevent). **Important** : Pour le moment, vous ne pouvez définir qu’un `<Override>` seul élément et il doit être de type `javascript`.|

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **resid**  |  Oui  | Spécifie l’emplacement URL de la page HTML de votre application. Il `resid` ne peut pas y avoir plus de 32 caractères `id` et doit correspondre à un attribut d’un `Url` élément dans l’élément `Resources` . |
|  [lifetime](#lifetime-attribute)  |  Non  | La valeur par défaut `lifetime` est `short` et n’a pas besoin d’être spécifiée. Outlook’activation basée sur des événements utilisent uniquement la `short` valeur. Si vous souhaitez utiliser un runtime partagé dans un Excel, définissez explicitement la valeur sur `long`. |

### <a name="lifetime-attribute"></a>attribut de durée de vie

Facultatif. Représente la durée d’exécuter le module.

**Valeurs disponibles**

`short`: Valeur par défaut. Utilisé uniquement pour Outlook’activation basée sur des événements. Une fois le add-in activé, il s’exécute pendant une durée maximale, comme spécifié par la plateforme. Actuellement, cela fait environ 5 minutes. Il s’agit de la seule valeur prise en charge par Outlook.

`long`: utilisé uniquement lors de la configuration [d’un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md). Le add-in peut démarrer sur le document ouvert et s’exécuter indéfiniment. Par exemple, le code du volet Des tâches continue d’être en cours d’exécution même lorsque l’utilisateur ferme le volet Des tâches. Il s’agit de la seule valeur prise en charge par le runtime partagé.

## <a name="see-also"></a>Voir aussi

- [Services d’exécution](runtimes.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](../../outlook/autolaunch.md)
