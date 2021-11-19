---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de add-in est désactivée au lancement du module.
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4c0107daaf73aee6ba116553a8d01250e9c7d981
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081434"
---
# <a name="enabled-element"></a>Élément Enabled

Spécifie si un [contrôle Bouton](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls) est activé au lancement du module. **L’élément Enabled** est un élément enfant de [Control](control.md). S’il est omis, la valeur par défaut est `true` .

**Type de complément :** volet Office

**Valide uniquement dans ces schémas VersionOverrides**:

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Cet élément n’est valide que dans Excel, c’est-à-dire lorsque l’attribut de l’élément `Name` [Host](host.md) est « Workbook ».

Le contrôle parent peut également être activé et désactivé par programme. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
