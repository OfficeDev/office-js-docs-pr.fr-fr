---
title: Élément Enabled dans le fichier manifeste
description: Découvrez comment spécifier qu’une commande de add-in est désactivée au lancement du module.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: fc635e91b005eb51c70e8517058fc03fa4f26c6c
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511266"
---
# <a name="enabled-element"></a>Élément Enabled

Spécifie si un contrôle [Bouton ou](control-button.md) [Menu](control-menu.md) est activé au lancement du module. **L’élément Enabled** est un élément enfant de [Control](control.md). S’il est omis, la valeur par défaut est `true`.

**Type de complément :** volet Office

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Cet élément n’est valide que dans Excel, PowerPoint et Word, c’est-à-dire lorsque `Name` l’attribut de l’élément [Host](host.md) est « Workbook », « Presentation » ou « Document ».

Le contrôle parent peut également être activé et désactivé par programme. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

## <a name="example"></a>Exemple

```xml
<Enabled>false</Enabled>
```
