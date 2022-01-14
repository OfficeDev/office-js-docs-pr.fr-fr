---
title: Élémznr VersionOverrides dans le fichier manifest
description: Documentation de référence de l’élément VersionOverrides Office fichiers manifeste (XML) des modules.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 657bdebbc88993badd9d0e60946239edd55d5533
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042146"
---
# <a name="versionoverrides-element"></a>Élément VersionOverrides

Cet élément contient des informations sur les fonctionnalités qui ne sont pas pris en charge dans le manifeste de base. Son markup enfant peut remplacer une partie du markup dans le manifeste de base (ou dans un **parent VersionOverrides**). **VersionOverrides est** un élément enfant de l’élément [OfficeApp](officeapp.md) racine dans le manifeste ou d’un **élément VersionOverrides** parent. Cet élément est pris en charge dans les versions 1.1 et ultérieures du schéma de manifeste, mais il est défini dans des schémas VersionOverrides distincts.

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Oui  |  Espace de noms de schéma VersionOverrides. Les valeurs autorisées varient en fonction de la valeur xsi:type de cet élément et de la valeur `<VersionOverrides>` **xsi:type** de l’élément  `<OfficeApp>` parent. Voir [les valeurs d’espace de noms](#namespace-values) ci-dessous.|
|  **xsi:type**  |  Oui  | Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valeurs des espaces de noms

La liste suivante répertorie la valeur requise de l’attribut **xmlns** en fonction de la valeur **xsi:type** de l’élément `<OfficeApp>` racine.

- **TaskPaneApp prend** en charge uniquement la version 1.0 de VersionOverrides, et les **xmlns** doivent être `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** prend en charge uniquement la version 1.0 de VersionOverrides, et les **xmlns** doivent être `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** prend en charge les versions 1.0 et 1.1 de VersionOverrides, de sorte que la valeur de **xmlns** varie en fonction de la valeur `<VersionOverrides>` **xsi:type** de cet élément :
  - Lorsque **xsi:type** est `VersionOverridesV1_0` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides` .
  - Lorsque **xsi:type** est `VersionOverridesV1_1` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> Actuellement, Outlook 2016 ou version ultérieure prend en charge le schéma VersionOverrides v1.1 et le `VersionOverridesV1_1` type.

## <a name="variant-schemas"></a>Schémas de variantes

Il existe un schéma différent pour chacune des **valeurs xmlns** possibles, de sorte que chacune possède une page de référence distincte.

- [VersionOverrides 1.0 TaskPane](versionoverrides-1-0-taskpane.md)
- [Contenu VersionOverrides 1.0](versionoverrides-1-0-content.md)
- [VersionOverrides 1.0 Mail](versionoverrides-1-0-mail.md)
- [VersionOverrides 1.1 Mail](versionoverrides-1-1-mail.md)
