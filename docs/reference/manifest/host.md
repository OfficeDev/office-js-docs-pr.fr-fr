---
title: Élément Host dans le fichier manifeste
description: Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: ea0f5c8bc07c72c0c888fb56b40d98c6030c2ebc
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340686"
---
# <a name="host-element"></a>Élément Host

Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.

> [!IMPORTANT]
> La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node). Toutefois, la fonctionnalité est identique.  

## <a name="basic-manifest"></a>Manifeste de base

Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.

### <a name="attributes"></a>Attributs

| Attribut     | Type   | Requis | Description                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | obligatoire | Nom du type d’application Office client. |

### <a name="name"></a>Nom

Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

### <a name="example"></a>Exemple

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>Nœud VersionOverrides

Lorsqu’il est défini dans [VersionOverrides](versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`. 

Cet élément remplace **l’élément Hosts** dans le manifeste de base.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

### <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Spécifie l’Office application dans laquelle ces paramètres s’appliquent.|

### <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Oui   |  Définit les paramètres pour le facteur de forme pour bureau. |
|  [MobileFormFactor](mobileformfactor.md)    |  Non   |  Définit les paramètres pour le facteur de forme pour environnement mobile. **Remarque :** Cet élément est uniquement pris en charge dans Outlook sur iOS et Android. |
|  [AllFormFactors](allformfactors.md)    |  Non   |  Définit les paramètres de tous les facteurs de forme. Utilisé uniquement par des fonctions personnalisées dans Excel. |

### <a name="xsitype"></a>xsi:type

Contrôle l Office application (Word, Excel, PowerPoint, Outlook, OneNote) dans laquelle les paramètres contenus s’appliquent. La valeur doit être l’une des suivantes :

- `Document` (Word)
- `MailHost` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Exemple d’hôte

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
