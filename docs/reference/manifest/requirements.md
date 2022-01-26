---
title: Élément Requirements dans le fichier manifest
description: L’élément Requirements spécifie l’ensemble minimal de conditions requises et les méthodes que votre Office doit activer par Office ou pour remplacer les paramètres de manifeste de base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85dcd08f3bfcffe34c4c479608f25ea0c2b6a134
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222282"
---
# <a name="requirements-element"></a>Élément Requirements

La signification de cet élément dépend de son utilisation dans le manifeste de [base](#in-the-base-manifest)] ou en tant qu’enfant d’un [élément **VersionOverrides**](#as-a-child-of-a-versionoverrides-element).

> [!TIP]
> Avant d’utiliser cet élément, familiarisez-vous avec [spécifier Office hôtes et les conditions requises de l’API](../../develop/specify-office-hosts-and-api-requirements.md)

## <a name="in-the-base-manifest"></a>Dans le manifeste de base

Lorsqu’il est utilisé dans le manifeste de base (c’est-à-dire, en tant qu’enfant direct [d’OfficeApp),](officeapp.md)l’élément **Requirements** spécifie l’ensemble minimal de conditions requises de l’API JavaScript Office [(ensembles](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) de conditions requises et/ou méthodes) que votre Office Add-in doit être activé par Office. Le add-in ne sera activé sur aucune combinaison de version et de plateforme Office (par exemple, Windows, Mac, web et iOS ou iPad) qui ne prend pas en charge les méthodes et ensembles de conditions requises spécifiés.

**Type de add-in :** Volet De tâches, Courrier

## <a name="as-a-child-of-a-versionoverrides-element"></a>Enfant d’un élément VersionOverrides

Lorsqu’il est utilisé en tant qu’enfant de [VersionOverrides,](versionoverrides.md)spécifie [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) l’ensemble minimal de conditions requises pour l’API JavaScript Office (ensembles de conditions requises et/ou méthodes) qui doivent être pris en charge par la version et la plateforme Office (telles que Windows, Mac, web et iOS ou iPad) afin que les paramètres de l’élément **VersionOverrides** remplacent les *paramètres* de manifeste de base  pour prendre effet.

Considérons un add-in qui spécifie l’exigence A dans le manifeste de base et spécifie la condition B à l’intérieur **de VersionOverrides**. 

- Si la plateforme et la version Office ne prend pas en charge A, le add-in n’est pas activé et Office n’active pas la section **VersionOverrides** du manifeste. 
- Si les deux versions A et B sont prises en charge, le add-in est activé et tous les marques de l’application **VersionOverrides** prennent effet. 
- Si A est pris en charge, mais pas B,  le module est activé et une partie du markup dans **VersionOverrides** prend effet. Plus précisément, les éléments enfants de **VersionOverrides** qui ne remplacent pas les éléments de manifeste de base prennent effet. Par exemple, un **élément WebApplicationInfo** ou **EquivalentAddins** prend effet. Toutefois, tous les éléments enfants de **VersionOverrides** qui remplacent un élément de manifeste de base, tels que **Hosts,** ne prennent pas effet. Au lieu de cela, Office utilise les valeurs du marques de manifeste de base qui auraient été autrement overridées. 

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

### <a name="remarks"></a>Remarques

**L’élément Requirements** n’a aucun rôle dans **versionOverrides** s’il ne spécifie aucune exigence supplémentaire qui n’est pas spécifiée dans un élément **Requirements** dans le manifeste de base. Si la version et la plateforme Office ne sont pas en charge dans le manifeste de base, le add-in n’est pas activé et l’élément **VersionOverrides** n’est pas l’élément. Pour cette raison, vous devez utiliser un élément **Requirements** dans **une VersionOverrides** uniquement lorsque ces deux conditions sont remplies :

- Votre add-in comporte des fonctionnalités supplémentaires qui sont implémentées avec la configuration dans **une VersionOverrides** (telles que les commandes de add-in) et qui nécessitent une méthode ou un ensemble de conditions requises qui n’est pas spécifié dans un élément **Requirements** dans le manifeste de base. 
- Votre add-in est utile et doit être activé (mais sans les fonctionnalités supplémentaires), même dans une combinaison de plateforme et de version Office qui ne prend pas en charge les conditions requises pour les fonctionnalités supplémentaires.

> [!TIP]
> Ne répétez pas **les éléments Requirement** à partir du manifeste de base à l’intérieur **d’une VersionOverrides**. Cela n’a aucun effet et peut être erronée quant à l’objectif de l’élément **Requirements** à l’intérieur d’une **VersionOverrides**.

> [!WARNING]
> Faites très attention avant d’utiliser un élément **Requirements** dans **une VersionOverrides,** car sur les combinaisons de plateforme et de version qui ne prend pas en charge la condition *requise,* aucune des commandes de module ne sera installée, même celles qui appellent des fonctionnalités qui *n’ont* pas besoin de cette exigence. Prenons l’exemple d’un add-in qui possède deux boutons de ruban personnalisés. L’un d’eux Office des API JavaScript disponibles dans l’ensemble de conditions **requises ExcelApi 1.4** (et version ultérieure). Les autres appellent des API qui sont uniquement disponibles dans **ExcelApi 1.9** (et ultérieur). Si vous avez ajouté une condition requise pour **ExcelApi 1.9** dans **VersionOverrides,** aucun des deux boutons n’apparaît sur le ruban.  Une meilleure stratégie dans ce scénario consisterait à utiliser la technique décrite dans les vérifications runtime pour la prise en charge des méthodes et des ensembles [de conditions requises.](../../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support) Le code appelé par le deuxième bouton utilise d’abord pour vérifier la prise en charge `isSetSupported` **d’ExcelApi 1.9**. S’il n’est pas pris en charge, le code envoie à l’utilisateur un message lui disant que cette fonctionnalité du module n’est pas disponible sur sa version de Office. 

> [!NOTE]
> Dans les modules de messagerie, il est possible d’imbrier **versionOverrides** 1.1 à l’intérieur d’une **version VersionOverrides** 1.0. Office utilisera toujours la version la plus élevée **VersionOverrides** prise en charge par la plateforme et Office version.

## <a name="syntax"></a>Syntaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md) 
 [VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>Peut contenir

|Élément|Contenu|Courrier|TaskPane|
|:-----|:-----|:-----|:-----|
|[Ensembles](sets.md)|x|x|x|
|[Méthodes](methods.md)|x||x|

## <a name="see-also"></a>Voir aussi

Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).
