---
title: Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôles personnalisés
description: Découvrez comment inclure des boutons Office prédéfinis dans vos groupes de commandes personnalisés et des onglets dans le ruban Office.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: e04107893b3c0dd453c84d38fdd5623e308b70e3
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088168"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a>Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôles personnalisés (aperçu)

Vous pouvez insérer des boutons Office intégrés dans vos groupes de contrôles personnalisés sur le ruban Office à l’aide de balises dans le manifeste du complément. (Vous ne pouvez pas insérer vos commandes de complément personnalisées dans un groupe Office prédéfini.) Vous pouvez également insérer des groupes de contrôles Office prédéfinis dans vos onglets de ruban personnalisés.

> [!NOTE]
> Cet article suppose que vous connaissez bien l’article [concepts de base pour les commandes de complément](add-in-commands.md). Vérifiez-le si vous ne l’avez pas encore fait.

> [!IMPORTANT]
>
> - La fonctionnalité de complément et le balisage décrits dans cet article sont dans l’aperçu et sont *disponibles uniquement dans PowerPoint sur le Web*. Nous vous recommandons d’essayer le balisage uniquement dans les environnements de test et de développement. N’utilisez pas les marques de révision dans un environnement de production ou dans des documents professionnels.
> - Le balisage décrit dans cet article fonctionne uniquement sur les plateformes qui prennent en charge l’ensemble de conditions requises **AddinCommands 1,3**. Voir le comportement de la section plus loin [sur les plateformes non prises en charge](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Insérer un groupe de contrôles prédéfini dans un onglet personnalisé

Pour insérer un groupe de contrôles Office prédéfini dans un onglet, ajoutez un élément [OfficeGroup](../reference/manifest/customtab.md#officegroup) en tant qu’élément enfant dans l' `<CustomTab>` élément parent. L' `id` attribut de l’élément de l' `<OfficeGroup>` élément est défini sur l’ID du groupe prédéfini. Voir [Rechercher les ID des contrôles et des groupes de](#find-the-ids-of-controls-and-control-groups)contrôles.

L’exemple de balisage suivant montre comment ajouter le groupe de contrôles paragraphe Office à un onglet personnalisé et le positionner de sorte qu’il apparaisse immédiatement après un groupe personnalisé.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Insérer un contrôle intégré dans un groupe personnalisé

Pour insérer un contrôle Office intégré dans un groupe personnalisé, ajoutez un élément [OfficeControl](../reference/manifest/group.md#officecontrol) en tant qu’élément enfant dans l’élément parent `<Group>` . L' `id` attribut de l' `<OfficeControl>` élément est défini sur l’ID du contrôle intégré. Voir [Rechercher les ID des contrôles et des groupes de](#find-the-ids-of-controls-and-control-groups)contrôles.

L’exemple de balisage suivant ajoute le contrôle Office Superscript à un groupe personnalisé et l’affiche juste après un bouton personnalisé.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> Les utilisateurs peuvent personnaliser le ruban dans l’application Office. Toutes les personnalisations utilisateur remplacent les paramètres de votre manifeste. Par exemple, un utilisateur peut supprimer un bouton d’un groupe et supprimer un groupe d’un onglet.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Rechercher les ID des contrôles et des groupes de contrôles

Les ID des contrôles et des groupes de contrôles pris en charge se trouvent dans des fichiers dans les [ID de contrôle Office](https://github.com/OfficeDev/office-control-ids)référentiel. Suivez les instructions du fichier Lisez-moi de cette référentiel.

## <a name="behavior-on-unsupported-platforms"></a>Comportement sur les plateformes non prises en charge

Si votre complément est installé sur une plateforme qui ne prend pas en charge l' [ensemble de conditions de AddinCommands 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), le balisage décrit dans cet article est ignoré et les contrôles/groupes Office prédéfinis n’apparaîtront pas dans vos groupes/onglets personnalisés. Pour empêcher l’installation de votre complément sur des plateformes qui ne prennent pas en charge le balisage, ajoutez une référence à l’ensemble de conditions requises dans la `<Requirements>` section du manifeste. Pour obtenir des instructions, voir [définir l’élément Requirements dans le manifeste](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Vous pouvez également concevoir votre complément de manière à ce qu’il ait une expérience secondaire lorsque **AddinCommands 1,3** n’est pas pris en charge, comme décrit dans [la rubrique use Runtime Checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). Par exemple, si votre complément contient des instructions qui supposent que les boutons intégrés se trouvent dans vos groupes personnalisés, vous pouvez avoir une autre version qui suppose que les boutons intégrés ne se trouvent qu’à leurs emplacements habituels.
