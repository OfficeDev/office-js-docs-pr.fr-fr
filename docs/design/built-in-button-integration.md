---
title: Intégrer des boutons de Office intégrés dans des onglets et des groupes de contrôles personnalisés
description: Découvrez comment inclure des boutons de Office intégrés dans vos groupes de commandes et onglets personnalisés sur Office ruban.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 91e64e3939ea83c6468b1f8b35ac189ad7d3d373
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496725"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Intégrer des boutons de Office intégrés dans des onglets et des groupes de contrôles personnalisés

Vous pouvez insérer des boutons Office intégrés dans vos groupes de contrôles personnalisés sur le ruban Office à l’aide de la marque dans le manifeste du module. (Vous ne pouvez pas insérer vos commandes de Office personnalisées.) Vous pouvez également insérer des groupes de contrôles Office intégrés dans vos onglets de ruban personnalisés.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec l’article [Concepts de base pour les commandes de add-in](add-in-commands.md). Si vous ne l’avez pas fait récemment, veuillez l’examiner.

> [!IMPORTANT]
>
> - La fonctionnalité et le markup du add-in décrits dans cet article sont disponibles *uniquement dans PowerPoint sur le web*.
> - Le markup décrit dans cet article fonctionne uniquement sur les plateformes qui supportent l’ensemble de conditions **requises AddinCommands 1.3**. Consultez la section [Comportement sur les plateformes non pris en cas de problème](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Insérer un groupe de contrôles intégré dans un onglet personnalisé

Pour insérer un groupe de contrôles Office dans un onglet, ajoutez un élément [OfficeGroup](/javascript/api/manifest/customtab#officegroup) en tant qu’élément enfant dans l’élément **CustomTab** parent. L’attribut `id` de **l’élément OfficeGroup** est définie sur l’ID du groupe intégré. Voir [Rechercher les ID des contrôles et des groupes de contrôles](#find-the-ids-of-controls-and-control-groups).

L’exemple de marques de Office ajoute le groupe de contrôles Paragraph à un onglet personnalisé et le positionnait pour qu’il apparaisse juste après un groupe personnalisé.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Insérer un contrôle intégré dans un groupe personnalisé

Pour insérer un contrôle Office dans un groupe personnalisé, ajoutez un élément [OfficeControl](/javascript/api/manifest/group#officecontrol) en tant qu’élément enfant dans l’élément **Group** parent. L’attribut `id` de **l’élément OfficeControl** est définie sur l’ID du contrôle intégré. Voir [Rechercher les ID des contrôles et des groupes de contrôles](#find-the-ids-of-controls-and-control-groups).

L’exemple de marques de Office suivant ajoute le contrôle Superscript à un groupe personnalisé et le place pour qu’il apparaisse juste après un bouton personnalisé.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button1">
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
> Les utilisateurs peuvent personnaliser le ruban dans l Office’application. Toutes les personnalisations utilisateur remplaceront vos paramètres de manifeste. Par exemple, un utilisateur peut supprimer un bouton de n’importe quel groupe et supprimer n’importe quel groupe d’un onglet.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Rechercher les ID des contrôles et des groupes de contrôles

Les ID des contrôles et des groupes de contrôles pris en charge se font dans les fichiers du [Office de contrôle.](https://github.com/OfficeDev/office-control-ids) Suivez les instructions du fichier ReadMe de ce dépôt.

## <a name="behavior-on-unsupported-platforms"></a>Comportement sur les plateformes non pris en place

Si votre add-in est installé sur une plateforme qui ne prend pas en charge l’ensemble de conditions [requises AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), le markup décrit dans cet article est ignoré et les contrôles/groupes Office intégrés n’apparaissent pas dans vos groupes/onglets personnalisés. Pour empêcher l’installation de votre add-in sur des plateformes qui ne le supportent pas, ajoutez une référence à l’ensemble de conditions requises dans **la section Conditions** requises du manifeste. Pour obtenir des instructions, voir [Spécifier Office versions et plateformes peuvent héberger votre module.](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) Vous pouvez également concevoir votre add-in pour une expérience lorsque **AddinCommands 1.3** n’est pas pris en charge, comme décrit dans La conception pour [d’autres expériences](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Par exemple, si votre add-in contient des instructions qui supposent que les boutons intégrés se trouveront dans vos groupes personnalisés, vous pouvez concevoir une version qui suppose que les boutons intégrés se trouveront uniquement à leurs endroits habituels.
