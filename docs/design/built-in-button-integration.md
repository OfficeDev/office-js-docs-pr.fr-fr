---
title: Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôle personnalisés
description: Découvrez comment inclure des boutons Office intégrés dans vos groupes de commandes et onglets personnalisés sur le ruban Office.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc706fcd0b049647847a73f7c40144dba9df0e2
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659786"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôle personnalisés

Vous pouvez insérer des boutons Office intégrés dans vos groupes de contrôles personnalisés sur le ruban Office à l’aide du balisage dans le manifeste du complément. (Vous ne pouvez pas insérer vos commandes de complément personnalisées dans un groupe Office intégré.) Vous pouvez également insérer des groupes de contrôles Office intégrés entiers dans vos onglets de ruban personnalisés.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les concepts de base de l’article [pour les commandes de complément](add-in-commands.md). Veuillez la consulter si vous ne l’avez pas fait récemment.

> [!IMPORTANT]
>
> - La fonctionnalité de complément et le balisage décrits dans cet article *sont disponibles uniquement dans PowerPoint sur le web*.
> - Le balisage décrit dans cet article fonctionne uniquement sur les plateformes qui prennent en charge l’ensemble de conditions requises **AddinCommands 1.3**. Consultez la section suivante [Comportement sur les plateformes non prises en charge](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Insérer un groupe de contrôle intégré dans un onglet personnalisé

Pour insérer un groupe de contrôles Office intégré dans un onglet, ajoutez un élément [OfficeGroup](/javascript/api/manifest/customtab#officegroup) en tant qu’élément enfant dans l’élément parent **\<CustomTab\>** . L’attribut `id` de l’élément **\<OfficeGroup\>** est défini sur l’ID du groupe intégré. Consultez [Rechercher les ID des contrôles et des groupes de contrôles](#find-the-ids-of-controls-and-control-groups).

L’exemple de balisage suivant ajoute le groupe de contrôle Paragraphes Office à un onglet personnalisé et le positionne pour qu’il apparaisse juste après un groupe personnalisé.

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

Pour insérer un contrôle Office intégré dans un groupe personnalisé, ajoutez un élément [OfficeControl](/javascript/api/manifest/group#officecontrol) en tant qu’élément enfant dans l’élément parent **\<Group\>** . L’attribut `id` de l’élément **\<OfficeControl\>** est défini sur l’ID du contrôle intégré. Consultez [Rechercher les ID des contrôles et des groupes de contrôles](#find-the-ids-of-controls-and-control-groups).

L’exemple de balisage suivant ajoute le contrôle Office Superscript à un groupe personnalisé et le positionne pour qu’il apparaisse juste après un bouton personnalisé.

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
> Les utilisateurs peuvent personnaliser le ruban dans l’application Office. Toutes les personnalisations utilisateur remplacent vos paramètres de manifeste. Par exemple, un utilisateur peut supprimer un bouton de n’importe quel groupe et supprimer n’importe quel groupe d’un onglet.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Rechercher les ID des contrôles et des groupes de contrôles

Les ID des contrôles pris en charge et des groupes de contrôles se trouvent dans les fichiers des [ID de contrôle Office](https://github.com/OfficeDev/office-control-ids) du référentiel. Suivez les instructions du fichier ReadMe de ce dépôt.

## <a name="behavior-on-unsupported-platforms"></a>Comportement sur les plateformes non prises en charge

Si votre complément est installé sur une plateforme qui ne prend pas en charge [l’ensemble de conditions requises AddinCommands 1.3, le balisage](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) décrit dans cet article est ignoré et les contrôles/groupes Office intégrés n’apparaissent pas dans vos groupes/onglets personnalisés. Pour empêcher l’installation de votre complément sur des plateformes qui ne prennent pas en charge le balisage, ajoutez une référence à l’ensemble de conditions requises dans la **\<Requirements\>** section du manifeste. Pour obtenir des instructions, consultez [Spécifier les versions et plateformes Office qui peuvent héberger votre complément](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Vous pouvez également concevoir votre complément pour avoir une expérience quand **AddinCommands 1.3** n’est pas pris en charge, comme décrit dans [la conception pour d’autres expériences](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Par exemple, si votre complément contient des instructions qui supposent que les boutons intégrés se trouvent dans vos groupes personnalisés, vous pouvez concevoir une version qui suppose que les boutons intégrés se trouvent uniquement à leur emplacement habituel.
