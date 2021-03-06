---
title: Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôles personnalisés
description: Découvrez comment inclure des boutons Office intégrés dans vos groupes de commandes et onglets personnalisés sur le ruban Office.
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505254"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a><span data-ttu-id="b633b-103">Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôles personnalisés</span><span class="sxs-lookup"><span data-stu-id="b633b-103">Integrate built-in Office buttons into custom control groups and tabs</span></span>

<span data-ttu-id="b633b-104">Vous pouvez insérer des boutons Office intégrés dans vos groupes de contrôles personnalisés sur le ruban Office à l’aide de la marque dans le manifeste du module.</span><span class="sxs-lookup"><span data-stu-id="b633b-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="b633b-105">(Vous ne pouvez pas insérer vos commandes de add-in personnalisées dans un groupe Office intégré.) Vous pouvez également insérer des groupes de contrôles Office intégrés entiers dans vos onglets de ruban personnalisés.</span><span class="sxs-lookup"><span data-stu-id="b633b-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="b633b-106">Cet article suppose que vous êtes familiarisé avec l’article Concepts de base pour les [commandes de add-in.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="b633b-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="b633b-107">Si vous ne l’avez pas fait récemment, veuillez l’examiner.</span><span class="sxs-lookup"><span data-stu-id="b633b-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="b633b-108">La fonctionnalité de l’application et le markup décrits dans cet article sont disponibles uniquement *dans PowerPoint sur le web.*</span><span class="sxs-lookup"><span data-stu-id="b633b-108">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="b633b-109">Le markup décrit dans cet article fonctionne uniquement sur les plateformes qui supportent l’ensemble de conditions **requises AddinCommands 1.3**.</span><span class="sxs-lookup"><span data-stu-id="b633b-109">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="b633b-110">Consultez la section [Comportement sur les plateformes](#behavior-on-unsupported-platforms)non pris en place.</span><span class="sxs-lookup"><span data-stu-id="b633b-110">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="b633b-111">Insérer un groupe de contrôles intégré dans un onglet personnalisé</span><span class="sxs-lookup"><span data-stu-id="b633b-111">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="b633b-112">Pour insérer un groupe de contrôles Office intégré dans un onglet, ajoutez un élément [OfficeGroup](../reference/manifest/customtab.md#officegroup) en tant qu’élément enfant dans l’élément `<CustomTab>` parent.</span><span class="sxs-lookup"><span data-stu-id="b633b-112">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="b633b-113">`id`L’attribut de l’élément est définie sur l’ID `<OfficeGroup>` du groupe intégré.</span><span class="sxs-lookup"><span data-stu-id="b633b-113">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="b633b-114">Voir [Rechercher les ID des contrôles et des groupes de contrôles.](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="b633b-114">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="b633b-115">L’exemple de marques de commande suivant ajoute le groupe de contrôles Office Paragraph à un onglet personnalisé et le place pour qu’il apparaisse juste après un groupe personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b633b-115">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="b633b-116">Insérer un contrôle intégré dans un groupe personnalisé</span><span class="sxs-lookup"><span data-stu-id="b633b-116">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="b633b-117">Pour insérer un contrôle Office intégré dans un groupe personnalisé, ajoutez un élément [OfficeControl](../reference/manifest/group.md#officecontrol) en tant qu’élément enfant dans l’élément `<Group>` parent.</span><span class="sxs-lookup"><span data-stu-id="b633b-117">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="b633b-118">`id`L’attribut de `<OfficeControl>` l’élément est définie sur l’ID du contrôle intégré.</span><span class="sxs-lookup"><span data-stu-id="b633b-118">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="b633b-119">Voir [Rechercher les ID des contrôles et des groupes de contrôles.](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="b633b-119">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="b633b-120">L’exemple de marques de commande suivant ajoute le contrôle Exposant Office à un groupe personnalisé et le place pour qu’il apparaisse juste après un bouton personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b633b-120">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

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
> <span data-ttu-id="b633b-121">Les utilisateurs peuvent personnaliser le ruban dans l’application Office.</span><span class="sxs-lookup"><span data-stu-id="b633b-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="b633b-122">Toutes les personnalisations utilisateur remplaceront vos paramètres de manifeste.</span><span class="sxs-lookup"><span data-stu-id="b633b-122">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="b633b-123">Par exemple, un utilisateur peut supprimer un bouton de n’importe quel groupe et supprimer n’importe quel groupe d’un onglet.</span><span class="sxs-lookup"><span data-stu-id="b633b-123">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="b633b-124">Rechercher les ID des contrôles et des groupes de contrôles</span><span class="sxs-lookup"><span data-stu-id="b633b-124">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="b633b-125">Les ID des contrôles et des groupes de contrôles pris en charge se contiennent dans les fichiers des [ID](https://github.com/OfficeDev/office-control-ids)de contrôle Office de repo.</span><span class="sxs-lookup"><span data-stu-id="b633b-125">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="b633b-126">Suivez les instructions du fichier ReadMe de ce dépôt.</span><span class="sxs-lookup"><span data-stu-id="b633b-126">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="b633b-127">Comportement sur les plateformes non pris en place</span><span class="sxs-lookup"><span data-stu-id="b633b-127">Behavior on unsupported platforms</span></span>

<span data-ttu-id="b633b-128">Si votre add-in est installé sur une plateforme qui ne prend pas en charge l’ensemble de conditions [requises AddinCommands 1.3,](../reference/requirement-sets/add-in-commands-requirement-sets.md)le markup décrit dans cet article est ignoré et les contrôles/groupes Office intégrés n’apparaissent pas dans vos groupes/onglets personnalisés.</span><span class="sxs-lookup"><span data-stu-id="b633b-128">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="b633b-129">Pour empêcher l’installation de votre add-in sur des plateformes qui ne la prisent pas en charge, ajoutez une référence à l’ensemble de conditions requises dans la `<Requirements>` section du manifeste.</span><span class="sxs-lookup"><span data-stu-id="b633b-129">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="b633b-130">Pour obtenir des instructions, [voir Définir l’élément Requirements dans le manifeste.](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="b633b-130">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="b633b-131">Vous pouvez également concevoir votre add-in pour qu’il offre une expérience de substitution lorsque **AddinCommands 1.3** n’est pas pris en charge, comme décrit dans utiliser les vérifications à l’runtime dans votre [code JavaScript.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="b633b-131">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="b633b-132">Par exemple, si votre add-in contient des instructions qui supposent que les boutons intégrés se trouveront dans vos groupes personnalisés, vous pouvez avoir une autre version qui suppose que les boutons intégrés se trouveront uniquement à leurs endroits habituels.</span><span class="sxs-lookup"><span data-stu-id="b633b-132">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
