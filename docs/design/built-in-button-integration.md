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
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a><span data-ttu-id="00123-103">Intégrer des boutons Office intégrés dans des onglets et des groupes de contrôles personnalisés (aperçu)</span><span class="sxs-lookup"><span data-stu-id="00123-103">Integrate built-in Office buttons into custom control groups and tabs (preview)</span></span>

<span data-ttu-id="00123-104">Vous pouvez insérer des boutons Office intégrés dans vos groupes de contrôles personnalisés sur le ruban Office à l’aide de balises dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="00123-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="00123-105">(Vous ne pouvez pas insérer vos commandes de complément personnalisées dans un groupe Office prédéfini.) Vous pouvez également insérer des groupes de contrôles Office prédéfinis dans vos onglets de ruban personnalisés.</span><span class="sxs-lookup"><span data-stu-id="00123-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="00123-106">Cet article suppose que vous connaissez bien l’article [concepts de base pour les commandes de complément](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="00123-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="00123-107">Vérifiez-le si vous ne l’avez pas encore fait.</span><span class="sxs-lookup"><span data-stu-id="00123-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="00123-108">La fonctionnalité de complément et le balisage décrits dans cet article sont dans l’aperçu et sont *disponibles uniquement dans PowerPoint sur le Web*.</span><span class="sxs-lookup"><span data-stu-id="00123-108">The add-in feature and markup described in this article is in preview and is *only available in PowerPoint on the web*.</span></span> <span data-ttu-id="00123-109">Nous vous recommandons d’essayer le balisage uniquement dans les environnements de test et de développement.</span><span class="sxs-lookup"><span data-stu-id="00123-109">We recommend that you try out the markup in test and development environments only.</span></span> <span data-ttu-id="00123-110">N’utilisez pas les marques de révision dans un environnement de production ou dans des documents professionnels.</span><span class="sxs-lookup"><span data-stu-id="00123-110">Do not use preview markup in a production environment or within business-critical documents.</span></span>
> - <span data-ttu-id="00123-111">Le balisage décrit dans cet article fonctionne uniquement sur les plateformes qui prennent en charge l’ensemble de conditions requises **AddinCommands 1,3**.</span><span class="sxs-lookup"><span data-stu-id="00123-111">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="00123-112">Voir le comportement de la section plus loin [sur les plateformes non prises en charge](#behavior-on-unsupported-platforms).</span><span class="sxs-lookup"><span data-stu-id="00123-112">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="00123-113">Insérer un groupe de contrôles prédéfini dans un onglet personnalisé</span><span class="sxs-lookup"><span data-stu-id="00123-113">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="00123-114">Pour insérer un groupe de contrôles Office prédéfini dans un onglet, ajoutez un élément [OfficeGroup](../reference/manifest/customtab.md#officegroup) en tant qu’élément enfant dans l' `<CustomTab>` élément parent.</span><span class="sxs-lookup"><span data-stu-id="00123-114">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="00123-115">L' `id` attribut de l’élément de l' `<OfficeGroup>` élément est défini sur l’ID du groupe prédéfini.</span><span class="sxs-lookup"><span data-stu-id="00123-115">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="00123-116">Voir [Rechercher les ID des contrôles et des groupes de](#find-the-ids-of-controls-and-control-groups)contrôles.</span><span class="sxs-lookup"><span data-stu-id="00123-116">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="00123-117">L’exemple de balisage suivant montre comment ajouter le groupe de contrôles paragraphe Office à un onglet personnalisé et le positionner de sorte qu’il apparaisse immédiatement après un groupe personnalisé.</span><span class="sxs-lookup"><span data-stu-id="00123-117">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="00123-118">Insérer un contrôle intégré dans un groupe personnalisé</span><span class="sxs-lookup"><span data-stu-id="00123-118">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="00123-119">Pour insérer un contrôle Office intégré dans un groupe personnalisé, ajoutez un élément [OfficeControl](../reference/manifest/group.md#officecontrol) en tant qu’élément enfant dans l’élément parent `<Group>` .</span><span class="sxs-lookup"><span data-stu-id="00123-119">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="00123-120">L' `id` attribut de l' `<OfficeControl>` élément est défini sur l’ID du contrôle intégré.</span><span class="sxs-lookup"><span data-stu-id="00123-120">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="00123-121">Voir [Rechercher les ID des contrôles et des groupes de](#find-the-ids-of-controls-and-control-groups)contrôles.</span><span class="sxs-lookup"><span data-stu-id="00123-121">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="00123-122">L’exemple de balisage suivant ajoute le contrôle Office Superscript à un groupe personnalisé et l’affiche juste après un bouton personnalisé.</span><span class="sxs-lookup"><span data-stu-id="00123-122">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

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
> <span data-ttu-id="00123-123">Les utilisateurs peuvent personnaliser le ruban dans l’application Office.</span><span class="sxs-lookup"><span data-stu-id="00123-123">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="00123-124">Toutes les personnalisations utilisateur remplacent les paramètres de votre manifeste.</span><span class="sxs-lookup"><span data-stu-id="00123-124">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="00123-125">Par exemple, un utilisateur peut supprimer un bouton d’un groupe et supprimer un groupe d’un onglet.</span><span class="sxs-lookup"><span data-stu-id="00123-125">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="00123-126">Rechercher les ID des contrôles et des groupes de contrôles</span><span class="sxs-lookup"><span data-stu-id="00123-126">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="00123-127">Les ID des contrôles et des groupes de contrôles pris en charge se trouvent dans des fichiers dans les [ID de contrôle Office](https://github.com/OfficeDev/office-control-ids)référentiel.</span><span class="sxs-lookup"><span data-stu-id="00123-127">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="00123-128">Suivez les instructions du fichier Lisez-moi de cette référentiel.</span><span class="sxs-lookup"><span data-stu-id="00123-128">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="00123-129">Comportement sur les plateformes non prises en charge</span><span class="sxs-lookup"><span data-stu-id="00123-129">Behavior on unsupported platforms</span></span>

<span data-ttu-id="00123-130">Si votre complément est installé sur une plateforme qui ne prend pas en charge l' [ensemble de conditions de AddinCommands 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), le balisage décrit dans cet article est ignoré et les contrôles/groupes Office prédéfinis n’apparaîtront pas dans vos groupes/onglets personnalisés.</span><span class="sxs-lookup"><span data-stu-id="00123-130">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="00123-131">Pour empêcher l’installation de votre complément sur des plateformes qui ne prennent pas en charge le balisage, ajoutez une référence à l’ensemble de conditions requises dans la `<Requirements>` section du manifeste.</span><span class="sxs-lookup"><span data-stu-id="00123-131">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="00123-132">Pour obtenir des instructions, voir [définir l’élément Requirements dans le manifeste](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="00123-132">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="00123-133">Vous pouvez également concevoir votre complément de manière à ce qu’il ait une expérience secondaire lorsque **AddinCommands 1,3** n’est pas pris en charge, comme décrit dans [la rubrique use Runtime Checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="00123-133">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="00123-134">Par exemple, si votre complément contient des instructions qui supposent que les boutons intégrés se trouvent dans vos groupes personnalisés, vous pouvez avoir une autre version qui suppose que les boutons intégrés ne se trouvent qu’à leurs emplacements habituels.</span><span class="sxs-lookup"><span data-stu-id="00123-134">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
