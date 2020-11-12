---
title: Comment trouver l’ordre approprié d’éléments manifeste
description: Découvrez comment trouver l’ordre correct dans lequel placer les éléments enfants dans un élément parent.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 8798031410e6d71fd6d9f3f08f89a4c6f78f4692
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996381"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="62415-103">Comment trouver l’ordre approprié d’éléments manifeste</span><span class="sxs-lookup"><span data-stu-id="62415-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="62415-104">Les éléments XML dans le fichier manifeste d’un complément Office doivent être sous l’élément parent approprié *et* dans un ordre spécifique, par rapport à d’autres, sous le parent.</span><span class="sxs-lookup"><span data-stu-id="62415-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="62415-105">Le classement requis est spécifié dans les fichiers XSD dans le dossier [schémas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="62415-105">The required ordering is specified in the XSD files in the [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) folder.</span></span> <span data-ttu-id="62415-106">Les fichiers XSD sont classés dans des sous-dossiers pour volet de tâches, contenu et compléments de courrier.</span><span class="sxs-lookup"><span data-stu-id="62415-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="62415-107">Par exemple, dans l’`<OfficeApp>`élément, le `<Id>`,`<Version>` ,`<ProviderName>` doit apparaître dans cet ordre.</span><span class="sxs-lookup"><span data-stu-id="62415-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="62415-108">Si un élément `<AlternateId>` est ajouté, il doit être compris entre l’élément `<Id>` et `<Version>`.</span><span class="sxs-lookup"><span data-stu-id="62415-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="62415-109">Votre manifeste ne sera pas valide et votre complément ne sera pas chargé, si un élément n’est pas dans l’ordre.</span><span class="sxs-lookup"><span data-stu-id="62415-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="62415-110">Le [validateur d’Office-AddIn-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) utilise le même message d’erreur lorsqu’un élément est absent de l’ordre lorsqu’un élément est sous un parent incorrect.</span><span class="sxs-lookup"><span data-stu-id="62415-110">The [validator within office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="62415-111">L’erreur indique que l’élément enfant n’est pas un enfant valide de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="62415-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="62415-112">Si vous recevez un message d’erreur mais que la documentation de référence pour l’élément enfant indique qu’elle *est* valide pour le parent, alors le problème est probablement que l’enfant a été placé dans l’ordre incorrect.</span><span class="sxs-lookup"><span data-stu-id="62415-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="62415-113">Les sections suivantes présentent les éléments de manifeste dans l’ordre dans lequel ils doivent apparaître.</span><span class="sxs-lookup"><span data-stu-id="62415-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="62415-114">Il existe des différences selon que l' `type` attribut de l' `<OfficeApp>` élément est `TaskPaneApp` , `ContentApp` ou `MailApp` .</span><span class="sxs-lookup"><span data-stu-id="62415-114">There are differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="62415-115">Pour éviter que ces sections deviennent trop encombrantes, l’élément hautement complexe `<VersionOverrides>` est divisé en sections distinctes.</span><span class="sxs-lookup"><span data-stu-id="62415-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="62415-116">Tous les éléments affichés ne sont pas obligatoires.</span><span class="sxs-lookup"><span data-stu-id="62415-116">Not all of the elements shown are mandatory.</span></span> <span data-ttu-id="62415-117">Si la `minOccurs` valeur d’un élément est **0** dans le [schéma](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), l’élément est facultatif.</span><span class="sxs-lookup"><span data-stu-id="62415-117">If the `minOccurs` value for a element is **0** in the [schema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="62415-118">Classement des éléments de complément du volet Office de base</span><span class="sxs-lookup"><span data-stu-id="62415-118">Basic task pane add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="TaskPaneApp">
    <Id>
    <AlternateID>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
        <Sets>
            <Set>
        <Methods>
            <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <Permissions>
    <Dictionary>
        <TargetDialects>
        <QueryUri>
        <CitationText>
        <DictionaryName>
        <DictionaryHomePage>
    <VersionOverrides>*
    <ExtendedOverrides>
```

<span data-ttu-id="62415-119">\*Voir classement des éléments de [complément du volet Office dans VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) pour l’ordre des éléments enfants de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="62415-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="62415-120">Classement des éléments des compléments de messagerie de base</span><span class="sxs-lookup"><span data-stu-id="62415-120">Basic mail add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="MailApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <FormSettings>
        <Form>
        <DesktopSettings>
            <SourceLocation>
            <RequestedHeight>
        <TabletSettings>
            <SourceLocation>
            <RequestedHeight>
        <PhoneSettings>
            <SourceLocation>
    <Permissions>
    <Rule>
    <DisableEntityHighlighting>
    <VersionOverrides>*
```

<span data-ttu-id="62415-121">\*Consultez l’ordre des éléments de [compléments de messagerie dans VersionOverrides ver. 1,0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) et classement des éléments de [complément de messagerie dans VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) pour l’ordre des éléments enfants de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="62415-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="62415-122">Classement des éléments de complément de contenu de base</span><span class="sxs-lookup"><span data-stu-id="62415-122">Basic content add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="ContentApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl >
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <Methods>
        <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <RequestedWidth>
    <RequestedHeight>
    <Permissions>
    <AllowSnapshot>
    <VersionOverrides>*
```

<span data-ttu-id="62415-123">\*Consultez l’ordre des éléments de [complément de contenu dans VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) pour obtenir l’ordre des éléments enfants de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="62415-123">\*See [Content add-in element ordering within VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="62415-124">Classement des éléments de complément du volet Office dans VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="62415-124">Task pane add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <Runtimes>
                <Runtime>
            <AllFormFactors>
                <ExtensionPoint>
                    <Script>
                        <SourceLocation>
                    <Page>
                        <SourceLocation>
                    <Metadata>
                        <SourceLocation>
                    <Namespace>
            <DesktopFormFactor>
                <GetStarted>
                    <Title>
                    <Description>
                    <LearnMoreUrl>
                <FunctionFile>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                    <CustomTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                        <Label>
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
        <Resources>
            <Images>
                <Image>
                    <Override>
            <Urls>
                <Url>
                    <Override>
            <ShortStrings>
                <String>
                    <Override>
            <LongStrings>
                <String>
                    <Override>
        <WebApplicationInfo>
            <Id>
            <MsaId>
            <Resource>
            <Scopes>
                <Scope>
            <Authorizations>
                <Authorization>
                    <Resource>
                    <Scopes>
                        <Scope>
        <EquivalentAddins>
            <EquivalentAddin>
                <ProgId>
                <DisplayName>
                <FileName>
                <Type>
```

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="62415-125">Classement des éléments de complément de messagerie dans VersionOverrides ver.</span><span class="sxs-lookup"><span data-stu-id="62415-125">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="62415-126">1.0</span><span class="sxs-lookup"><span data-stu-id="62415-126">1.0</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <DesktopFormFactor>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>
                                <Action>
                                    <SourceLocation>
                                    <FunctionName>
                    <CustomTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>  
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                <Items>
                                    <Item>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <Label>
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <VersionOverrides>*
```

<span data-ttu-id="62415-127">\* Un VersionOverrides avec une `type` valeur `VersionOverridesV1_1` , au lieu de `VersionOverridesV1_0` , peut être imbriqué à la fin de l’VersionOverrides externe.</span><span class="sxs-lookup"><span data-stu-id="62415-127">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="62415-128">Consultez la rubrique ordre des éléments de [complément de messagerie dans VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) pour l’ordre des éléments dans `VersionOverridesV1_1` .</span><span class="sxs-lookup"><span data-stu-id="62415-128">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="62415-129">Classement des éléments de complément de messagerie dans VersionOverrides ver.</span><span class="sxs-lookup"><span data-stu-id="62415-129">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="62415-130">1.1</span><span class="sxs-lookup"><span data-stu-id="62415-130">1.1</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
    <Sets>
        <Set>
    <Hosts>
    <Host>
        <DesktopFormFactor>
            <ExtensionPoint>
                <OfficeTab>
                    <Group>
                        <Label>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <SourceLocation>
                                <FunctionName>
                <CustomTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                    <Label>
                <OfficeMenu>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                    <SourceLocation>
                <SourceLocation>
                <Label>
                <CommandSurface>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="content-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="62415-131">Classement des éléments de complément de contenu dans VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="62415-131">Content add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a><span data-ttu-id="62415-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="62415-132">See also</span></span>

- [<span data-ttu-id="62415-133">Référence de schéma pour les manifestes des compléments Office (version 1.1)</span><span class="sxs-lookup"><span data-stu-id="62415-133">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
