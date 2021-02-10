---
title: Comment trouver l’ordre approprié d’éléments manifeste
description: Découvrez comment trouver l’ordre correct dans lequel placer les éléments enfants dans un élément parent.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 2ee80167a76861209e814dc6c272720feb3a9cf1
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173912"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="2565a-103">Comment trouver l’ordre approprié d’éléments manifeste</span><span class="sxs-lookup"><span data-stu-id="2565a-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="2565a-104">Les éléments XML dans le fichier manifeste d’un complément Office doivent être sous l’élément parent approprié *et* dans un ordre spécifique, par rapport à d’autres, sous le parent.</span><span class="sxs-lookup"><span data-stu-id="2565a-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="2565a-105">Le classement requis est spécifié dans les fichiers XSD dans le dossier [schémas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="2565a-105">The required ordering is specified in the XSD files in the [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) folder.</span></span> <span data-ttu-id="2565a-106">Les fichiers XSD sont classés dans des sous-dossiers pour volet de tâches, contenu et compléments de courrier.</span><span class="sxs-lookup"><span data-stu-id="2565a-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="2565a-107">Par exemple, dans l’`<OfficeApp>`élément, le `<Id>`,`<Version>` ,`<ProviderName>` doit apparaître dans cet ordre.</span><span class="sxs-lookup"><span data-stu-id="2565a-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="2565a-108">Si un élément `<AlternateId>` est ajouté, il doit être compris entre l’élément `<Id>` et `<Version>`.</span><span class="sxs-lookup"><span data-stu-id="2565a-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="2565a-109">Votre manifeste ne sera pas valide et votre complément ne sera pas chargé, si un élément n’est pas dans l’ordre.</span><span class="sxs-lookup"><span data-stu-id="2565a-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="2565a-110">Le [validateur dans office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) utilise le même message d’erreur lorsqu’un élément est hors de l’ordre que lorsqu’un élément se trouve sous le mauvais parent.</span><span class="sxs-lookup"><span data-stu-id="2565a-110">The [validator within office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="2565a-111">L’erreur indique que l’élément enfant n’est pas un enfant valide de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="2565a-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="2565a-112">Si vous recevez un message d’erreur mais que la documentation de référence pour l’élément enfant indique qu’elle *est* valide pour le parent, alors le problème est probablement que l’enfant a été placé dans l’ordre incorrect.</span><span class="sxs-lookup"><span data-stu-id="2565a-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="2565a-113">Les sections suivantes montrent les éléments de manifeste dans l’ordre dans lequel ils doivent apparaître.</span><span class="sxs-lookup"><span data-stu-id="2565a-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="2565a-114">Il existe des différences selon que `type` l’attribut de `<OfficeApp>` l’élément est , ou `TaskPaneApp` `ContentApp` `MailApp` .</span><span class="sxs-lookup"><span data-stu-id="2565a-114">There are differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="2565a-115">Pour éviter que ces sections ne deviennent trop complexes, l’élément hautement complexe est `<VersionOverrides>` décomposé en sections distinctes.</span><span class="sxs-lookup"><span data-stu-id="2565a-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="2565a-116">Tous les éléments affichés ne sont pas obligatoires.</span><span class="sxs-lookup"><span data-stu-id="2565a-116">Not all of the elements shown are mandatory.</span></span> <span data-ttu-id="2565a-117">Si la `minOccurs` valeur d’un élément [](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3) **est 0** dans le schéma, l’élément est facultatif.</span><span class="sxs-lookup"><span data-stu-id="2565a-117">If the `minOccurs` value for a element is **0** in the [schema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="2565a-118">Ordre des éléments de l’élément de volet De tâches de base</span><span class="sxs-lookup"><span data-stu-id="2565a-118">Basic task pane add-in element ordering</span></span>

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

<span data-ttu-id="2565a-119">\*Voir l’ordre des éléments de l’élément du volet Des tâches dans [VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) pour l’ordre des éléments enfants de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="2565a-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="2565a-120">Ordre des éléments de messagerie de base</span><span class="sxs-lookup"><span data-stu-id="2565a-120">Basic mail add-in element ordering</span></span>

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

<span data-ttu-id="2565a-121">\*Voir l’ordre des éléments de l’élément De messagerie dans [VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) et l’ordre des éléments de la messagerie dans [VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) pour l’ordre des éléments enfants de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="2565a-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="2565a-122">Ordre des éléments de contenu de base des éléments</span><span class="sxs-lookup"><span data-stu-id="2565a-122">Basic content add-in element ordering</span></span>

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

<span data-ttu-id="2565a-123">\*Voir l’ordre des éléments de contenu dans [VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) pour l’ordre des éléments enfants de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="2565a-123">\*See [Content add-in element ordering within VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="2565a-124">Ordre des éléments de l’élément du volet De tâches dans VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="2565a-124">Task pane add-in element ordering within VersionOverrides</span></span>

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
                        <OverriddenByRibbonApi>
                        <Group> (can be below <ControlGroup>)
                            <OverriddenByRibbonApi>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <OverriddenByRibbonApi>
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
                                        <OverriddenByRibbonApi>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <ControlGroup> (can be above <Group>)
                        <Label>
                        <InsertAfter> (or <InsertBefore>)
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="2565a-125">Ordre des éléments du add-in de messagerie dans VersionOverrides Ver.</span><span class="sxs-lookup"><span data-stu-id="2565a-125">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="2565a-126">1.0</span><span class="sxs-lookup"><span data-stu-id="2565a-126">1.0</span></span>

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

<span data-ttu-id="2565a-127">\* Une versionOverrides avec la valeur, au lieu de , peut être imbriée à la fin `type` `VersionOverridesV1_1` de `VersionOverridesV1_0` l’extérieur VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="2565a-127">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="2565a-128">Voir l’ordre des éléments de la messagerie dans [VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) pour l’ordre des éléments dans `VersionOverridesV1_1` .</span><span class="sxs-lookup"><span data-stu-id="2565a-128">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="2565a-129">Ordre des éléments du add-in de messagerie dans VersionOverrides Ver.</span><span class="sxs-lookup"><span data-stu-id="2565a-129">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="2565a-130">1.1</span><span class="sxs-lookup"><span data-stu-id="2565a-130">1.1</span></span>

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

## <a name="content-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="2565a-131">Ordre des éléments de l’élément de contenu dans VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="2565a-131">Content add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a><span data-ttu-id="2565a-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2565a-132">See also</span></span>

- [<span data-ttu-id="2565a-133">Référence pour les manifestes des add-ins Office (v1.1)</span><span class="sxs-lookup"><span data-stu-id="2565a-133">Reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="2565a-134">Définitions de schéma officiel</span><span class="sxs-lookup"><span data-stu-id="2565a-134">Official schema definitions</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
