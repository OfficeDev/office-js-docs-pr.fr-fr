---
title: Comment trouver l’ordre approprié d’éléments manifeste
description: Découvrez comment trouver l’ordre correct dans lequel placer les éléments enfants dans un élément parent.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: d418f796592a0e4c247e717a5ce75d1c40c18d79
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302573"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Comment trouver l’ordre approprié d’éléments manifeste

Les éléments XML dans le fichier manifeste d’un complément Office doivent être sous l’élément parent approprié *et* dans un ordre spécifique, par rapport à d’autres, sous le parent.

Le classement requis est spécifié dans les fichiers XSD dans le dossier [schémas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Les fichiers XSD sont classés dans des sous-dossiers pour volet de tâches, contenu et compléments de courrier.

Par exemple, dans l’`<OfficeApp>`élément, le `<Id>`,`<Version>` ,`<ProviderName>` doit apparaître dans cet ordre. Si un élément `<AlternateId>` est ajouté, il doit être compris entre l’élément `<Id>` et `<Version>`. Votre manifeste ne sera pas valide et votre complément ne sera pas chargé, si un élément n’est pas dans l’ordre.

> [!NOTE]
> Le [validateur au sein de la boîte à outils Office](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-toolbox) utilise le même message d’erreur lorsqu’un élément est absent de l’ordre lorsqu’un élément est sous un parent incorrect. L’erreur indique que l’élément enfant n’est pas un enfant valide de l’élément parent. Si vous recevez un message d’erreur mais que la documentation de référence pour l’élément enfant indique qu’elle *est* valide pour le parent, alors le problème est probablement que l’enfant a été placé dans l’ordre incorrect.

Les sections suivantes présentent les éléments de manifeste dans l’ordre dans lequel ils doivent apparaître. Il existe de légères différences selon que l' `type` attribut de l' `<OfficeApp>` élément est `TaskPaneApp`, `ContentApp`ou `MailApp`. Pour éviter que ces sections deviennent trop encombrantes, l’élément hautement complexe `<VersionOverrides>` est divisé en sections distinctes.

> [!Note]
> Tous les éléments affichés ne sont pas obligatoires. Si la `minOccurs` valeur d’un élément est **0** dans le [schéma](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), l’élément est facultatif.

## <a name="basic-task-pane-add-in-element-ordering"></a>Classement des éléments de complément du volet Office de base

```
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
```

\*Voir classement des éléments de [complément du volet Office dans VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) pour l’ordre des éléments enfants de VersionOverrides.

## <a name="basic-mail-add-in-element-ordering"></a>Classement des éléments des compléments de messagerie de base

```
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

\*Consultez l’ordre des éléments de compléments de [messagerie dans VersionOverrides ver. 1,0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) et classement des éléments de [complément de messagerie dans VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) pour l’ordre des éléments enfants de VersionOverrides.

## <a name="basic-content-add-in-element-ordering"></a>Classement des éléments de complément de contenu de base

```
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
    <VersionOverrides>
```

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>Classement des éléments de complément du volet Office dans VersionOverrides

```
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
      <Hosts>
        <Host>
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>Classement des éléments de complément de messagerie dans VersionOverrides ver. 1.0

```
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

\*Un VersionOverrides avec `type` une `VersionOverridesV1_1`valeur, au `VersionOverridesV1_0`lieu de, peut être imbriqué à la fin de l’VersionOverrides externe. Consultez la rubrique ordre des éléments de [complément de messagerie dans VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) pour l’ordre `VersionOverridesV1_1`des éléments dans.

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>Classement des éléments de complément de messagerie dans VersionOverrides ver. 1.1

```
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

## <a name="see-also"></a>Voir aussi

- [Référence de schéma pour les manifestes des compléments Office (version 1.1)](../develop/add-in-manifests.md)
