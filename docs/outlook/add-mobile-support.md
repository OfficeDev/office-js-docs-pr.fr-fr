---
title: Ajout d’une prise en charge mobile pour un complément Outlook
description: Découvrez comment ajouter la prise en charge d’Outlook Mobile, notamment comment mettre à jour le manifeste du complément et modifier votre code pour les scénarios mobiles, si nécessaire.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c84b4aeb04cd2c8b3c2f0a7afa9fd1631c22afc5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607542"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Ajouter la prise en charge des commandes de complément pour Outlook Mobile

L’utilisation de commandes de complément dans Outlook Mobile permet à vos utilisateurs d’accéder aux [mêmes fonctionnalités](#code-considerations) (avec certaines limitations) qu’ils ont déjà dans Outlook sur le web, Windows et Mac. L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.

## <a name="updating-the-manifest"></a>Mise à jour du manifeste

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

L’exemple suivant montre un bouton de volet Office unique dans un `MobileFormFactor` élément.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

Cet exemple est semblable aux éléments qui apparaissent dans un élément [DesktopFormFactor](/javascript/api/manifest/desktopformfactor), avec toutefois quelques différences importantes.

- L’élément [OfficeTab](/javascript/api/manifest/officetab) n’est pas utilisé.
- The [ExtensionPoint](/javascript/api/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](/javascript/api/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](/javascript/api/manifest/group) element that contains multiple `Control` elements.
- Il n’existe aucun équivalent de type `Menu` pour l’élément `Control`.
- L’élément [Supertip](/javascript/api/manifest/supertip) n’est pas utilisé.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## <a name="code-considerations"></a>Éléments à prendre en compte pour le code

La conception d’un complément pour mobile implique certaines considérations supplémentaires.

### <a name="use-rest-instead-of-exchange-web-services"></a>Utiliser REST plutôt que les services web Exchange

The [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.

L’ensemble de conditions requises de boîte aux lettres 1.5 a introduit une nouvelle version [d’Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) qui peut demander un jeton d’accès compatible avec les API REST et une nouvelle propriété [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) qui peut être utilisée pour rechercher le point de terminaison de l’API REST pour l’utilisateur.

### <a name="pinch-zoom"></a>Pincer pour zoomer

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### <a name="close-task-panes"></a>Fermeture des volets Office

In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) method to close the task pane when your scenario is complete.

### <a name="compose-mode-and-appointments"></a>Mode composition et rendez-vous

Actuellement, les compléments dans Outlook Mobile prennent uniquement en charge l’activation lors de la lecture des messages. Les compléments ne sont pas activés lors de la composition des messages, ou lors de l’affichage ou de la rédaction des rendez-vous. Toutefois, il existe deux exceptions :

1. Les compléments intégrés du fournisseur de réunions en ligne peuvent être activés en mode Organisateur de rendez-vous. Pour plus d’informations sur cette exception (y compris les API disponibles), consultez [Créer un complément mobile Outlook pour un fournisseur de réunions en ligne](online-meeting.md#available-apis).
1. Les compléments qui journalisent les notes de rendez-vous et d’autres détails sur la gestion des relations client (CRM) ou les services de prise de notes peuvent être activés en mode Participant au rendez-vous. Pour plus d’informations sur cette exception (y compris les API disponibles), [reportez-vous aux notes de rendez-vous du journal sur une application externe dans les compléments mobiles Outlook](mobile-log-appointments.md#available-apis).

### <a name="unsupported-apis"></a>API non prises en charge

Les API introduites dans l’ensemble de conditions requises 1.6 ou version ultérieure ne sont pas prises en charge par Outlook Mobile. Les API suivantes des ensembles de conditions requises antérieurs ne sont pas non plus prises en charge.

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## <a name="see-also"></a>Voir aussi

[Ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)