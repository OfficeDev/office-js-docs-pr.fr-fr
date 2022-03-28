---
title: Ajout d’une prise en charge mobile pour un complément Outlook
description: L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 451476e9cd7eda9902ad156966558c182752547f
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484303"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Ajouter la prise en charge des commandes de complément pour Outlook Mobile

L’utilisation des commandes de Outlook Mobile permet à vos utilisateurs d’accéder aux mêmes fonctionnalités (avec certaines [limitations](#code-considerations)) que dans Outlook sur le web, Windows et Mac. L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.

## <a name="updating-the-manifest"></a>Mise à jour du manifeste

La première étape de l’activation des commandes de complément dans Outlook Mobile est de les définir dans le manifeste du complément. Le schéma [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 définit un nouveau facteur de forme pour les versions mobiles, [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

Cet élément contient toutes les informations pour charger le complément dans des clients mobiles. Cela vous permet de définir entièrement différents éléments de l’interface utilisateur et fichiers JavaScript pour l’expérience mobile.

L’exemple suivant montre un bouton de volet de tâches unique dans un `MobileFormFactor` élément.

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
- L’élément [ExtensionPoint](/javascript/api/manifest/extensionpoint) doit avoir un seul élément enfant. Si le complément ajoute uniquement un bouton, l’élément enfant doit être un élément [Control](/javascript/api/manifest/control). Si le complément ajoute plusieurs boutons, l’élément enfant doit être un élément [Group](/javascript/api/manifest/group) qui contient plusieurs éléments `Control`.
- Il n’existe aucun équivalent de type `Menu` pour l’élément `Control`.
- L’élément [Supertip](/javascript/api/manifest/supertip) n’est pas utilisé.
- Les tailles d’icône requises sont différentes. Au minimum, les compléments mobiles doivent prendre en charge les icônes 25 x 25, 32 x 32 et 48 x 48 pixels.

## <a name="code-considerations"></a>Éléments à prendre en compte pour le code

La conception d’un complément pour mobile implique certaines considérations supplémentaires.

### <a name="use-rest-instead-of-exchange-web-services"></a>Utiliser REST plutôt que les services web Exchange

La méthode [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) n’est pas prise en charge dans Outlook Mobile. Les compléments doivent privilégier l’obtention d’informations auprès de l’API Office.js lorsque cela est possible. Si les compléments requièrent des informations non exposées par l’API Office.js, ils doivent utiliser les [API REST Outlook](/outlook/rest/) pour accéder à la boîte aux lettres de l’utilisateur.

L’ensemble de conditions requises de la boîte aux lettres 1.5 a introduit une nouvelle version de [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) qui peut demander un jeton d’accès compatible avec les API REST et une nouvelle propriété [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) qui peut être utilisée pour rechercher le point de terminaison de l’API REST pour l’utilisateur.

### <a name="pinch-zoom"></a>Pincer pour zoomer

Par défaut les utilisateurs peuvent utiliser le mouvement pincer pour zoomer sur les volets Office. Si ce mouvement n’est pas pertinent pour votre scénario, veillez à désactiver la fonction « pincer pour zoomer » dans votre code HTML.

### <a name="close-task-panes"></a>Fermeture des volets Office

Dans Outlook Mobile, les volets Office occupent la totalité de l’écran et exigent par défaut que l’utilisateur les ferme pour revenir au message. Envisagez d’utiliser la méthode [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) pour fermer le volet Office lorsque votre scénario est terminé.

### <a name="compose-mode-and-appointments"></a>Mode composition et rendez-vous

Actuellement, les compléments dans Outlook Mobile ne peuvent être activés que lors de la lecture de messages. Les compléments ne sont pas activés lors de la composition des messages, ou lors de l’affichage ou de la rédaction des rendez-vous. Toutefois, les modules intégrés du fournisseur de réunions en ligne peuvent être activés en mode Organisateur de rendez-vous. Pour plus d’informations sur cette exception (y compris les API disponibles), voir Créer un Outlook mobile pour un fournisseur de [réunion en ligne](online-meeting.md#available-apis).

### <a name="unsupported-apis"></a>API non prises en charge

Les API introduites dans l’ensemble de conditions requises 1.6 ou ultérieure ne sont pas Outlook Mobile. Les API suivantes des ensembles de conditions requises antérieures ne sont pas non plus pris en charge.

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

[Ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook](/javascript/api/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)