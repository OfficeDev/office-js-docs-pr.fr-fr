---
title: Ajout d’une prise en charge mobile pour un complément Outlook
description: L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.
ms.date: 07/16/2021
localization_priority: Normal
ms.openlocfilehash: 270042d61077ae28abee79db024243bfbd5b6dc2
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937313"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Ajouter la prise en charge des commandes de complément pour Outlook Mobile

L’utilisation des commandes de Outlook Mobile permet à vos utilisateurs d’accéder aux mêmes fonctionnalités (avec certaines [limitations)](#code-considerations)que dans Outlook sur le web, Windows et Mac. L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.

## <a name="updating-the-manifest"></a>Mise à jour du manifeste

La première étape de l’activation des commandes de complément dans Outlook Mobile est de les définir dans le manifeste du complément. Le schéma [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 définit un nouveau facteur de forme pour les versions mobiles, [MobileFormFactor](../reference/manifest/mobileformfactor.md).

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

Cet exemple est semblable aux éléments qui apparaissent dans un élément [DesktopFormFactor](../reference/manifest/desktopformfactor.md), avec toutefois quelques différences importantes.

- L’élément [OfficeTab](../reference/manifest/officetab.md) n’est pas utilisé.
- L’élément [ExtensionPoint](../reference/manifest/extensionpoint.md) doit avoir un seul élément enfant. Si le complément ajoute uniquement un bouton, l’élément enfant doit être un élément [Control](../reference/manifest/control.md). Si le complément ajoute plusieurs boutons, l’élément enfant doit être un élément [Group](../reference/manifest/group.md) qui contient plusieurs éléments `Control`.
- Il n’existe aucun équivalent de type `Menu` pour l’élément `Control`.
- L’élément [Supertip](../reference/manifest/supertip.md) n’est pas utilisé.
- Les tailles d’icône requises sont différentes. Au minimum, les compléments mobiles doivent prendre en charge les icônes 25 x 25, 32 x 32 et 48 x 48 pixels.

## <a name="code-considerations"></a>Éléments à prendre en compte pour le code

La conception d’un complément pour mobile implique certaines considérations supplémentaires.

### <a name="use-rest-instead-of-exchange-web-services"></a>Utiliser REST plutôt que les services web Exchange

La méthode [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) n’est pas prise en charge dans Outlook Mobile. Les compléments doivent privilégier l’obtention d’informations auprès de l’API Office.js lorsque cela est possible. Si les compléments requièrent des informations non exposées par l’API Office.js, ils doivent utiliser les [API REST Outlook](/outlook/rest/) pour accéder à la boîte aux lettres de l’utilisateur.

L’ensemble de conditions requises de la boîte aux lettres 1.5 a introduit une nouvelle version de [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) qui peut demander un jeton d’accès compatible avec les API REST, ainsi qu’une nouvelle propriété [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) qui peut être utilisée pour rechercher le point de terminaison de l’API REST pour l’utilisateur.

### <a name="pinch-zoom"></a>Pincer pour zoomer

Par défaut les utilisateurs peuvent utiliser le mouvement pincer pour zoomer sur les volets Office. Si ce mouvement n’est pas pertinent pour votre scénario, veillez à désactiver la fonction « pincer pour zoomer » dans votre code HTML.

### <a name="close-task-panes"></a>Fermeture des volets Office

Dans Outlook Mobile, les volets Office occupent la totalité de l’écran et exigent par défaut que l’utilisateur les ferme pour revenir au message. Envisagez d’utiliser la méthode [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closeContainer__) pour fermer le volet Office lorsque votre scénario est terminé.

### <a name="compose-mode-and-appointments"></a>Mode composition et rendez-vous

Actuellement, les compléments dans Outlook Mobile ne peuvent être activés que lors de la lecture de messages. Les compléments ne sont pas activés lors de la composition des messages, ou lors de l’affichage ou de la rédaction des rendez-vous. Toutefois, les modules intégrés du fournisseur de réunions en ligne peuvent être activés en mode Organisateur de rendez-vous. Pour plus d’informations sur cette exception (y compris les API disponibles), reportez-vous à Créer un Outlook mobile pour un fournisseur de réunion [en ligne.](online-meeting.md#available-apis)

### <a name="unsupported-apis"></a>API non prises en charge

Les API introduites dans l’ensemble de conditions requises 1.6 ou ultérieure ne sont pas Outlook Mobile. Les API suivantes des ensembles de conditions requises antérieures ne sont pas non plus pris en charge.

- [Office.context.officeTheme](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
- [Office.context.mailbox.convertToEwsId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [Office.context.mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [Office.context.mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
- [Office.context.mailbox.item.displayReplyAllForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [Office.context.mailbox.item.displayReplyForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [Office.context.mailbox.item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [Office.context.mailbox.item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [Office.context.mailbox.item.getRegexMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a>Voir aussi

[Ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)