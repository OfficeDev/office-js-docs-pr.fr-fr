---
title: Implémenter un volet Office épinglable dans un complément Outlook
description: La commande de forme UX taskpane pour complément ouvre un volet Office vertical à droite d’un message ou demande de réunion, ce qui permet au complément de fournir une interface utilisateur pour des interactions plus détaillées.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39af3a532d553835b02709301c998a78dc9958bb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093867"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Implémenter un volet Office épinglable dans Outlook

The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Bien que la fonctionnalité des volets des tâches épinglables ait été introduite dans l' [ensemble de conditions requises 1,5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), elle est actuellement uniquement disponible pour les abonnés Microsoft 365 à l’aide des éléments suivants.
> - Outlook 2016 ou version ultérieure sur Windows (Build 7668,2000 ou version ultérieure pour les utilisateurs des canaux actifs ou Office Insider, générer 7900. xxxx ou une version ultérieure pour les utilisateurs des canaux différés)
> - Outlook 2016 ou version ultérieure sur Mac (version 16.13.503 ou ultérieure)
> - Outlook moderne sur le web

> [!IMPORTANT]
> Les volets des tâches pouvant être épinglés ne sont pas disponibles pour les éléments suivants.
> - Rendez-vous/réunions
> - Outlook.com

## <a name="support-task-pane-pinning"></a>Prise en charge de l’épinglage des volets des tâches

The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.

L’élément `SupportsPinning` est défini dans le schéma VersionOverrides v1.1, vous devez donc inclure un élément [VersionOverrides](../reference/manifest/versionoverrides.md) pour les versions 1.0 et 1.1.

> [!NOTE]
> Si vous envisagez de [publier](../publish/publish.md) votre complément Outlook sur [AppSource](https://appsource.microsoft.com), lorsque vous utilisez l’élément **SupportsPinning** afin d’obtenir la [validation d’AppSource](/legal/marketplace/certification-policies), le contenu de votre complément ne doit pas être statique et doit afficher clairement les données liées au message qui est ouvert ou sélectionné dans la boîte aux lettres.

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

Pour obtenir un exemple complet, consultez le contrôle `msgReadOpenPaneButton` dans l’[exemple de manifeste command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>Gestion des mises à jour de l’interface utilisateur en fonction du message actuellement sélectionné

Pour mettre à jour l’interface utilisateur ou les variables internes de votre volet Office en fonction de l’élément actif, vous devez enregistrer un gestionnaire d’événements pour être notifié de la modification.

### <a name="implement-the-event-handler"></a>Mettre en œuvre le gestionnaire d’événements

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> L’implémentation des gestionnaires d’événements pour un événement ItemChanged implique de vérifier si l’élément Office.content.mailbox.item est null.
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a>Enregistrement du gestionnaire d’événements

Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a>Voir aussi

Pour un exemple de complément qui implémente un volet Office épinglables, consultez [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) sur GitHub.
