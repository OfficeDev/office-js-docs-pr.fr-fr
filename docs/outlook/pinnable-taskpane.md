---
title: Implémenter un volet Office épinglable dans un complément Outlook
description: La commande de forme UX taskpane pour complément ouvre un volet Office vertical à droite d’un message ou demande de réunion, ce qui permet au complément de fournir une interface utilisateur pour des interactions plus détaillées.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 57a17a90fe565adb3ffb9d23e3b169bc83be2735
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348880"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="74a38-103">Implémenter un volet Office épinglable dans Outlook</span><span class="sxs-lookup"><span data-stu-id="74a38-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="74a38-p101">La commande de forme UX [taskpane](add-in-commands-for-outlook.md#launching-a-task-pane) pour complément ouvre un volet Office vertical à droite d’un message ou demande de réunion, ce qui permet au complément de fournir une interface utilisateur pour des interactions plus détaillées (remplissage de plusieurs champs, etc.). Ce volet Office peut être affiché dans le volet de lecture lorsque vous affichez une liste des messages, ce qui permet un traitement rapide d’un message.</span><span class="sxs-lookup"><span data-stu-id="74a38-p101">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="74a38-p102">Toutefois, par défaut, si un utilisateur a un complément de volet Office ouvert pour un message dans le volet de lecture et sélectionne un nouveau message, le volet Office est automatiquement fermé. Pour un complément très sollicité, l’utilisateur peut préférer conserver ce volet ouvert, supprimant ainsi le besoin de réactiver le complément sur chaque message. Avec les volets Office épinglables, votre complément peut donner à l’utilisateur cette option.</span><span class="sxs-lookup"><span data-stu-id="74a38-p102">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="74a38-109">Bien que la fonctionnalité des volets des tâches épinglables soit une nouveauté de l’ensemble de conditions requises [1.5,](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)elle n’est actuellement disponible que pour les abonnés Microsoft 365 utilisant ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="74a38-109">Although the pinnable task panes feature was introduced in [requirement set 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only available to Microsoft 365 subscribers using the following:</span></span>
>
> - <span data-ttu-id="74a38-110">Outlook 2016 ou version ultérieure sur Windows (build 7668.2000 ou ultérieure pour les utilisateurs des canaux Insider actuels ou Office, build 7900.xxxx ou version ultérieure pour les utilisateurs dans les canaux différés)</span><span class="sxs-lookup"><span data-stu-id="74a38-110">Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels)</span></span>
> - <span data-ttu-id="74a38-111">Outlook 2016 version ultérieure sur Mac (version 16.13.503 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="74a38-111">Outlook 2016 or later on Mac (version 16.13.503 or later)</span></span>
> - <span data-ttu-id="74a38-112">Outlook moderne sur le web</span><span class="sxs-lookup"><span data-stu-id="74a38-112">Modern Outlook on the web</span></span>

> [!IMPORTANT]
> <span data-ttu-id="74a38-113">Les volets Des tâches épinglables ne sont pas disponibles pour les tâches suivantes :</span><span class="sxs-lookup"><span data-stu-id="74a38-113">Pinnable task panes are not available for the following:</span></span>
>
> - <span data-ttu-id="74a38-114">Rendez-vous/réunions</span><span class="sxs-lookup"><span data-stu-id="74a38-114">Appointments/Meetings</span></span>
> - <span data-ttu-id="74a38-115">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="74a38-115">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="74a38-116">Prise en charge de l’épinglage des volets des tâches</span><span class="sxs-lookup"><span data-stu-id="74a38-116">Support task pane pinning</span></span>

<span data-ttu-id="74a38-p103">La première étape consiste à ajouter une prise en charge de l’épinglage, ce qui est effectué dans le [manifeste](manifests.md) du complément. Cette opération est effectuée en ajoutant l’élément [SupportsPinning](../reference/manifest/action.md#supportspinning) à l’élément `Action` qui décrit le bouton du volet Office.</span><span class="sxs-lookup"><span data-stu-id="74a38-p103">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="74a38-119">L’élément `SupportsPinning` est défini dans le schéma VersionOverrides v1.1, vous devez donc inclure un élément [VersionOverrides](../reference/manifest/versionoverrides.md) pour les versions 1.0 et 1.1.</span><span class="sxs-lookup"><span data-stu-id="74a38-119">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="74a38-120">Si vous envisagez de [publier](../publish/publish.md) votre complément Outlook sur [AppSource](https://appsource.microsoft.com), lorsque vous utilisez l’élément **SupportsPinning** afin d’obtenir la [validation d’AppSource](/legal/marketplace/certification-policies), le contenu de votre complément ne doit pas être statique et doit afficher clairement les données liées au message qui est ouvert ou sélectionné dans la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="74a38-120">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/legal/marketplace/certification-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

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

<span data-ttu-id="74a38-121">Pour obtenir un exemple complet, consultez le contrôle `msgReadOpenPaneButton` dans l’[exemple de manifeste command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="74a38-121">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="74a38-122">Gestion des mises à jour de l’interface utilisateur en fonction du message actuellement sélectionné</span><span class="sxs-lookup"><span data-stu-id="74a38-122">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="74a38-123">Pour mettre à jour l’interface utilisateur ou les variables internes de votre volet Office en fonction de l’élément actif, vous devez enregistrer un gestionnaire d’événements pour être notifié de la modification.</span><span class="sxs-lookup"><span data-stu-id="74a38-123">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="74a38-124">Mettre en œuvre le gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="74a38-124">Implement the event handler</span></span>

<span data-ttu-id="74a38-p104">Le gestionnaire d’événements doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` de cet objet est réglée sur `Office.EventType.ItemChanged`. Lorsque l’événement est appelé, l’objet `Office.context.mailbox.item` est déjà mis à jour pour refléter l’élément actuellement sélectionné.</span><span class="sxs-lookup"><span data-stu-id="74a38-p104">The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="74a38-128">L’implémentation des gestionnaires d’événements pour un événement ItemChanged implique de vérifier si l’élément Office.content.mailbox.item est null.</span><span class="sxs-lookup"><span data-stu-id="74a38-128">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="74a38-129">Enregistrement du gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="74a38-129">Register the event handler</span></span>

<span data-ttu-id="74a38-p105">Utilisez la méthode [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour inscrire votre gestionnaire d’événements pour l’événement `Office.EventType.ItemChanged`. Cette opération doit être effectuée dans la fonction `Office.initialize` de votre volet Office.</span><span class="sxs-lookup"><span data-stu-id="74a38-p105">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="74a38-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="74a38-132">See also</span></span>

<span data-ttu-id="74a38-133">Pour un exemple de complément qui implémente un volet Office épinglables, consultez [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) sur GitHub.</span><span class="sxs-lookup"><span data-stu-id="74a38-133">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
