---
title: Implémenter un volet Office épinglable dans un complément Outlook
description: La commande de forme UX taskpane pour complément ouvre un volet Office vertical à droite d’un message ou demande de réunion, ce qui permet au complément de fournir une interface utilisateur pour des interactions plus détaillées.
ms.date: 02/28/2020
localization_priority: Normal
ms.openlocfilehash: ea9dc255bfb3b689a05d880007282da011edef3e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605317"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Implémenter un volet Office épinglable dans Outlook

La commande de forme UX [taskpane](add-in-commands-for-outlook.md#launching-a-task-pane) pour complément ouvre un volet Office vertical à droite d’un message ou demande de réunion, ce qui permet au complément de fournir une interface utilisateur pour des interactions plus détaillées (remplissage de plusieurs champs, etc.). Ce volet Office peut être affiché dans le volet de lecture lorsque vous affichez une liste des messages, ce qui permet un traitement rapide d’un message.

Toutefois, par défaut, si un utilisateur a un complément de volet Office ouvert pour un message dans le volet de lecture et sélectionne un nouveau message, le volet Office est automatiquement fermé. Pour un complément très sollicité, l’utilisateur peut préférer conserver ce volet ouvert, supprimant ainsi le besoin de réactiver le complément sur chaque message. Avec les volets Office épinglables, votre complément peut donner à l’utilisateur cette option.

> [!NOTE]
> Bien que la fonctionnalité des volets des tâches épinglables ait été introduite dans l' [ensemble de conditions requises 1,5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), elle est actuellement uniquement disponible pour les abonnés Office 365 à l’aide de la commande suivante.
> - Outlook 2016 ou version ultérieure sur Windows (Build 7668,2000 ou version ultérieure pour les utilisateurs des canaux actifs ou Office Insider, générer 7900. xxxx ou une version ultérieure pour les utilisateurs des canaux différés)
> - Outlook 2016 ou version ultérieure sur Mac (version 16.13.503 ou ultérieure)
> - Outlook moderne sur le web

> [!IMPORTANT]
> Les volets des tâches pouvant être épinglés ne sont pas disponibles pour les éléments suivants.
> - Rendez-vous/réunions
> - Outlook.com

## <a name="support-task-pane-pinning"></a>Prise en charge de l’épinglage des volets des tâches

La première étape consiste à ajouter une prise en charge de l’épinglage, ce qui est effectué dans le [manifeste](manifests.md) du complément. Cette opération est effectuée en ajoutant l’élément [SupportsPinning](../reference/manifest/action.md#supportspinning) à l’élément `Action` qui décrit le bouton du volet Office.

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

Le gestionnaire d’événements doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` de cet objet est réglée sur `Office.EventType.ItemChanged`. Lorsque l’événement est appelé, l’objet `Office.context.mailbox.item` est déjà mis à jour pour refléter l’élément actuellement sélectionné.

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

Utilisez la méthode [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour inscrire votre gestionnaire d’événements pour l’événement `Office.EventType.ItemChanged`. Cette opération doit être effectuée dans la fonction `Office.initialize` de votre volet Office.

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
