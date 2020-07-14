---
title: Commandes de complément Outlook
description: Les commandes de complément Outlook permettent de lancer des actions de complément spécifiques à partir du ruban en ajoutant des boutons ou des menus déroulants.
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 7705c168077d2a704ff16b05bfb82416cd7f4154
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094028"
---
# <a name="add-in-commands-for-outlook"></a>Commandes de complément pour Outlook

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> Les commandes complémentaires sont disponibles uniquement dans Outlook 2013 ou plus récent sur Windows, Outlook 2016 ou plus récent sur Mac, Outlook sur iOS, Outlook sur Android, Outlook sur le web pour Exchange 2016 ou plus récent, et Outlook sur le web pour Microsoft 365 et Outlook.com.
>
> Pour qu’Outlook 2013 prenne en charge les commandes de complément, trois mises à jour doivent être installées :
> - [Mise à jour de sécurité pour Outlook du 8 mars 2016](https://support.microsoft.com/kb/3114829)
> - [Mise à jour de sécurité pour Office du 8 mars 2016 (KB3114816)](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [Mise à jour de sécurité pour Office du 8 mars 2016 (KB3114828)](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> La prise en charge des commandes de complément dans Exchange 2016 nécessite la [mise à jour cumulative 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).

Add-in commands are only available for add-ins that do not use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](activation-rules.md) to limit the types of items they activate on. However, [contextual add-ins](contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).

## <a name="creating-the-add-in-command"></a>Création d’une commande de complément

Add-in commands are declared in the add-in manifest in the [VersionOverrides element](../reference/manifest/versionoverrides.md). This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support `VersionOverrides`, existing add-ins will continue to function as they did without add-in commands.

Les entrées de manifeste `VersionOverrides` spécifient plusieurs éléments pour le complément, notamment l’hôte, les types de contrôles à ajouter au ruban, le texte, les icônes et toutes les fonctions associées.

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>Comment les commandes de complément apparaissent-elles ?

An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the **Home** or **Message** tab. For the calendar, the default is the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tab. For module extensions, the default is a custom tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.

Au fur et à mesure du ruban, les commandes de complément s’affichent dans le menu de dépassement de capacité. Les commandes de complément pour un complément sont généralement regroupées.

![Boutons de commande du complément sur le ruban](../images/commands-normal.png)

![Boutons de commande du complément sur le ruban et le menu de dépassement de capacité](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>Outlook moderne sur le web

Dans Outlook sur le web, le nom du complément s’affiche dans un menu de dépassement de capacité. Si le complément inclut plusieurs commandes de complément, vous pouvez développer le menu du complément pour afficher le groupe de boutons portant le nom du complément.

![Menu de dépassement de capacité dans lequel se trouvent les boutons de commande du complément](../images/commands-overflow-menu-web.png)

![Menu de dépassement de capacité affichant les boutons de commande du complément](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a>Quelles formes d’expérience utilisateur existent pour les commandes de complément ?

The UX shape for an add-in command consists of a ribbon tab in the host application that contains buttons that can perform various functions. Currently, three UI shapes are supported:

- Un bouton qui exécute une fonction JavaScript
- Un bouton qui lance un volet Office
- Un bouton qui affiche un menu déroulant avec un ou plusieurs boutons des deux autres types

### <a name="executing-a-javascript-function"></a>Exécuter une fonction JavaScript

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

Dans les extensions de module, le bouton de commande de complément peut exécuter les fonctions JavaScript qui interagissent avec le contenu de l’interface utilisateur principale.

![Bouton exécutant une fonction sur le ruban Outlook.](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a>Lancement d’un volet Office

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![Bouton permettant d’ouvrir un volet Office sur le ruban Outlook.](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>Menu déroulant

A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.

![Bouton permettant de développer un menu sur le ruban Outlook.](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>Où les commandes de complément apparaissent-elles dans l’interface utilisateur ?

Les commandes de complément sont prises en charge pour quatre scénarios :

### <a name="reading-a-message"></a>Lecture d’un message

Lorsque l’utilisateur lit un message dans le volet de lecture ou dans l’onglet **Message** pour un formulaire de lecture contextuel, les commandes de complément ajoutées à l’onglet par défaut apparaissent dans l’onglet **Accueil**.

### <a name="composing-a-message"></a>Composition d’un message

Lorsque l’utilisateur crée un message, les commandes de complément ajoutées à l’onglet par défaut apparaissent dans l’onglet **Message**.

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a>Création ou affichage d’un rendez-vous ou d’une réunion en tant qu’organisateur

When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.

### <a name="viewing-a-meeting-as-an-attendee"></a>Affichage d’une réunion en tant que participant

When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon

### <a name="using-a-module-extension"></a>Utilisation d’une extension de module

Quand vous utilisez une extension de module, les commandes de complément apparaissent dans l’onglet personnalisé de l’extension.

## <a name="see-also"></a>Voir aussi

- [Démonstration de la commande de l'add-in Add-in Outlook](https://github.com/officedev/outlook-add-in-command-demo)
- [Créer des commandes complémentaires dans votre formulaire pour Excel, PowerPoint et Word](../develop/create-addin-commands.md)
