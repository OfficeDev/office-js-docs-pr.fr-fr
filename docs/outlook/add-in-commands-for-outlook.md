---
title: Commandes de complément Outlook
description: Les commandes de complément Outlook permettent de lancer des actions de complément spécifiques à partir du ruban en ajoutant des boutons ou des menus déroulants.
ms.date: 10/11/2022
ms.localizationpriority: high
ms.openlocfilehash: d029fd4acc1a32c912c73d6e5f468b9c217b9262
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546458"
---
# <a name="add-in-commands-for-outlook"></a>Commandes de complément pour Outlook

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> Les commandes complémentaires sont disponibles uniquement dans Outlook 2013 ou plus récent sur Windows, Outlook 2016 ou plus récent sur Mac, Outlook sur iOS, Outlook sur Android, Outlook sur le web pour Exchange 2016 ou plus récent, et Outlook sur le web pour Microsoft 365 et Outlook.com.
>
> Pour qu’Outlook 2013 prenne en charge les commandes de complément, trois mises à jour doivent être installées :
> - [Mise à jour de sécurité pour Outlook du 8 mars 2016](https://support.microsoft.com/kb/3114829)
> - [Mise à jour de sécurité pour Office du 8 mars 2016 (KB3114816)](https://support.microsoft.com/topic/3d3eb171-78c2-0e61-62a2-85723bc4bcc0)
> - [Mise à jour de sécurité pour Office du 8 mars 2016 (KB3114828)](https://support.microsoft.com/topic/54437016-d1e0-7aac-dbb7-4ecfbd57f5f0)
>
> La prise en charge des commandes de complément dans Exchange 2016 nécessite la [mise à jour cumulative 5](https://support.microsoft.com/topic/d67d7693-96a4-fb6e-b60b-e64984e267bd).

> [!TIP]
> Si votre complément utilise un manifeste XML, les commandes de complément sont uniquement disponibles pour les compléments qui n’utilisent pas [les règles ItemHasAttachment, ItemHasKnownEntity ou ItemHasRegularExpressionMatch](activation-rules.md) pour limiter les types d’éléments sur lesquels ils s’activent. Toutefois, [les compléments contextuels](contextual-outlook-add-ins.md) peuvent présenter différentes commandes selon que l’élément actuellement sélectionné est un message ou un rendez-vous, et peuvent choisir d’apparaître dans des scénarios de lecture ou de composition. L’utilisation des commandes de complément constitue une [meilleure pratique](../concepts/add-in-development-best-practices.md).

## <a name="create-the-ui-for-the-add-in-command"></a>Créer l’interface utilisateur de la commande de complément

Les commandes de complément sont déclarées dans le manifeste du complément. Le balisage dépend du type de manifeste.

# <a name="xml-manifest"></a>[Manifeste XML](#tab/xmlmanifest)

Les commandes de complément sont déclarées dans [l’élément VersionOverrides](/javascript/api/manifest/versionoverrides). Cet élément est un ajout au schéma de manifeste XML v1.1 qui garantit la compatibilité descendante. Dans un client qui ne prend pas en charge **\<VersionOverrides\>**, les compléments existants continuent à fonctionner comme ils le feraient sans commande de complément.

Les entrées de manifeste **\<VersionOverrides\>** spécifient plusieurs éléments pour le complément, notamment l’application, les types de contrôles à ajouter au ruban, le texte, les icônes et toutes les fonctions associées.

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

# <a name="teams-manifest-developer-preview"></a>[Manifeste Teams (préversion du développeur)](#tab/jsonmanifest)

Les commandes de complément sont déclarées avec les propriétés « extensions.runtimes » et « extensions.ribbons ». Ces propriétés spécifient de nombreuses choses pour le complément, telles que l’application, les types de contrôles à ajouter au ruban, le texte, les icônes et toutes les fonctions associées.

Lorsqu’un complément doit fournir des mises à jour d’état, telles que des indicateurs de progression ou des messages d’erreur, il doit le faire à travers les [API de notification](/javascript/api/outlook/office.notificationmessages). Le traitement des notifications doit également être défini dans un fichier HTML distinct spécifié dans la propriété « runtimes.code.page » du manifeste.

---
### <a name="icons"></a>Icônes

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>Comment les commandes de complément apparaissent-elles ?

Une commande de complément s’affiche sur le ruban sous la forme d’un bouton ou d’un élément dans un menu déroulant. Lorsqu’un utilisateur installe un complément, ses commandes apparaissent dans l’interface utilisateur sous la forme d’un groupe de boutons. Il peut s'agir de l'onglet par défaut du ruban ou d'un onglet personnalisé. Pour les messages, l'onglet par défaut est l'onglet **Accueil** ou **Message**. Pour le calendrier, l'onglet par défaut est **Réunion**, **Occurrence de réunion**, **Série de réunions** ou **Rendez-vous**. Pour les extensions de module, la valeur par défaut est un onglet personnalisé. Dans l'onglet par défaut, chaque module complémentaire peut avoir un groupe de rubans comportant jusqu'à 6 commandes. Dans les onglets personnalisés, le complément peut avoir jusqu’à 10 groupes, avec 6 commandes chacun. Les compléments sont limités à un seul onglet personnalisé.

Au fur et à mesure du ruban, les commandes de complément s’affichent dans le menu de dépassement de capacité. Les commandes de complément pour un complément sont généralement regroupées.

![Boutons de commande des modules complémentaires sur le ruban.](../images/commands-normal.png)

![Boutons de commande des modules complémentaires sur le ruban et dans le menu déroulant.](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>Outlook moderne sur le web

Dans Outlook sur le web, le nom du complément s’affiche dans un menu de dépassement de capacité. Si le complément inclut plusieurs commandes de complément, vous pouvez développer le menu du complément pour afficher le groupe de boutons portant le nom du complément.

![Menu de débordement où se trouvent les boutons de commande du module d'extension.](../images/commands-overflow-menu-web.png)

![Menu de débordement affichant les boutons de commande des modules complémentaires.](../images/commands-overflow-menu-expand-web.png)

## <a name="what-are-the-types-of-add-in-commands"></a>Quels sont les types de commandes de complément?

L’interface utilisateur d’une commande de complément se compose d’un bouton de ruban ou d’un élément dans un menu déroulant. Il existe deux types de commandes de complément en fonction du type d’action déclenchée par la commande.

- **Commandes du volet Office** : le bouton ou l’élément de menu qui ouvre le volet Office du complément. Vous ajoutez ce type de commande de complément avec des marques dans le manifeste. Le « code-behind » de la commande est fourni par Office.
- **Commandes de fonction** : le bouton ou l’élément de menu exécute n’importe quel Code JavaScript arbitraire. Le code appelle presque toujours des API dans la bibliothèque JavaScript Office, mais cela n’est pas nécessaire. Ce type de complément n’affiche généralement aucune autre interface utilisateur que le bouton ou l’élément de menu lui-même. Notez ce qui suit sur les commandes de fonction :

   - La fonction déclenchée peut appeler la méthode [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) pour afficher une boîte de dialogue, ce qui est un bon moyen d’afficher une erreur, d’afficher la progression ou d’inviter l’utilisateur à entrer des données.
   - Le runtime dans lequel la commande de fonction s’exécute est un [runtime complet basé sur un navigateur](../testing/runtimes.md#browser-runtime). Il peut afficher un code HTML et appeler Internet pour envoyer ou obtenir des données.

### <a name="run-a-function-command"></a>Exécuter une commande de fonction

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

Dans les extensions de module, le bouton de commande de complément peut exécuter les fonctions JavaScript qui interagissent avec le contenu de l’interface utilisateur principale.

![Bouton exécutant une fonction sur le ruban Outlook.](../images/commands-uiless-button-1.png)

### <a name="launch-a-task-pane"></a>Lancement d’un volet Office

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![Bouton permettant d’ouvrir un volet Office sur le ruban Outlook.](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>Menu déroulant

Une commande de complément de menu déroulant définit une liste statique d’éléments. Les boutons dans le menu peuvent correspondre à n’importe quelle combinaison de boutons qui exécutent une fonction ou qui ouvrent un volet Office. Les sous-menus ne sont pas pris en charge.

![Bouton permettant de développer un menu sur le ruban Outlook.](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>Où les commandes de complément apparaissent-elles dans l’interface utilisateur ?

Les commandes de complément sont prises en charge pour quatre scénarios :

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
- [Commandes de fonction de débogage dans les compléments Outlook](debug-ui-less.md)
- [Didacticiel : créer un complément de composition de message Outlook](../tutorials/outlook-tutorial.md)
