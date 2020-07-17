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

Les commandes de complément Outlook permettent d’initier des actions de complément spécifiques à partir du ruban en ajoutant des boutons ou des menus déroulants. Les utilisateurs peuvent ainsi accéder aux compléments d’une manière simple, intuitive et discrète. Parce qu’elles offrent des fonctionnalités optimales en toute transparence, les commandes de complément vous permettent de créer des solutions plus attrayantes.

> [!NOTE]
> Les commandes complémentaires sont disponibles uniquement dans Outlook 2013 ou plus récent sur Windows, Outlook 2016 ou plus récent sur Mac, Outlook sur iOS, Outlook sur Android, Outlook sur le web pour Exchange 2016 ou plus récent, et Outlook sur le web pour Microsoft 365 et Outlook.com.
>
> Pour qu’Outlook 2013 prenne en charge les commandes de complément, trois mises à jour doivent être installées :
> - [Mise à jour de sécurité pour Outlook du 8 mars 2016](https://support.microsoft.com/kb/3114829)
> - [Mise à jour de sécurité pour Office du 8 mars 2016 (KB3114816)](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [Mise à jour de sécurité pour Office du 8 mars 2016 (KB3114828)](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> La prise en charge des commandes de complément dans Exchange 2016 nécessite la [mise à jour cumulative 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).

Les commandes de complément sont uniquement disponibles pour les compléments qui n’utilisent pas les règles [ItemHasAttachment, ItemHasKnownEntity ou ItemHasRegularExpressionMatch](activation-rules.md) pour limiter les types d’éléments sur lesquels elles s’activent. Toutefois, les [compléments contextuels](contextual-outlook-add-ins.md) peuvent présenter diverses commandes selon que l’élément actuellement sélectionné est un message ou un rendez-vous, et peuvent apparaître dans des scénarios de lecture ou de composition. L’utilisation des commandes de complément constitue une [meilleure pratique](../concepts/add-in-development-best-practices.md).

## <a name="creating-the-add-in-command"></a>Création d’une commande de complément

Les commandes de complément sont déclarées dans le manifeste de complément dans l’[élément VersionOverrides](../reference/manifest/versionoverrides.md). Cet élément est un ajout au schéma de manifeste version 1.1 qui assure la compatibilité descendante. Dans un client qui ne prend pas en charge `VersionOverrides`, les compléments existants continuent à fonctionner comme ils le feraient sans commande de complément.

Les entrées de manifeste `VersionOverrides` spécifient plusieurs éléments pour le complément, notamment l’hôte, les types de contrôles à ajouter au ruban, le texte, les icônes et toutes les fonctions associées.

Lorsqu’un complément doit fournir des mises à jour d’état, telles que des indicateurs de progression ou des messages d’erreur, il doit le faire via les [API de notification](/javascript/api/outlook/office.notificationmessages). Le traitement pour les notifications doit également être défini dans un fichier HTML distinct qui est spécifié dans le nœud `FunctionFile` du manifeste.

Les développeurs doivent définir des icônes pour toutes les tailles requises afin que les commandes de complément s’ajustent parfaitement avec le ruban. Les tailles d’icône requises sont 80 x 80 pixels, 32 x 32 pixels, 16 x 16 pixels pour les versions de bureau et 48 x 48 pixels, 32 x 32 pixels et 25 x 25 pixels pour les versions mobiles.

## <a name="how-do-add-in-commands-appear"></a>Comment les commandes de complément apparaissent-elles ?

Une commande de complément apparaît dans le ruban, comme un bouton. Lorsqu’un utilisateur installe un complément, ses commandes apparaissent dans l’interface utilisateur sous la forme d’un groupe de boutons. Le groupe peut apparaître dans l’onglet par défaut du ruban ou dans un onglet personnalisé. Pour les messages, il apparaît par défaut dans l’onglet **Accueil** ou **Message**. Pour le calendrier, il apparaît par défaut dans l’onglet **Réunion**, **Occurrence de réunion**, **Série de réunions** ou **Rendez-vous**. Pour les extensions de module, il apparaît par défaut dans un onglet personnalisé. Dans l’onglet par défaut, chaque complément peut avoir un groupe Ruban incluant 6 commandes maximum. Dans les onglets personnalisés, le complément peut avoir jusqu’à 10 groupes, avec 6 commandes chacun. Les compléments sont limités à un seul onglet personnalisé.

Au fur et à mesure du ruban, les commandes de complément s’affichent dans le menu de dépassement de capacité. Les commandes de complément pour un complément sont généralement regroupées.

![Boutons de commande du complément sur le ruban](../images/commands-normal.png)

![Boutons de commande du complément sur le ruban et le menu de dépassement de capacité](../images/commands-collapsed.png)

Lorsqu’une commande de complément est ajoutée à un complément, le nom du complément est supprimé de la barre d’application. Seul le bouton de commande du complément dans le ruban est conservé.

### <a name="modern-outlook-on-the-web"></a>Outlook moderne sur le web

Dans Outlook sur le web, le nom du complément s’affiche dans un menu de dépassement de capacité. Si le complément inclut plusieurs commandes de complément, vous pouvez développer le menu du complément pour afficher le groupe de boutons portant le nom du complément.

![Menu de dépassement de capacité dans lequel se trouvent les boutons de commande du complément](../images/commands-overflow-menu-web.png)

![Menu de dépassement de capacité affichant les boutons de commande du complément](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a>Quelles formes d’expérience utilisateur existent pour les commandes de complément ?

La forme d’expérience utilisateur d’une commande de complément inclut un onglet de ruban dans l’application hôte qui contient des boutons permettant d’effectuer diverses actions. Actuellement, trois formes d’expérience utilisateur sont prises en charge :

- Un bouton qui exécute une fonction JavaScript
- Un bouton qui lance un volet Office
- Un bouton qui affiche un menu déroulant avec un ou plusieurs boutons des deux autres types

### <a name="executing-a-javascript-function"></a>Exécuter une fonction JavaScript

Utilisez un bouton de commande de complément qui exécute une fonction JavaScript pour les scénarios dans lesquels l’utilisateur n’a pas besoin d’effectuer de sélections supplémentaires pour lancer l’action. Cela peut être utile, entre autres, pour les actions de suivi, de rappel, d’impression ou de scénario quand l’utilisateur souhaite obtenir des informations supplémentaires d’un service.

Dans les extensions de module, le bouton de commande de complément peut exécuter les fonctions JavaScript qui interagissent avec le contenu de l’interface utilisateur principale.

![Bouton exécutant une fonction sur le ruban Outlook.](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a>Lancement d’un volet Office

Utilisez un bouton de commande de complément pour lancer un volet Office pour les scénarios dans lesquels l’utilisateur doit interagir avec un complément pour une durée plus longue. Par exemple, le complément nécessite des modifications de paramètres ou la saisie de données dans de nombreux champs.

La largeur par défaut du volet de tâche vertical est de 320 px. Celui-ci peut être redimensionné à la fois dans l’explorateur Outlook et dans l’inspecteur. Le volet peut redimensionner le volet de tâche et l’affichage Liste de façon identique.

![Bouton permettant d’ouvrir un volet Office sur le ruban Outlook.](../images/commands-task-pane-button-1.png)

<br/>

Cette capture d’écran montre un exemple de volet Office vertical. Le volet s’ouvre avec le nom de la commande du complément dans le coin supérieur gauche. Les utilisateurs peuvent utiliser le bouton **X** situé dans le coin supérieur droit du volet pour fermer le complément lorsqu’ils ont terminé de l’utiliser. Par défaut, ce volet n’est pas conservé sur plusieurs messages. Les compléments peuvent [prendre en charge l’épinglage](pinnable-taskpane.md) pour le volet Office et recevoir des événements lorsqu’un nouveau message est sélectionné. Tous les éléments d’interface utilisateur affichés dans le volet Office, mis à part le nom et le bouton Fermer, sont fournis par le complément.

Si l’utilisateur sélectionne une autre commande de complément qui ouvre un volet Office, le volet est remplacé par la commande qui vient d’être utilisée. Si l’utilisateur sélectionne un bouton de commande de complément qui exécute une fonction ou sur un menu déroulant alors que le volet Office est ouvert, l’action est exécutée et le volet Office reste ouvert.

### <a name="drop-down-menu"></a>Menu déroulant

Une commande de complément de menu déroulant définit une liste statique de boutons. Les boutons dans le menu peuvent correspondre à n’importe quelle combinaison de boutons qui exécutent une fonction ou qui ouvrent un volet Office. Les sous-menus ne sont pas pris en charge.

![Bouton permettant de développer un menu sur le ruban Outlook.](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>Où les commandes de complément apparaissent-elles dans l’interface utilisateur ?

Les commandes de complément sont prises en charge pour quatre scénarios :

### <a name="reading-a-message"></a>Lecture d’un message

Lorsque l’utilisateur lit un message dans le volet de lecture ou dans l’onglet **Message** pour un formulaire de lecture contextuel, les commandes de complément ajoutées à l’onglet par défaut apparaissent dans l’onglet **Accueil**.

### <a name="composing-a-message"></a>Composition d’un message

Lorsque l’utilisateur crée un message, les commandes de complément ajoutées à l’onglet par défaut apparaissent dans l’onglet **Message**.

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a>Création ou affichage d’un rendez-vous ou d’une réunion en tant qu’organisateur

Lorsque vous créez ou visualisez un rendez-vous ou une réunion en tant qu’organisateur, les commandes de complément ajoutées à l’onglet par défaut apparaissent dans les onglets **Réunion**, **Occurrence de réunion**, **Série de réunions** ou **Rendez-vous** des formulaires contextuels. Toutefois, si l’utilisateur sélectionne un élément dans le calendrier sans ouvrir la fenêtre contextuelle, le groupe Ruban du complément n’apparaît pas sur le ruban.

### <a name="viewing-a-meeting-as-an-attendee"></a>Affichage d’une réunion en tant que participant

Lorsque vous visualisez une réunion en tant que participant, les commandes de complément ajoutées à l’onglet par défaut apparaissent dans les onglets **Réunion**, **Occurrence de réunion** ou **Série de réunions** des formulaires contextuels. Toutefois, si l’utilisateur sélectionne un élément dans le calendrier sans ouvrir la fenêtre contextuelle, le groupe Ruban du complément n’apparaît pas sur le ruban.

### <a name="using-a-module-extension"></a>Utilisation d’une extension de module

Quand vous utilisez une extension de module, les commandes de complément apparaissent dans l’onglet personnalisé de l’extension.

## <a name="see-also"></a>Voir aussi

- [Démonstration de la commande de l'add-in Add-in Outlook](https://github.com/officedev/outlook-add-in-command-demo)
- [Créer des commandes complémentaires dans votre formulaire pour Excel, PowerPoint et Word](../develop/create-addin-commands.md)
