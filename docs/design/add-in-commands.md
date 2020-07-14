---
title: Concepts basiques pour les commandes de complément
description: Découvrez l'ajout de boutons et d'éléments de menu personnalisés au ruban dans Office dans le cadre d’un complément Office.
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093874"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Commandes de complément pour Excel, PowerPoint et Word

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

Pour une vue d'ensemble du reportage, voir la vidéo [Ruban de l'application commandes complémentaires au sein du Bureau](https://channel9.msdn.com/events/Build/2016/P551).

> [!NOTE]
> SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.

> [!IMPORTANT]
> Les commandes de complément sont actuellement prises en charge dans Outlook. Pour plus d’informations, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).

*Figure 1. Complément incluant des commandes en cours d’exécution dans Excel (version de bureau)*

![Capture d’écran d’une commande de complément dans Excel](../images/add-in-commands-1.png)

*Figure 2. Complément incluant des commandes en cours d’exécution dans Excel sur le web*

![Capture d’écran d’une commande de complément dans Excel sur le web](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a>Fonctionnalités de commande

Les fonctionnalités de commande suivantes sont actuellement prises en charge.

> [!NOTE]
> Les compléments de contenu ne prennent actuellement pas en charge les commandes de complément.

### <a name="extension-points"></a>Points d’extension

- Onglets de ruban - Permet d’étendre les onglets prédéfinis ou de créer un onglet personnalisé.
- Menus contextuels - Permet d’étendre les menus contextuels sélectionnés.

### <a name="control-types"></a>Types de contrôles

- Boutons simples - Permettent de déclencher des actions spécifiques.
- Menus - Menu déroulant simple avec des boutons qui déclenchent des actions.

### <a name="actions"></a>Actions

- ShowTaskpane - Affiche un ou plusieurs volets où sont chargées des pages HTML personnalisées.
- ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.  

### <a name="default-enabled-or-disabled-status-preview"></a>État Activé ou Désactivé par défaut (préversion)

Vous pouvez spécifier si la commande est activée ou désactivée lors du lancement de votre complément et modifier le paramètre par programme.

> [!NOTE]
> Cette fonctionnalité est en préversion et n’est pas prise en charge dans tous les hôtes ou scénarios. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](disable-add-in-commands.md).

## <a name="supported-platforms"></a>Plateformes prises en charge

Les commandes de complément sont actuellement prises en charge sur les plateformes suivantes.

- Office on Windows (build 16.0.6769+, connecté à l'abonnement Microsoft 365)
- Office 2019 sur Windows
- Office sur Mac (build 15.33+, connecté à l'abonnement Microsoft 365)
- Office 2019 sur Mac
- Office sur le web

> [!NOTE]
> Pour plus d’informations sur la prise en charge dans Outlook, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).

## <a name="debugging"></a>Débogage

Pour déboguer une commande de complément, vous devez l’exécuter dans Office sur le web. Pour plus de détails, voir [Débogage de compléments dans Office sur le web](../testing/debug-add-ins-in-office-online.md).

## <a name="best-practices"></a>Meilleures pratiques

Appliquez les meilleures pratiques suivantes lorsque vous développez des commandes de complément :

- Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.
- Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.
- Pour le placement de vos commandes dans le ruban d'application de l'Office :
    - Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).
    - Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).  
    - Placez des commandes dans un onglet personnalisé si vous avez plus de six commandes de niveau supérieur.
    - Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.
    - N’ajoutez pas de boutons superflus pour augmenter la valeur de votre complément.

     > [!NOTE]
     > Les compléments qui occupent trop d’espace peuvent ne pas obtenir la [validation d’AppSource](/legal/marketplace/certification-policies).

- Pour toutes les icônes, suivez les [règles de conception d’icône](add-in-icons.md).
- Proposez une version de complément qui fonctionne aussi sur les hôtes qui ne prennent pas en charge les commandes. Un seul manifeste de complément peut fonctionner sur les hôtes tenant compte ou non des commandes (par exemple, un volet Office dans le second cas).

   *Figure 3. Complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016*

   ![Capture d’écran illustrant un complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a>Étapes suivantes

La meilleure façon de commencer à utiliser des commandes de complément consiste à consulter des [exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) sur GitHub.

Pour plus d’informations sur la spécification des commandes de complément dans votre manifeste, reportez-vous à l’article expliquant comment [créer des commandes de complément dans votre manifeste](../develop/create-addin-commands.md) et au contenu de référence sur [VersionOverrides](../reference/manifest/versionoverrides.md).
