---
title: Concepts basiques pour les commandes de complément
description: Découvrez l'ajout de boutons et d'éléments de menu personnalisés au ruban dans Office dans le cadre d’un complément Office.
ms.date: 05/10/2022
ms.localizationpriority: high
ms.openlocfilehash: 5d08ba9958d8c2f7002e32f726b087a15dbf27e0
ms.sourcegitcommit: fd04b41f513dbe9e623c212c1cbd877ae2285da0
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2022
ms.locfileid: "65313190"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Commandes de complément pour Excel, PowerPoint et Word

Les commandes de complément sont des éléments d’interface utilisateur qui étendent l’interface utilisateur d’Office et lancent des actions dans votre complément. Vous pouvez les utiliser pour ajouter un bouton sur le ruban ou un élément dans le menu contextuel. Lorsque les utilisateurs sélectionnent une commande de complément, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage d’une page du complément dans le volet Office. Les commandes de complément aident les utilisateurs à trouver et utiliser votre complément, ce qui favorise l’adoption et la réutilisation de votre complément, et améliore la fidélisation des clients.

> [!NOTE]
> Les catalogues SharePoint ne prennent pas en charge les commandes de complément. Vous pouvez déployer des commandes de complément via [Integrated Apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) ou [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), ou utiliser [le chargement latéral](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) pour déployer votre commande de complément à des fins de test.

> [!IMPORTANT]
> Les commandes de complément sont actuellement prises en charge dans Outlook. Pour plus d’informations, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).

*Figure 1. Complément incluant des commandes en cours d’exécution dans Excel (version de bureau)*

![Capture d’écran affichant les commandes de complément mises en évidence dans le ruban Excel.](../images/add-in-commands-1.png)

*Figure 2. Complément incluant des commandes en cours d’exécution dans Excel sur le web*

![Capture d’écran affichant des commandes de complément dans Excel sur le web.](../images/add-in-commands-2.png)

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
- ExecuteFunction - Charge une page HTML invisible, puis y exécute une fonction JavaScript. Pour afficher l’interface utilisateur au sein de votre fonction (par exemple, erreurs, avancement, entrées supplémentaires), vous pouvez utiliser l’API [displayDialog](/javascript/api/office/office.ui).  

### <a name="default-enabled-or-disabled-status"></a>État Activé ou Désactivé par défaut

Vous pouvez spécifier si la commande est activée ou désactivée lors du lancement de votre complément et modifier le paramètre par programme.

> [!NOTE]
> Cette fonctionnalité n’est pas prise en charge dans toutes les applications Office ni tous les scénarios. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](disable-add-in-commands.md).

### <a name="position-on-the-ribbon-preview"></a>Position sur le ruban (aperçu)

Vous pouvez spécifier l’emplacement où s’affiche un onglet personnalisé sur le ruban de l’application Office, par exemple, « juste à droite de l’onglet Accueil ».

> [!NOTE]
> Cette fonctionnalité n’est pas prise en charge dans toutes les applications Office ni dans tous les scénarios. Pour plus d’informations, voir [Positionner un onglet personnalisé sur le ruban](custom-tab-placement.md).

### <a name="integration-of-built-in-office-buttons"></a>Intégration des boutons Office intégrés

Vous pouvez insérer les boutons prédéfinis du ruban Office dans vos groupes personnalisés de commandes et onglets personnalisés du ruban.

> [!NOTE]
> Cette fonctionnalité n’est pas prise en charge dans toutes les applications Office ni dans tous les scénarios. Pour plus d’informations, voir [Intégrer des boutons prédéfinis Office dans les onglets personnalisés](built-in-button-integration.md).

### <a name="contextual-tabs"></a>Onglets contextuels

Vous pouvez spécifier qu’un onglet n’est visible que dans le ruban dans certains contextes, par exemple lorsque vous sélectionnez un graphique dans Excel.

> [!NOTE]
> Cette fonctionnalité n’est pas prise en charge dans toutes les applications Office ni dans tous les scénarios. Si vous souhaitez en savoir, veuillez consulter la rubrique [Créer des onglets contextuels personnalisés dans des compléments Office](contextual-tabs.md).

## <a name="supported-platforms"></a>Plateformes prises en charge

Les commandes de complément sont actuellement prises en charge sur les plateformes suivantes, à l’exception des limitations spécifiées plus haut dans les sous-sections de [Fonctionnalités de commande](#command-capabilities).

- Office on Windows (build 16.0.6769+, connecté à un abonnement Microsoft 365)
- Office 2019 ou version ultérieure sous Windows
- Office sur Mac (build 15.33+, connecté à un abonnement Microsoft 365)
- Office 2019 ou version ultérieure sur Mac
- Office sur le web

> [!NOTE]
> Pour plus d’informations sur la prise en charge dans Outlook, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).

## <a name="debug"></a>Débogage

Pour déboguer une commande de complément, vous devez l’exécuter dans Office sur le web. Pour plus de détails, voir [Débogage de compléments dans Office sur le web](../testing/debug-add-ins-in-office-online.md).

## <a name="best-practices"></a>Meilleures pratiques

Appliquez les meilleures pratiques suivantes lorsque vous développez des commandes de complément.

- Utilisez les commandes pour représenter une action spécifique avec un résultat clair et précis pour les utilisateurs. Ne combinez pas plusieurs actions dans un seul bouton.
- Proposez des actions détaillées permettant de réaliser plus efficacement des tâches courantes dans votre complément. Réduisez le nombre d’étapes nécessaires à la réalisation d’une action.
- Pour le placement de vos commandes dans le ruban d'application de l'Office :
  - Placez les commandes sur un onglet existant (Insertion, Révision, etc.) si la fonctionnalité ajoutée lui correspond. Par exemple, si votre complément permet aux utilisateurs d’insérer un élément multimédia, ajoutez un groupe à l’onglet Insertion. Notez que l’ensemble des onglets ne sont pas nécessairement disponibles dans toutes les versions d’Office. Pour plus d’informations, voir le [manifeste XML de compléments Office](../develop/add-in-manifests.md).
  - Placez les commandes sous l’onglet Accueil si la fonctionnalité ne correspond à aucun autre onglet, et si vous avez moins de six commandes de niveau supérieur. Vous pouvez également ajouter des commandes à l’onglet Accueil si votre complément doit fonctionner sur toutes les versions d’Office (par exemple, Office sur le web ou le bureau) et si un onglet n’est pas disponible dans toutes les versions (par exemple, si l’onglet Création n’existe pas dans Office sur le web).  
  - Placez des commandes dans un onglet personnalisé si vous avez plus de six commandes de niveau supérieur.
  - Nommez votre groupe en fonction du nom de votre complément. Si vous avez plusieurs groupes, nommez chaque groupe en fonction de la fonctionnalité offerte par les commandes de ce groupe.
  - N’ajoutez pas de boutons superflus pour augmenter la valeur de votre complément.
  - Ne positionnez pas un onglet personnalisé à gauche de l’onglet Accueil, ou donnez-lui le focus par défaut lorsque le document s’ouvre, sauf si votre complément est la principale façon dont les utilisateurs interagissent avec le document. Donner une importance excessive à votre complément importune et gêne les utilisateurs et les administrateurs.
  - Si votre complément est le principal mode d’interaction des utilisateurs avec le document et que vous avez un onglet de ruban personnalisé, envisagez d’intégrer dans l’onglet les boutons de fonctions d’Office dont les utilisateurs ont fréquemment besoin.
  - Si la fonctionnalité fournie avec un onglet personnalisé ne doit être disponible que dans certains contextes, utilisez les [onglets contextuels personnalisés](contextual-tabs.md). Si vous utilisez des onglets contextuels personnalisés, veillez à implémenter une [expérience de secours pour les cas d’utilisation de votre complément sur des plateformes qui ne prennent pas en charge les onglets contextuels personnalisés](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

  > [!NOTE]
  > Les compléments qui occupent trop d’espace peuvent ne pas obtenir la [validation d’AppSource](/legal/marketplace/certification-policies).

- Pour toutes les icônes, suivez les [règles de conception d’icône](add-in-icons.md).
- Fournissez une version de votre complément qui fonctionne également sur les applications Office qui ne prennent pas en charge les commandes. Un seul manifeste de complément peut fonctionner dans des applications qui prennent en charge les commandes (avec des commandes) et qui ne prennent pas en charge les commandes (en tant que volet Office).

   *Figure 3. Complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016*

   ![Capture d’écran comparant un complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016. Dans la version 2013, le volet Office doit contenir toutes les commandes, tandis que dans la version 2016, les commandes peuvent se trouver dans le ruban.](../images/office-task-pane-add-ins.png)

## <a name="next-steps"></a>Étapes suivantes

La meilleure façon de commencer à utiliser des commandes de complément consiste à consulter des [exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) sur GitHub.

Pour plus d’informations sur la spécification des commandes de complément dans votre manifeste, reportez-vous à l’article expliquant comment [créer des commandes de complément dans votre manifeste](../develop/create-addin-commands.md) et au contenu de référence sur [VersionOverrides](/javascript/api/manifest/versionoverrides).
