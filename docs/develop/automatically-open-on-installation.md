---
title: Ouvrir automatiquement un volet Office lorsqu’un complément est installé
description: Découvrez comment configurer un complément Office pour qu’il s’ouvre automatiquement lorsqu’il est installé.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d6ff4b8b5b68236d435ec91b2dcbe121f211081d
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674764"
---
# <a name="automatically-open-a-task-pane-when-an-add-in-is-installed"></a>Ouvrir automatiquement un volet Office lorsqu’un complément est installé

Vous pouvez configurer le volet Office de votre complément pour qu’il soit lancé immédiatement après son installation. Cette fonctionnalité augmente l’utilisation. 

Par défaut, les compléments du volet Office qui n’incluent *aucune* [commande de complément](../design/add-in-commands.md) ouvrent le volet Office immédiatement après l’installation. Toutefois, lorsqu’un complément a une ou plusieurs commandes de complément, l’utilisateur est averti du nouveau complément, mais le complément ne se lance pas automatiquement. Ce comportement par défaut historique change de sorte que les compléments qui ont des commandes de complément se lancent automatiquement dans certaines situations. En outre, si le complément comporte plusieurs pages du volet Office, il est possible de contrôler si le complément est lancé lors de l’installation et, le cas échéant, quelle page s’ouvre dans le volet Office.

> [!NOTE]
> 
> - Cette fonctionnalité est actuellement disponible uniquement dans Office sur le Web. Nous nous efforçons d’apporter ce comportement à d’autres plateformes, mais elles présentent toujours le comportement par défaut historique décrit précédemment.
> - Cette fonctionnalité s’applique uniquement aux compléments installés par un utilisateur final, et non aux compléments déployés de manière centralisée.
> - Cette fonctionnalité ne s’applique pas aux compléments de contenu ou aux compléments de messagerie (Outlook).
> - Cette fonctionnalité s’applique uniquement aux compléments qui ont au moins une commande de complément de [type « commande du volet Office ».](../design/add-in-commands.md#types-of-add-in-commands)

## <a name="new-behavior"></a>Nouveau comportement

Le nouveau comportement est le suivant :

- Si le complément n’a qu’une seule [commande de volet office](../design/add-in-commands.md#types-of-add-in-commands), l’onglet du ruban du complément est sélectionné et le volet Office s’ouvre automatiquement lors de l’installation. Vous n’avez pas besoin de configurer quoi que ce soit.
- Si le complément a plusieurs commandes de volet office et que l’une d’elles est configurée pour être la valeur par défaut (voir [Configurer le volet Office par défaut](#configure-default-task-pane)), l’onglet du ruban du complément est sélectionné et le volet Office par défaut s’ouvre automatiquement lors de l’installation.
- Si le complément a plusieurs commandes de volet Office, mais qu’aucune n’est configurée comme valeur par défaut, l’onglet du ruban du complément est sélectionné automatiquement lors de l’installation et une légende s’affiche près de celui-ci pour avertir l’utilisateur du nouveau complément, mais aucun volet Office n’est ouvert. Il s’agit du même comportement que le comportement par défaut historique.

> [!NOTE]
> Si, pour une raison quelconque, la commande de complément qui lance le volet Office ne peut pas être sélectionnée manuellement par un utilisateur au démarrage, par exemple lorsqu’elle est [configurée pour être désactivée](../design/disable-add-in-commands.md) au démarrage, elle n’est pas ouverte automatiquement, quelle que soit la configuration. 

## <a name="configure-default-task-pane"></a>Configurer le volet Office par défaut

Pour désigner un volet Office par défaut, ajoutez un élément [TaskpaneId](/javascript/api/manifest/action#taskpaneid) comme premier enfant de l’élément **\<Action\>** et définissez sa valeur sur **Office.AutoShowTaskpaneWithDocument**. Voici un exemple.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

> [!TIP]
> Si vous souhaitez que votre complément se lance automatiquement chaque fois que l’utilisateur rouvre le document, vous devez effectuer d’autres étapes de configuration. Pour plus d’informations et des conseils sur l’utilisation de cette fonctionnalité, consultez [Ouvrir automatiquement un volet Office avec un document](automatically-open-a-task-pane-with-a-document.md). 

## <a name="see-also"></a>Voir aussi

- [Ouvrir automatiquement un volet de tâches avec un document](automatically-open-a-task-pane-with-a-document.md)
