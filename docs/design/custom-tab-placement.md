---
title: Positionnement d’un onglet personnalisé sur le ruban
description: Découvrez comment contrôler l’emplacement où un onglet personnalisé apparaît dans le ruban Office et s’il a le focus par défaut.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 2c1e2ae66805212e78868cf7c07a0e5c14cd4025
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088171"
---
# <a name="position-a-custom-tab-on-the-ribbon-preview"></a>Positionnement d’un onglet personnalisé sur le ruban (aperçu)

Vous pouvez spécifier où l’onglet personnalisé de votre complément doit apparaître sur le ruban de l’application Office à l’aide de balises dans le manifeste du complément.

> [!NOTE]
> Cet article suppose que vous connaissez bien l’article [concepts de base pour les commandes de complément](add-in-commands.md). Vérifiez-le si vous ne l’avez pas fait récemment.

> [!IMPORTANT]
>
> - La fonctionnalité de complément et le balisage décrits dans cet article sont dans l’aperçu et sont *disponibles uniquement dans PowerPoint sur le Web*. Nous vous recommandons d’essayer le balisage uniquement dans les environnements de test et de développement. N’utilisez pas les marques de révision dans un environnement de production ou dans des documents professionnels.
> - Le balisage décrit dans cet article fonctionne uniquement sur les plateformes qui prennent en charge l’ensemble de conditions requises **AddinCommands 1,3**. Voir [comportement sur les plateformes non prises en charge](#behavior-on-unsupported-platforms) ci-dessous.

Spécifiez l’emplacement où vous souhaitez afficher un onglet personnalisé en identifiant l’onglet Office prédéfini à côté duquel vous souhaitez le placer et en spécifiant s’il doit se trouver à gauche ou à droite de l’onglet intégré. Pour ce faire, vous devez inclure un élément [InsertBefore](../reference/manifest/customtab.md#insertbefore) (Left) ou [InsertAfter](../reference/manifest/customtab.md#insertafter) (Right) dans l’élément [CustomTab](../reference/manifest/customtab.md) du manifeste de votre complément. (Vous ne pouvez pas avoir les deux éléments.)

Dans l’exemple suivant, l’onglet personnalisé est configuré pour apparaître *juste après* l’onglet **révision** . Notez que la valeur de l' `<InsertAfter>` élément est l’ID de l’onglet Office prédéfini. 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

Gardez les points suivants à l’esprit.

- Les  `<InsertBefore>`  `<InsertAfter>` éléments et sont facultatifs. Si vous ne l’utilisez pas, votre onglet personnalisé apparaîtra en tant qu’onglet le plus à droite sur le ruban.
- Les  `<InsertBefore>` éléments et s’excluent  `<InsertAfter>` mutuellement. Vous ne pouvez pas utiliser les deux.
- Si l’utilisateur installe plus d’un complément dont l’onglet personnalisé est configuré pour le même emplacement, par exemple, après l’onglet **révision** , l’onglet du complément installé le plus récemment est situé à cet endroit. Les onglets des compléments précédemment installés seront déplacés d’un endroit à un autre. Par exemple, l’utilisateur installe des compléments A, B et C dans cet ordre et tous sont configurés pour insérer un onglet après l’onglet **révision** , les onglets apparaissent dans cet ordre : **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Les utilisateurs peuvent personnaliser le ruban dans l’application Office. Par exemple, un utilisateur peut déplacer ou masquer l’onglet de votre complément. Vous ne pouvez pas empêcher cela ou ne pas détecter qu’il s’est produit.
- Si un utilisateur déplace l’un des onglets intégrés, Office interprète les `<InsertBefore>`  `<InsertAfter>` éléments et en fonction de *l’emplacement par défaut de l’onglet intégré*. Par exemple, si l’utilisateur déplace l’onglet **révision** à l’extrémité droite du ruban, Office interprète le balisage de l’exemple ci-dessus comme « placer l’onglet personnalisé juste à droite de l' *onglet **révision** par défaut*».

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>Spécification de l’onglet qui a le focus lors de l’ouverture du document

Office active toujours le focus par défaut sur l’onglet situé immédiatement à droite de l’onglet **fichier** . Par défaut, il s’agit de l’onglet **Accueil** . Si vous configurez votre onglet personnalisé de sorte qu’il se trouve avant l’onglet **Accueil** , avec `<InsertBefore>TabHome</InsertBefore>` , l’onglet personnalisé est activé lorsque le document s’ouvre.

> [!IMPORTANT]
> En donnant une importance excessive à votre complément, vous désactivez les utilisateurs et les administrateurs. Ne positionnez pas un onglet personnalisé avant l’onglet **Accueil** , sauf si votre complément est le principal moyen pour les utilisateurs d’interagir avec le document.

## <a name="behavior-on-unsupported-platforms"></a>Comportement sur les plateformes non prises en charge

Si votre complément est installé sur une plateforme qui ne prend pas en charge l' [ensemble de conditions de AddinCommands 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), le balisage décrit dans cet article est ignoré et votre onglet personnalisé apparaîtra en tant qu’onglet le plus à droite sur le ruban. Pour empêcher l’installation de votre complément sur des plateformes qui ne prennent pas en charge le balisage, ajoutez une référence à l’ensemble de conditions requises dans la `<Requirements>` section du manifeste. Pour obtenir des instructions, voir [définir l’élément Requirements dans le manifeste](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Vous pouvez également concevoir votre complément de manière à ce qu’il ait une expérience secondaire lorsque **AddinCommands 1,3** n’est pas pris en charge, comme décrit dans [la rubrique use Runtime Checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). Par exemple, si votre complément contient des instructions qui partent de l’emplacement où vous le souhaitez, vous pouvez avoir une autre version qui suppose que l’onglet est le plus à droite.
