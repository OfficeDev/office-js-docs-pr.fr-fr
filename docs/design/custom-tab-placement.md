---
title: Positionner un onglet personnalisé sur le ruban
description: Découvrez comment contrôler l’emplacement d’un onglet personnalisé dans le ruban Office et s’il a le focus par défaut.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 42445898623e082c3c85e756625307dc5a237c28
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659814"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>Positionner un onglet personnalisé sur le ruban

Vous pouvez spécifier l’emplacement où vous souhaitez que l’onglet personnalisé de votre complément apparaisse sur le ruban de l’application Office à l’aide du balisage dans le manifeste du complément.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les concepts de base de l’article [pour les commandes de complément](add-in-commands.md). Veuillez la consulter si vous ne l’avez pas fait récemment.

> [!IMPORTANT]
>
> - La fonctionnalité de complément et le balisage décrits dans cet article *sont disponibles uniquement dans PowerPoint sur le web*.
> - Le balisage décrit dans cet article fonctionne uniquement sur les plateformes qui prennent en charge l’ensemble de conditions requises **AddinCommands 1.3**. Consultez [comportement sur les plateformes non prises en charge ci-dessous](#behavior-on-unsupported-platforms) .

Spécifiez l’endroit où vous souhaitez qu’un onglet personnalisé s’affiche en identifiant l’onglet Office intégré auquel vous voulez qu’il se trouve à côté et en spécifiant s’il doit se trouver sur le côté gauche ou droit de l’onglet intégré. Effectuez ces spécifications en incluant un élément [InsertBefore](/javascript/api/manifest/customtab#insertbefore) (à gauche) ou [InsertAfter](/javascript/api/manifest/customtab#insertafter) (à droite) dans l’élément [CustomTab](/javascript/api/manifest/customtab) du manifeste de votre complément. (Vous ne pouvez pas avoir les deux éléments.)

Dans l’exemple suivant, l’onglet personnalisé est configuré pour apparaître *juste après* l’onglet **Révision** . Notez que la valeur de l’élément **\<InsertAfter\>** est l’ID de l’onglet Office intégré. 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

Gardez à l’esprit les points suivants.

- Les éléments et **\<InsertAfter\>** les **\<InsertBefore\>** éléments sont facultatifs. Si vous n’utilisez ni l’un ni l’autre, votre onglet personnalisé apparaîtra sous la forme de l’onglet le plus à droite du ruban.
- Les **\<InsertBefore\>** éléments et **\<InsertAfter\>** les éléments s’excluent mutuellement. Vous ne pouvez pas utiliser les deux.
- Si l’utilisateur installe plusieurs compléments dont l’onglet personnalisé est configuré pour le même emplacement, par exemple après l’onglet **Révision** , l’onglet du complément le plus récemment installé se trouve à cet emplacement. Les onglets des compléments précédemment installés seront déplacés sur un seul emplacement. Par exemple, l’utilisateur installe les compléments A, B et C dans cet ordre et tous sont configurés pour insérer un onglet après l’onglet **Révision** , puis les onglets s’affichent dans cet ordre : **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Les utilisateurs peuvent personnaliser le ruban dans l’application Office. Par exemple, un utilisateur peut déplacer ou masquer l’onglet de votre complément. Vous ne pouvez pas empêcher cela ou détecter qu’il s’est produit.
- Si un utilisateur déplace l’un des onglets intégrés, Office interprète les éléments et **\<InsertAfter\>** les **\<InsertBefore\>** éléments en termes *d’emplacement par défaut de l’onglet intégré*. Par exemple, si l’utilisateur déplace l’onglet **Révision** à l’extrémité droite du ruban, Office interprète le balisage dans l’exemple précédent comme « placez l’onglet personnalisé juste à droite de *l’emplacement par défaut de l’onglet **Révision*** ».

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>Spécifier l’onglet qui a le focus lorsque le document s’ouvre

Office accorde toujours le focus par défaut à l’onglet qui se trouve immédiatement à droite de l’onglet **Fichier** . Par défaut, il s’agit de l’onglet **Accueil** . Si vous configurez votre onglet personnalisé pour qu’il soit avant l’onglet **Accueil** , avec `<InsertBefore>TabHome</InsertBefore>`, votre onglet personnalisé aura le focus lorsque le document s’ouvre.

> [!IMPORTANT]
> Donner une importance excessive à votre complément dérange et contrarie les utilisateurs et les administrateurs. Ne positionnez pas un onglet personnalisé avant l’onglet **Accueil** , sauf si votre complément est la principale façon dont les utilisateurs interagissent avec le document.

## <a name="behavior-on-unsupported-platforms"></a>Comportement sur les plateformes non prises en charge

Si votre complément est installé sur une plateforme qui ne prend pas en charge [l’ensemble de conditions requises AddinCommands 1.3, le balisage](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) décrit dans cet article est ignoré et votre onglet personnalisé apparaît comme l’onglet le plus à droite du ruban. Pour empêcher l’installation de votre complément sur des plateformes qui ne prennent pas en charge le balisage, ajoutez une référence à l’ensemble de conditions requises dans la **\<Requirements\>** section du manifeste. Pour obtenir des instructions, consultez [Spécifier les versions et plateformes Office qui peuvent héberger votre complément](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Vous pouvez également concevoir votre complément pour avoir une autre expérience lorsque **AddinCommands 1.3** n’est pas pris en charge, comme décrit dans [La Conception pour les expériences alternatives](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Par exemple, si votre complément contient des instructions qui supposent que l’onglet personnalisé est l’emplacement souhaité, vous pouvez avoir une autre version qui suppose que l’onglet est le plus à droite.
