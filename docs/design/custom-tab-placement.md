---
title: Positionner un onglet personnalisé sur le ruban
description: Découvrez comment contrôler l’endroit où un onglet personnalisé apparaît sur Office ruban et s’il a le focus par défaut.
ms.date: 02/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3c1955caaf7dc8004257307fb41f33ddfbaded4d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150143"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>Positionner un onglet personnalisé sur le ruban

Vous pouvez spécifier l’endroit où vous souhaitez que l’onglet personnalisé de votre application de Office apparaisse sur le ruban de l’application Office à l’aide de la marque dans le manifeste du module.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec l’article Concepts de base pour [les commandes de add-in.](add-in-commands.md) Si vous ne l’avez pas fait récemment, veuillez l’examiner.

> [!IMPORTANT]
>
> - La fonctionnalité et le markup du add-in décrits dans cet article sont disponibles *uniquement dans PowerPoint sur le web*.
> - Le markup décrit dans cet article fonctionne uniquement sur les plateformes qui supportent l’ensemble de conditions **requises AddinCommands 1.3**. Voir [Comportement sur les plateformes non](#behavior-on-unsupported-platforms) pris en cas de problème ci-dessous.

Spécifiez l’endroit où vous souhaitez qu’un onglet personnalisé apparaisse en identifiant l’onglet Office intégré que vous souhaitez qu’il soit à côté et en spécifiant s’il doit se trouver à gauche ou à droite de l’onglet intégré. Faites ces spécifications en incluant un [élément InsertBefore](../reference/manifest/customtab.md#insertbefore) (gauche) ou [InsertAfter](../reference/manifest/customtab.md#insertafter) (à droite) dans l’élément [CustomTab](../reference/manifest/customtab.md) du manifeste de votre add-in. (Vous ne pouvez pas avoir les deux éléments.)

Dans l’exemple suivant, l’onglet personnalisé est configuré pour apparaître juste *après* **l’onglet** Révision. Notez que la valeur de l’élément est l’ID de l’onglet `<InsertAfter>` Office intégré. 

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

- Les  `<InsertBefore>` éléments et les éléments sont  `<InsertAfter>` facultatifs. Si vous n’utilisez ni l’un ni l’autre, votre onglet personnalisé apparaîtra comme onglet le plus à droite du ruban.
- Les  `<InsertBefore>` éléments et les éléments  `<InsertAfter>` s’excluent mutuellement. Vous ne pouvez pas utiliser les deux.
- Si l’utilisateur installe plusieurs modules dont l’onglet personnalisé est configuré  au même endroit, par exemple après l’onglet Révision, l’onglet du dernier module installé se trouve à cet endroit. Les onglets des add-ins précédemment installés sont déplacés d’un endroit à l’autre. Par exemple, l’utilisateur installe les add-ins A, B et C dans cet  ordre et tous sont configurés pour insérer un onglet après l’onglet Révision, puis les onglets apparaissent dans cet ordre : **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Les utilisateurs peuvent personnaliser le ruban dans l Office’application. Par exemple, un utilisateur peut déplacer ou masquer l’onglet de votre add-in. Vous ne pouvez pas l’empêcher ou détecter qu’il s’est produit.
- Si un utilisateur déplace l’un des onglets intégrés, Office interprète les éléments et les éléments en termes d’emplacement par défaut de `<InsertBefore>` `<InsertAfter>` l’onglet *intégré.* Par exemple, si l’utilisateur déplace l’onglet Révision à l’extrémité droite du ruban, Office interprète le marques de révision dans l’exemple ci-dessus comme « placer l’onglet personnalisé à droite de l’endroit où se trouve l’onglet Révision par défaut ». ** 

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>Spécification de l’onglet qui a le focus à l’ouverture du document

Office permet toujours d’avoir le focus par défaut sur l’onglet qui se trouve immédiatement à droite de **l’onglet** Fichier. Par défaut, il s’agit de **l’onglet** Accueil. Si vous configurez votre onglet  personnalisé pour qu’il se place avant l’onglet Accueil, avec , votre onglet personnalisé aura le focus à l’ouverture `<InsertBefore>TabHome</InsertBefore>` du document.

> [!IMPORTANT]
> Donner une importance excessive à votre complément dérange et contrarie les utilisateurs et les administrateurs. Ne positionnez pas  un onglet personnalisé avant l’onglet Accueil, sauf si votre module est le principal moyen pour les utilisateurs d’interagir avec le document.

## <a name="behavior-on-unsupported-platforms"></a>Comportement sur les plateformes non pris en place

Si votre add-in est installé sur une plateforme qui ne prend pas en charge l’ensemble de conditions [requises AddinCommands 1.3,](../reference/requirement-sets/add-in-commands-requirement-sets.md)le markup décrit dans cet article est ignoré et votre onglet personnalisé apparaît comme onglet le plus à droite sur le ruban. Pour empêcher l’installation de votre add-in sur des plateformes qui ne la prisent pas en charge, ajoutez une référence à l’ensemble de conditions requises dans la `<Requirements>` section du manifeste. Pour obtenir des instructions, [voir Définir l’élément Requirements dans le manifeste.](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) Vous pouvez également concevoir votre add-in pour qu’il offre une expérience de substitution lorsque **AddinCommands 1.3** n’est pas pris en charge, comme décrit dans utiliser les vérifications à l’runtime dans votre [code JavaScript.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) Par exemple, si votre add-in contient des instructions qui supposent que l’onglet personnalisé est l’endroit où vous le souhaitez, vous pouvez avoir une autre version qui suppose que l’onglet est le plus à droite.
