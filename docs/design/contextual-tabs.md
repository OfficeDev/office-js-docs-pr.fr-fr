---
title: Créez des onglets contextuels personnalisés Office add-ins
description: Découvrez comment ajouter des onglets contextuels personnalisés à Office add-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d03ac2c01c03353f3e2d1b54ba20616d7b42d93f
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555205"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="de8c7-103">Créez des onglets contextuels personnalisés Office add-ins</span><span class="sxs-lookup"><span data-stu-id="de8c7-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="de8c7-104">Un onglet contextuel est un contrôle d’onglet caché dans le ruban Office qui s’affiche dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office’affichage.</span><span class="sxs-lookup"><span data-stu-id="de8c7-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="de8c7-105">Par exemple, **l’onglet Conception** de table qui apparaît sur Excel ruban lorsqu’une table est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="de8c7-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="de8c7-106">Vous pouvez inclure des onglets contextuels personnalisés dans votre Office Add-in et spécifier quand ils sont visibles ou cachés, en créant des gestionnaires d’événements qui modifient la visibilité.</span><span class="sxs-lookup"><span data-stu-id="de8c7-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="de8c7-107">(Toutefois, les onglets contextuels personnalisés ne répondent pas aux modifications de mise au point.)</span><span class="sxs-lookup"><span data-stu-id="de8c7-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="de8c7-108">Cet article suppose que vous connaissez la documentation décrite ci-après.</span><span class="sxs-lookup"><span data-stu-id="de8c7-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="de8c7-109">Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).</span><span class="sxs-lookup"><span data-stu-id="de8c7-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="de8c7-110">Concepts basiques pour les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="de8c7-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="de8c7-111">Les onglets contextuels personnalisés ne sont actuellement pris en charge Excel et uniquement sur ces plates-formes et builds :</span><span class="sxs-lookup"><span data-stu-id="de8c7-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="de8c7-112">Excel sur Windows (Microsoft 365 abonnement uniquement): Version 2102 (Build 13801.20294) ou plus tard.</span><span class="sxs-lookup"><span data-stu-id="de8c7-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="de8c7-113">Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="de8c7-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="de8c7-114">Les onglets contextuels personnalisés ne fonctionnent que sur les plates-formes qui supporte les ensembles d’exigences suivants.</span><span class="sxs-lookup"><span data-stu-id="de8c7-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="de8c7-115">Pour en savoir plus sur les ensembles d’exigences et la façon de travailler avec eux, [consultez spécifier Office applications et les exigences de l’API](../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="de8c7-116">RibbonApi 1.2 RubanApi 1.2</span><span class="sxs-lookup"><span data-stu-id="de8c7-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="de8c7-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="de8c7-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="de8c7-118">Vous pouvez utiliser les vérifications de temps d’exécution dans votre code pour vérifier si l’hôte de l’utilisateur et la combinaison de plate-forme prend en charge ces ensembles d’exigences tels [que décrits dans spécifier les applications Office et les exigences de l’API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="de8c7-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="de8c7-119">(La technique de spécifier les ensembles d’exigences dans le manifeste, qui est également décrit dans cet article, ne fonctionne pas actuellement pour RibbonApi 1.2.) Alternativement, vous pouvez [implémenter une expérience d’interface utilisateur alternative lorsque les onglets contextuels personnalisés ne sont pas pris en charge](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span><span class="sxs-lookup"><span data-stu-id="de8c7-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="de8c7-120">Comportement des onglets contextuels personnalisés</span><span class="sxs-lookup"><span data-stu-id="de8c7-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="de8c7-121">L’expérience utilisateur des onglets contextuels personnalisés suit le modèle des onglets Office intégrés intégrés.</span><span class="sxs-lookup"><span data-stu-id="de8c7-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="de8c7-122">Voici les principes de base pour les onglets contextuels personnalisés de placement :</span><span class="sxs-lookup"><span data-stu-id="de8c7-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="de8c7-123">Lorsqu’un onglet contextuel personnalisé est visible, il apparaît à l’extrémité droite du ruban.</span><span class="sxs-lookup"><span data-stu-id="de8c7-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="de8c7-124">Si un ou plusieurs onglets contextuels intégrés et un ou plusieurs onglets contextuels personnalisés provenant d’add-ins sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.</span><span class="sxs-lookup"><span data-stu-id="de8c7-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="de8c7-125">Si votre module a plus d’un onglet contextuel et qu’il y a des contextes dans lesquels plus d’un est visible, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre module.</span><span class="sxs-lookup"><span data-stu-id="de8c7-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="de8c7-126">(La direction est dans la même direction que la langue Office; c’est-à-dire de gauche à droite dans les langues de gauche à droite, mais de droite à gauche dans les langues de droite à gauche.) Voir [Définir les groupes et les contrôles qui apparaissent sur l’onglet pour plus](#define-the-groups-and-controls-that-appear-on-the-tab) de détails sur la façon dont vous les définissez.</span><span class="sxs-lookup"><span data-stu-id="de8c7-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="de8c7-127">Si plus d’un add-in a un onglet contextuel qui est visible dans un contexte spécifique, alors ils apparaissent dans l’ordre dans lequel les add-ins ont été lancés.</span><span class="sxs-lookup"><span data-stu-id="de8c7-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="de8c7-128">Les *onglets* contextuels personnalisés, contrairement aux onglets de base personnalisés, ne sont pas ajoutés en permanence Office ruban de l’application.</span><span class="sxs-lookup"><span data-stu-id="de8c7-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="de8c7-129">Ils ne sont présents que dans Office documents sur lesquels votre module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="de8c7-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="de8c7-130">Étapes majeures pour inclure un onglet contextuel dans un module</span><span class="sxs-lookup"><span data-stu-id="de8c7-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="de8c7-131">Voici les principales étapes pour inclure un onglet contextuel personnalisé dans un module :</span><span class="sxs-lookup"><span data-stu-id="de8c7-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="de8c7-132">Configurez l’add-in pour utiliser un temps d’exécution partagé.</span><span class="sxs-lookup"><span data-stu-id="de8c7-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="de8c7-133">Définissez l’onglet et les groupes et contrôles qui y apparaissent.</span><span class="sxs-lookup"><span data-stu-id="de8c7-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="de8c7-134">Enregistrez l’onglet contextuel avec Office.</span><span class="sxs-lookup"><span data-stu-id="de8c7-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="de8c7-135">Spécifiez les circonstances dans lesquelles l’onglet sera visible.</span><span class="sxs-lookup"><span data-stu-id="de8c7-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="de8c7-136">Configurez l’add-in pour utiliser un temps d’exécution partagé</span><span class="sxs-lookup"><span data-stu-id="de8c7-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="de8c7-137">L’ajout d’onglets contextuels personnalisés nécessite votre module d’utilisation pour utiliser le temps d’exécution partagé.</span><span class="sxs-lookup"><span data-stu-id="de8c7-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="de8c7-138">Pour plus d’informations, [consultez Configurer un module d’accès pour utiliser un temps d’exécution partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="de8c7-139">Définir les groupes et les contrôles qui apparaissent sur l’onglet</span><span class="sxs-lookup"><span data-stu-id="de8c7-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="de8c7-140">Contrairement aux onglets de base personnalisés, qui sont définis avec XML dans le manifeste, les onglets contextuels personnalisés sont définis à l’exécution avec un blob JSON.</span><span class="sxs-lookup"><span data-stu-id="de8c7-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="de8c7-141">Votre code analyse le blob dans un objet JavaScript, puis passe l’objet [à la méthode Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="de8c7-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="de8c7-142">Les onglets contextuels personnalisés ne sont présents que dans les documents sur lesquels votre module est actuellement en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="de8c7-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="de8c7-143">Ceci est différent des onglets de base personnalisés qui sont ajoutés au ruban d’application Office lorsque l’add-in est installé et restent présents lorsqu’un autre document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="de8c7-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="de8c7-144">En outre, la `requestCreateControls` méthode ne peut être utilisée qu’une seule fois dans une session de votre module d’ajout.</span><span class="sxs-lookup"><span data-stu-id="de8c7-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="de8c7-145">Si elle est appelée à nouveau, une erreur est lancée.</span><span class="sxs-lookup"><span data-stu-id="de8c7-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="de8c7-146">La structure des propriétés et des sous-propriétés du blob JSON (et des noms clés) est à peu près parallèle à la structure de [l’élément CustomTab](../reference/manifest/customtab.md) et de ses éléments descendants dans le manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="de8c7-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="de8c7-147">Nous allons construire un exemple d’onglets contextuels JSON blob étape par étape.</span><span class="sxs-lookup"><span data-stu-id="de8c7-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="de8c7-148">Le schéma complet de l’onglet contextuel JSON est [ àdynamic-ribbon.schema.jssur](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="de8c7-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="de8c7-149">Si vous travaillez dans Visual Studio Code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON.</span><span class="sxs-lookup"><span data-stu-id="de8c7-149">If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="de8c7-150">Pour plus d’informations, [voir Édition JSON avec Visual Studio Code - Schémas et paramètres JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="de8c7-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="de8c7-151">Commencez par créer une chaîne JSON avec deux propriétés de tableau nommées `actions` et `tabs` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="de8c7-152">Le `actions` tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des contrôles sur l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 20*.</span><span class="sxs-lookup"><span data-stu-id="de8c7-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="de8c7-153">Cet exemple simple d’onglet contextuel n’aura qu’un seul bouton et, par conséquent, une seule action.</span><span class="sxs-lookup"><span data-stu-id="de8c7-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="de8c7-154">Ajoutez ce qui suit en tant que seul membre du `actions` tableau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="de8c7-155">A propos de ce balisage, notez:</span><span class="sxs-lookup"><span data-stu-id="de8c7-155">About this markup, note:</span></span>

    - <span data-ttu-id="de8c7-156">Les `id` propriétés et les propriétés sont `type` obligatoires.</span><span class="sxs-lookup"><span data-stu-id="de8c7-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="de8c7-157">La valeur de `type` peut être soit « ExecuteFunction » ou « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="de8c7-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="de8c7-158">La `functionName` propriété n’est utilisée que lorsque la valeur `type` de est `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="de8c7-159">C’est le nom d’une fonction définie dans le Fichier de fonction.</span><span class="sxs-lookup"><span data-stu-id="de8c7-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="de8c7-160">Pour plus d’informations sur le Fichier de fonction, [consultez les concepts de base pour les commandes complémentaires](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="de8c7-161">Dans une étape ultérieure, vous cartographiez cette action à un bouton sur l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="de8c7-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="de8c7-162">Ajoutez ce qui suit en tant que seul membre du `tabs` tableau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="de8c7-163">A propos de ce balisage, notez:</span><span class="sxs-lookup"><span data-stu-id="de8c7-163">About this markup, note:</span></span>

    - <span data-ttu-id="de8c7-164">La propriété `id` est requise.</span><span class="sxs-lookup"><span data-stu-id="de8c7-164">The `id` property is required.</span></span> <span data-ttu-id="de8c7-165">Utilisez un bref id descriptif unique parmi tous les onglets contextuels de votre module.</span><span class="sxs-lookup"><span data-stu-id="de8c7-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="de8c7-166">La propriété `label` est requise.</span><span class="sxs-lookup"><span data-stu-id="de8c7-166">The `label` property is required.</span></span> <span data-ttu-id="de8c7-167">Il s’agit d’une chaîne conviviale pour servir d’étiquette de l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="de8c7-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="de8c7-168">La propriété `groups` est requise.</span><span class="sxs-lookup"><span data-stu-id="de8c7-168">The `groups` property is required.</span></span> <span data-ttu-id="de8c7-169">Il définit les groupes de contrôles qui apparaîtront sur l’onglet. Il doit avoir au moins un membre *et pas plus de 20*.</span><span class="sxs-lookup"><span data-stu-id="de8c7-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="de8c7-170">(Il ya aussi des limites sur le nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui limitera également le nombre de groupes que vous avez.</span><span class="sxs-lookup"><span data-stu-id="de8c7-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="de8c7-171">Voir l’étape suivante pour plus d’informations.)</span><span class="sxs-lookup"><span data-stu-id="de8c7-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="de8c7-172">L’objet onglet peut également avoir une propriété `visible` optionnelle qui spécifie si l’onglet est visible immédiatement lorsque l’add-in démarre.</span><span class="sxs-lookup"><span data-stu-id="de8c7-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="de8c7-173">Étant donné que les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche leur visibilité (comme l’utilisateur sélectionnant une entité d’un certain type dans le document), la propriété ne `visible` se présente `false` pas lorsqu’elle n’est pas présente.</span><span class="sxs-lookup"><span data-stu-id="de8c7-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="de8c7-174">Dans une section ultérieure, nous montrons comment définir la propriété `true` en réponse à un événement.</span><span class="sxs-lookup"><span data-stu-id="de8c7-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="de8c7-175">Dans l’exemple continu simple, l’onglet contextuel n’a qu’un seul groupe.</span><span class="sxs-lookup"><span data-stu-id="de8c7-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="de8c7-176">Ajoutez ce qui suit en tant que seul membre du `groups` tableau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="de8c7-177">A propos de ce balisage, notez:</span><span class="sxs-lookup"><span data-stu-id="de8c7-177">About this markup, note:</span></span>

    - <span data-ttu-id="de8c7-178">Toutes les propriétés sont requises.</span><span class="sxs-lookup"><span data-stu-id="de8c7-178">All the properties are required.</span></span>
    - <span data-ttu-id="de8c7-179">La `id` propriété doit être unique parmi tous les groupes de l’onglet. Utilisez une pièce d’identité brève et descriptive.</span><span class="sxs-lookup"><span data-stu-id="de8c7-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="de8c7-180">Il `label` s’agit d’une chaîne conviviale pour servir d’étiquette du groupe.</span><span class="sxs-lookup"><span data-stu-id="de8c7-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="de8c7-181">La `icon` valeur de la propriété est un tableau d’objets qui spécifient les icônes que le groupe aura sur le ruban en fonction de la taille du ruban et de la fenêtre d’application Office’application.</span><span class="sxs-lookup"><span data-stu-id="de8c7-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="de8c7-182">La `controls` valeur de la propriété est un éventail d’objets qui spécifient les boutons et les menus du groupe.</span><span class="sxs-lookup"><span data-stu-id="de8c7-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="de8c7-183">Il doit y en avoir au moins un.</span><span class="sxs-lookup"><span data-stu-id="de8c7-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="de8c7-184">*Le nombre total de contrôles sur l’ensemble de l’onglet ne peut pas être supérieur à 20.*</span><span class="sxs-lookup"><span data-stu-id="de8c7-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="de8c7-185">Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun, et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.</span><span class="sxs-lookup"><span data-stu-id="de8c7-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. <span data-ttu-id="de8c7-186">Chaque groupe doit avoir une icône d’au moins deux tailles, 32x32 px et 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="de8c7-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="de8c7-187">En option, vous pouvez également avoir des icônes de tailles 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, et 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="de8c7-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="de8c7-188">Office quelle icône utiliser en fonction de la taille du ruban et de la fenêtre d Office’application.</span><span class="sxs-lookup"><span data-stu-id="de8c7-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="de8c7-189">Ajoutez les objets suivants au tableau d’icônes.</span><span class="sxs-lookup"><span data-stu-id="de8c7-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="de8c7-190">(Si la taille de la fenêtre et du ruban est suffisamment grande pour qu’au moins une des *commandes* du groupe apparaisse, aucune icône de groupe n’apparaît.</span><span class="sxs-lookup"><span data-stu-id="de8c7-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="de8c7-191">Par exemple, regardez le groupe **Styles sur** le ruban Word lorsque vous rétrécissez et élargissez la fenêtre Word.) A propos de ce balisage, notez:</span><span class="sxs-lookup"><span data-stu-id="de8c7-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="de8c7-192">Les deux propriétés sont requises.</span><span class="sxs-lookup"><span data-stu-id="de8c7-192">Both the properties are required.</span></span>
    - <span data-ttu-id="de8c7-193">`size`L’unité de mesure de propriété est pixels.</span><span class="sxs-lookup"><span data-stu-id="de8c7-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="de8c7-194">Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.</span><span class="sxs-lookup"><span data-stu-id="de8c7-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="de8c7-195">La `sourceLocation` propriété spécifie l’URL complète de l’icône.</span><span class="sxs-lookup"><span data-stu-id="de8c7-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="de8c7-196">Tout comme vous devez généralement modifier les URL dans le manifeste de l’add-in lorsque vous passez du développement à la production (comme changer le domaine de localhost à contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.</span><span class="sxs-lookup"><span data-stu-id="de8c7-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. <span data-ttu-id="de8c7-197">Dans notre exemple continu simple, le groupe n’a qu’un seul bouton.</span><span class="sxs-lookup"><span data-stu-id="de8c7-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="de8c7-198">Ajoutez l’objet suivant comme seul membre du `controls` tableau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="de8c7-199">A propos de ce balisage, notez:</span><span class="sxs-lookup"><span data-stu-id="de8c7-199">About this markup, note:</span></span>

    - <span data-ttu-id="de8c7-200">Toutes les propriétés, sauf `enabled` , sont nécessaires.</span><span class="sxs-lookup"><span data-stu-id="de8c7-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="de8c7-201">`type` spécifie le type de contrôle.</span><span class="sxs-lookup"><span data-stu-id="de8c7-201">`type` specifies the type of control.</span></span> <span data-ttu-id="de8c7-202">Les valeurs peuvent être « Bouton », « Menu » ou « MobileButton ».</span><span class="sxs-lookup"><span data-stu-id="de8c7-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="de8c7-203">`id` peut être jusqu’à 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="de8c7-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="de8c7-204">`actionId` doit être l’ID d’une action définie dans le `actions` tableau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="de8c7-205">(Voir l’étape 1 de cette section.)</span><span class="sxs-lookup"><span data-stu-id="de8c7-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="de8c7-206">`label` est une chaîne conviviale pour servir d’étiquette du bouton.</span><span class="sxs-lookup"><span data-stu-id="de8c7-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="de8c7-207">`superTip` représente une forme riche de pointe d’outil.</span><span class="sxs-lookup"><span data-stu-id="de8c7-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="de8c7-208">Les propriétés `title` et les propriétés sont `description` requises.</span><span class="sxs-lookup"><span data-stu-id="de8c7-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="de8c7-209">`icon` spécifie les icônes pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="de8c7-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="de8c7-210">Les remarques précédentes sur l’icône du groupe s’appliquent ici aussi.</span><span class="sxs-lookup"><span data-stu-id="de8c7-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="de8c7-211">`enabled` (facultatif) précise si le bouton est activé lorsque l’onglet contextuel apparaît démarre.</span><span class="sxs-lookup"><span data-stu-id="de8c7-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="de8c7-212">La valeur par défaut si elle n’est pas présente est `true` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-212">The default if not present is `true`.</span></span> 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
<span data-ttu-id="de8c7-213">Voici l’exemple complet du blob JSON :</span><span class="sxs-lookup"><span data-stu-id="de8c7-213">The following is the complete example of the JSON blob:</span></span>

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="de8c7-214">Enregistrez l’onglet contextuel Office avec requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="de8c7-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="de8c7-215">L’onglet contextuel est enregistré Office en [appelant la méthode Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="de8c7-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="de8c7-216">Cela se fait généralement soit dans la fonction qui est assignée `Office.initialize` à ou avec la `Office.onReady` méthode.</span><span class="sxs-lookup"><span data-stu-id="de8c7-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="de8c7-217">Pour en savoir plus sur ces méthodes et parasinant l’add-in, [voir Initialiser Office add-in](../develop/initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="de8c7-218">Vous pouvez toutefois appeler la méthode à tout moment après l’initialisation.</span><span class="sxs-lookup"><span data-stu-id="de8c7-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="de8c7-219">La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un add-in.</span><span class="sxs-lookup"><span data-stu-id="de8c7-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="de8c7-220">Une erreur est lancée si elle est appelée à nouveau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="de8c7-221">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="de8c7-221">The following is an example.</span></span> <span data-ttu-id="de8c7-222">Notez que la chaîne JSON doit être convertie en objet JavaScript avec la méthode `JSON.parse` avant qu’elle puisse être transmise à une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="de8c7-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="de8c7-223">Spécifiez les contextes lorsque l’onglet sera visible avec requestUpdate</span><span class="sxs-lookup"><span data-stu-id="de8c7-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="de8c7-224">En règle générale, un onglet contextuel personnalisé doit apparaître lorsqu’un événement initié par l’utilisateur modifie le contexte d’ajout.</span><span class="sxs-lookup"><span data-stu-id="de8c7-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="de8c7-225">Considérez un scénario dans lequel l’onglet doit être visible lorsque, et seulement quand, un graphique (sur la feuille de travail par défaut d’un Excel de travail) est activé.</span><span class="sxs-lookup"><span data-stu-id="de8c7-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="de8c7-226">Commencez par affecter des gestionnaires.</span><span class="sxs-lookup"><span data-stu-id="de8c7-226">Begin by assigning handlers.</span></span> <span data-ttu-id="de8c7-227">Cela est généralement fait dans la `Office.onReady` méthode comme dans l’exemple suivant qui attribue les gestionnaires (créés dans une étape ultérieure) aux événements et aux événements de tous les `onActivated` graphiques de la feuille de `onDeactivated` travail.</span><span class="sxs-lookup"><span data-stu-id="de8c7-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

<span data-ttu-id="de8c7-228">Ensuite, définissez les gestionnaires.</span><span class="sxs-lookup"><span data-stu-id="de8c7-228">Next, define the handlers.</span></span> <span data-ttu-id="de8c7-229">Ce qui suit est un exemple simple d’un `showDataTab` , mais voir Manipulation de [l’erreur HostRestartNeeded plus](#handle-the-hostrestartneeded-error) tard dans cet article pour une version plus robuste de la fonction.</span><span class="sxs-lookup"><span data-stu-id="de8c7-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="de8c7-230">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="de8c7-230">About this code, note:</span></span>

- <span data-ttu-id="de8c7-231">Office effectue un contrôle lorsqu’il met à jour l’état du ruban.</span><span class="sxs-lookup"><span data-stu-id="de8c7-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="de8c7-232">La [méthode Office.ribbon.requestUpdate fait](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) la queue pour une demande de mise à jour.</span><span class="sxs-lookup"><span data-stu-id="de8c7-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="de8c7-233">La méthode résoudra `Promise` l’objet dès qu’il aura mis la demande en file d’attente, et non lorsque le ruban sera mis à jour.</span><span class="sxs-lookup"><span data-stu-id="de8c7-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="de8c7-234">Le paramètre de `requestUpdate` la méthode est un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID exactement *comme spécifié dans le JSON* et (2) spécifie la visibilité de l’onglet.</span><span class="sxs-lookup"><span data-stu-id="de8c7-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="de8c7-235">Si vous avez plus d’un onglet contextuel personnalisé qui devrait être visible dans le même contexte, il vous suffit d’ajouter des objets d’onglet supplémentaires au `tabs` tableau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

<span data-ttu-id="de8c7-236">Le gestionnaire pour cacher l’onglet est presque identique, sauf qu’il définit la `visible` propriété de nouveau à `false` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="de8c7-237">La Office javascript fournit également plusieurs interfaces (types) pour faciliter la construction de `RibbonUpdateData` l’objet.</span><span class="sxs-lookup"><span data-stu-id="de8c7-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="de8c7-238">Ce qui suit est `showDataTab` la fonction dans TypeScript et il fait usage de ces types.</span><span class="sxs-lookup"><span data-stu-id="de8c7-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="de8c7-239">Visibilité de l’onglet Basculement et état activé d’un bouton en même temps</span><span class="sxs-lookup"><span data-stu-id="de8c7-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="de8c7-240">La `requestUpdate` méthode est également utilisée pour basculer l’état activé ou désactivé d’un bouton personnalisé sur un onglet contextuel personnalisé ou un onglet de base personnalisé. Pour plus de détails à ce [sujet, consultez activer et désactiver les commandes add-in](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="de8c7-241">Il peut y avoir des scénarios dans lesquels vous souhaitez modifier à la fois la visibilité d’un onglet et l’état activé d’un bouton en même temps.</span><span class="sxs-lookup"><span data-stu-id="de8c7-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="de8c7-242">Vous pouvez le faire avec un seul appel de `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="de8c7-243">Ce qui suit est un exemple dans lequel un bouton sur un onglet de base est activé en même temps qu’un onglet contextuel est rendu visible.</span><span class="sxs-lookup"><span data-stu-id="de8c7-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

<span data-ttu-id="de8c7-244">Dans l’exemple suivant, le bouton activé est sur le même onglet contextuel qui est rendu visible.</span><span class="sxs-lookup"><span data-stu-id="de8c7-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="localizing-the-json-blob"></a><span data-ttu-id="de8c7-245">Localisation du blob JSON</span><span class="sxs-lookup"><span data-stu-id="de8c7-245">Localizing the JSON blob</span></span>

<span data-ttu-id="de8c7-246">Le blob JSON qui est transmis n’est `requestCreateControls` pas localisé de la même manière que le balisage manifeste pour les onglets de base personnalisés est localisé (qui est décrit à la localisation de contrôle à partir du [manifeste).](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="de8c7-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="de8c7-247">Au lieu de cela, la localisation doit se produire à l’heure d’exécution en utilisant des blobs JSON distincts pour chaque lieu.</span><span class="sxs-lookup"><span data-stu-id="de8c7-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="de8c7-248">Nous vous suggérons `switch` d’utiliser une instruction qui [teste Office.context.display Propriété](/javascript/api/office/office.context#displayLanguage) de la langue.</span><span class="sxs-lookup"><span data-stu-id="de8c7-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="de8c7-249">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="de8c7-249">The following is an example:</span></span>

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

<span data-ttu-id="de8c7-250">Ensuite, votre code appelle la fonction pour obtenir le blob localisé qui est transmis `requestCreateControls` à , comme dans l’exemple suivant:</span><span class="sxs-lookup"><span data-stu-id="de8c7-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="de8c7-251">Meilleures pratiques pour les onglets contextuels personnalisés</span><span class="sxs-lookup"><span data-stu-id="de8c7-251">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="de8c7-252">Implémentez une expérience d’interface utilisateur alternative lorsque les onglets contextuels personnalisés ne sont pas pris en charge</span><span class="sxs-lookup"><span data-stu-id="de8c7-252">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="de8c7-253">Certaines combinaisons de plate-forme, Office’application et Office construire ne supportent pas `requestCreateControls` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-253">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="de8c7-254">Votre module d’ajout doit être conçu pour offrir une expérience alternative aux utilisateurs qui font fonctionner l’add-in sur l’une de ces combinaisons.</span><span class="sxs-lookup"><span data-stu-id="de8c7-254">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="de8c7-255">Les sections suivantes décrivent deux façons d’offrir une expérience de repli.</span><span class="sxs-lookup"><span data-stu-id="de8c7-255">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="de8c7-256">Utilisez des onglets ou des commandes non textuels</span><span class="sxs-lookup"><span data-stu-id="de8c7-256">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="de8c7-257">Il ya un élément manifeste, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), qui est conçu pour créer une expérience de secours dans un add-in qui implémente des onglets contextuels personnalisés lorsque l’add-in est en cours d’exécution sur une application ou une plate-forme qui ne prend pas en charge les onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="de8c7-257">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="de8c7-258">La stratégie la plus simple pour l’utilisation de cet élément est que vous définissez dans le manifeste un ou plusieurs onglets de base personnalisés *(c’est-à-dire* des onglets personnalisés non textuels) qui dupliquent les personnalisations de ruban des onglets contextuels personnalisés de votre module.</span><span class="sxs-lookup"><span data-stu-id="de8c7-258">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="de8c7-259">Mais vous ajoutez `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` comme premier élément enfant du [CustomTab](../reference/manifest/customtab.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-259">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="de8c7-260">L’effet de cette chose est le suivant :</span><span class="sxs-lookup"><span data-stu-id="de8c7-260">The effect of doing so is the following:</span></span>

- <span data-ttu-id="de8c7-261">Si l’add-in s’exécute sur une application et une plate-forme qui supporte les onglets contextuels personnalisés, alors l’onglet de base personnalisé n’apparaîtra pas sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="de8c7-261">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="de8c7-262">Au lieu de cela, l’onglet contextuel personnalisé sera créé lorsque l’add-in appelle la `requestCreateControls` méthode.</span><span class="sxs-lookup"><span data-stu-id="de8c7-262">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="de8c7-263">Si *l’add-in s’exécute* sur une application ou une plate-forme qui ne prend pas en `requestCreateControls` charge, alors l’onglet de base personnalisé apparaît sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="de8c7-263">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="de8c7-264">Ce qui suit est un exemple de cette stratégie simple.</span><span class="sxs-lookup"><span data-stu-id="de8c7-264">The following is an example of this simple strategy.</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="de8c7-265">Cette stratégie simple utilise un onglet de base personnalisé qui reflète un onglet contextuel personnalisé avec ses groupes d’enfants et ses contrôles, mais vous pouvez utiliser une stratégie plus complexe.</span><span class="sxs-lookup"><span data-stu-id="de8c7-265">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="de8c7-266">`<OverriddenByRibbonApi>`L’élément peut également être ajouté en tant que (premier) élément enfant aux éléments [groupe](../reference/manifest/group.md) [et contrôle](../reference/manifest/control.md) [(type de bouton et](../reference/manifest/control.md#button-control) type de [menu)](../reference/manifest/control.md#menu-dropdown-button-controls)et éléments de `<Item>` menu.</span><span class="sxs-lookup"><span data-stu-id="de8c7-266">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="de8c7-267">Ce fait vous permet de distribuer les groupes et les contrôles qui apparaîtraient autrement sur l’onglet contextuel entre différents groupes, boutons et menus dans divers onglets de base personnalisés.</span><span class="sxs-lookup"><span data-stu-id="de8c7-267">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="de8c7-268">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="de8c7-268">The following is an example.</span></span> <span data-ttu-id="de8c7-269">Notez que « MyButton » n’apparaîtra sur l’onglet de base personnalisé que lorsque les onglets contextuels personnalisés ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="de8c7-269">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="de8c7-270">Mais le groupe parent et l’onglet de base personnalisé s’afficheront indépendamment du fait que les onglets contextuels personnalisés soient pris en charge.</span><span class="sxs-lookup"><span data-stu-id="de8c7-270">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="de8c7-271">Pour plus d’exemples, [voir OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-271">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="de8c7-272">Lorsqu’un onglet parent, un groupe ou un menu est marqué `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` avec, alors il n’est pas visible, et tout son balisage enfant est ignoré, lorsque les onglets contextuels personnalisés ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="de8c7-272">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="de8c7-273">Donc, il n’a pas d’importance si l’un de ces éléments enfant ont `<OverriddenByRibbonApi>` l’élément ou ce que sa valeur est.</span><span class="sxs-lookup"><span data-stu-id="de8c7-273">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="de8c7-274">L’implication de ceci est que si un élément de menu, le contrôle, ou le groupe doit être visible dans tous les contextes, alors non seulement devrait-il pas être marqué `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` avec, mais *son menu d’ancêtre, groupe, et onglet ne doit pas non plus être marqué de cette façon.*</span><span class="sxs-lookup"><span data-stu-id="de8c7-274">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="de8c7-275">Ne marquez pas *tous les* éléments enfant d’un onglet, d’un groupe ou d’un menu avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-275">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="de8c7-276">Cela ne va pas si l’élément parent est marqué pour `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` des raisons données dans le paragraphe précédent.</span><span class="sxs-lookup"><span data-stu-id="de8c7-276">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="de8c7-277">En outre, si vous laissez de côté `<OverriddenByRibbonApi>` le sur le parent (ou le définir à ), alors le parent apparaîtra indépendamment du fait que les `false` onglets contextuels personnalisés sont pris en charge, mais il sera vide quand ils sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="de8c7-277">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="de8c7-278">Ainsi, si tous les éléments enfant ne devraient pas apparaître lorsque les onglets contextuels personnalisés sont pris en charge, marquez le parent, et seulement le parent, avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="de8c7-278">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="de8c7-279">Utilisez des API qui affichent ou cachent un volet de tâche dans des contextes spécifiques</span><span class="sxs-lookup"><span data-stu-id="de8c7-279">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="de8c7-280">Comme alternative à , votre add-in peut définir un volet de tâche avec des contrôles `<OverriddenByRibbonApi>` d’interface utilisateur qui dupliquent la fonctionnalité des contrôles sur un onglet contextuel personnalisé. Ensuite, utilisez [les méthodes Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) [et Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) pour afficher le volet de tâche quand, et seulement quand, l’onglet contextuel aurait été affiché s’il avait été pris en charge.</span><span class="sxs-lookup"><span data-stu-id="de8c7-280">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="de8c7-281">Pour plus de détails sur la façon d’utiliser ces [méthodes, voir Afficher ou masquer le volet de tâche de votre Office Add-in](../develop/show-hide-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="de8c7-281">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="de8c7-282">Gérer l’erreur HostRestartNeeded</span><span class="sxs-lookup"><span data-stu-id="de8c7-282">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="de8c7-283">Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="de8c7-283">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="de8c7-284">Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau.</span><span class="sxs-lookup"><span data-stu-id="de8c7-284">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="de8c7-285">La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué.</span><span class="sxs-lookup"><span data-stu-id="de8c7-285">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="de8c7-286">Votre code doit gérer cette erreur.</span><span class="sxs-lookup"><span data-stu-id="de8c7-286">Your code should handle this error.</span></span> <span data-ttu-id="de8c7-287">Ce qui suit est un exemple de la façon dont.</span><span class="sxs-lookup"><span data-stu-id="de8c7-287">The following is an example of how.</span></span> <span data-ttu-id="de8c7-288">Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="de8c7-288">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
