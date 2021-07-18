---
title: Créer des onglets contextuels personnalisés dans Office de recherche
description: Découvrez comment ajouter des onglets contextuels personnalisés à votre Office de recherche.
ms.date: 07/15/2021
localization_priority: Normal
ms.openlocfilehash: a8eaffe0402601ee11a063d0df5670ff208be4fd
ms.sourcegitcommit: b20041962a7f921a8c40eb9ae55bc6992450b243
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/16/2021
ms.locfileid: "53456228"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="da214-103">Créer des onglets contextuels personnalisés dans Office de recherche</span><span class="sxs-lookup"><span data-stu-id="da214-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="da214-104">Un onglet contextuel est un contrôle onglet masqué dans le ruban Office qui est affiché dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office document.</span><span class="sxs-lookup"><span data-stu-id="da214-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="da214-105">Par exemple, **l’onglet Création** de table qui apparaît sur le ruban Excel lors de la sélection d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="da214-106">Vous incluez des onglets contextuels personnalisés dans votre Office et spécifiez quand ils sont visibles ou masqués en créant des handlers d’événements qui modifient la visibilité.</span><span class="sxs-lookup"><span data-stu-id="da214-106">You include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="da214-107">(Toutefois, les onglets contextuels personnalisés ne répondent pas aux changements de focus.)</span><span class="sxs-lookup"><span data-stu-id="da214-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="da214-108">Cet article suppose que vous connaissez la documentation décrite ci-après.</span><span class="sxs-lookup"><span data-stu-id="da214-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="da214-109">Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).</span><span class="sxs-lookup"><span data-stu-id="da214-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="da214-110">Concepts basiques pour les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da214-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="da214-111">Les onglets contextuels personnalisés sont actuellement uniquement pris en charge sur Excel et uniquement sur ces plateformes et builds :</span><span class="sxs-lookup"><span data-stu-id="da214-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="da214-112">Excel sur Windows (abonnement Microsoft 365 uniquement) : version 2102 (build 13801.20294) ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="da214-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="da214-113">Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="da214-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="da214-114">Les onglets contextuels personnalisés fonctionnent uniquement sur les plateformes qui supportent les ensembles de conditions requises suivants.</span><span class="sxs-lookup"><span data-stu-id="da214-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="da214-115">Pour plus d’informations sur les ensembles de conditions requises et sur leur utilisation, voir Spécifier Office [applications et les exigences d’API.](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="da214-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="da214-116">RibbonApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da214-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="da214-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="da214-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="da214-118">Vous pouvez utiliser les vérifications à l’runtime dans votre code pour tester si la combinaison hôte et plateforme de l’utilisateur prend en charge ces ensembles de conditions requises, comme décrit dans [Spécifier les applications Office](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)et les conditions requises de l’API.</span><span class="sxs-lookup"><span data-stu-id="da214-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="da214-119">(La technique de spécification des ensembles de conditions requises dans le manifeste, également décrite dans cet article, ne fonctionne actuellement pas pour RibbonApi 1.2.) Vous pouvez également implémenter [une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge.](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)</span><span class="sxs-lookup"><span data-stu-id="da214-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="da214-120">Comportement des onglets contextuels personnalisés</span><span class="sxs-lookup"><span data-stu-id="da214-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="da214-121">L’expérience utilisateur pour les onglets contextuels personnalisés suit le modèle des onglets Office contextuels intégrés.</span><span class="sxs-lookup"><span data-stu-id="da214-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="da214-122">Voici les principes de base pour l’emplacement des onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="da214-122">The following are the basic principles for the placement custom contextual tabs.</span></span>

- <span data-ttu-id="da214-123">Lorsqu’un onglet contextuel personnalisé est visible, il apparaît à l’extrémité droite du ruban.</span><span class="sxs-lookup"><span data-stu-id="da214-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="da214-124">Si un ou plusieurs onglets contextuels intégrés et un ou plusieurs onglets contextuels personnalisés des modules sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.</span><span class="sxs-lookup"><span data-stu-id="da214-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="da214-125">Si votre add-in possède plusieurs onglets contextuels et qu’il existe des contextes dans lesquels plusieurs onglets sont visibles, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre add-in.</span><span class="sxs-lookup"><span data-stu-id="da214-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="da214-126">(Le sens est identique à celui de la langue Office ; c’est-à-dire de gauche à droite dans les langues de gauche à droite, mais de droite à gauche dans les langues de droite à gauche.) Pour [plus d’informations sur](#define-the-groups-and-controls-that-appear-on-the-tab) leur définition, voir Définir les groupes et les contrôles qui apparaissent sous l’onglet.</span><span class="sxs-lookup"><span data-stu-id="da214-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="da214-127">Si plusieurs d’entre eux ont un onglet contextuel visible dans un contexte spécifique, ils apparaissent dans l’ordre dans lequel les modules ont été lancés.</span><span class="sxs-lookup"><span data-stu-id="da214-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="da214-128">Contrairement *aux* onglets principaux personnalisés, les onglets contextuels personnalisés ne sont pas ajoutés Office le ruban de l’application.</span><span class="sxs-lookup"><span data-stu-id="da214-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="da214-129">Ils sont présents uniquement dans Office documents sur lesquels votre module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="da214-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="da214-130">Étapes principales pour l’ajout d’un onglet contextuel dans un add-in</span><span class="sxs-lookup"><span data-stu-id="da214-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="da214-131">Voici les principales étapes à suivre pour inclure un onglet contextuel personnalisé dans un add-in.</span><span class="sxs-lookup"><span data-stu-id="da214-131">The following are the major steps for including a custom contextual tab in an add-in.</span></span>

1. <span data-ttu-id="da214-132">Configurez le add-in pour utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="da214-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="da214-133">Définissez l’onglet, ainsi que les groupes et les contrôles qui y apparaissent.</span><span class="sxs-lookup"><span data-stu-id="da214-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="da214-134">Inscrivez l’onglet contextuel avec Office.</span><span class="sxs-lookup"><span data-stu-id="da214-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="da214-135">Spécifiez les circonstances dans le cas où l’onglet sera visible.</span><span class="sxs-lookup"><span data-stu-id="da214-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="da214-136">Configurer le add-in pour utiliser un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="da214-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="da214-137">L’ajout d’onglets contextuels personnalisés nécessite que votre add-in utilise le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="da214-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="da214-138">Pour plus d’informations, [voir Configurer un module complémentaire pour utiliser un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="da214-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="da214-139">Définir les groupes et les contrôles qui apparaissent sous l’onglet</span><span class="sxs-lookup"><span data-stu-id="da214-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="da214-140">Contrairement aux onglets principaux personnalisés, qui sont définis avec du XML dans le manifeste, les onglets contextuels personnalisés sont définis lors de l’runtime avec un blob JSON.</span><span class="sxs-lookup"><span data-stu-id="da214-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="da214-141">Votre code parse le blob dans un objet JavaScript, puis passe l’objet à la [méthode Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="da214-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="da214-142">Les onglets contextuels personnalisés sont uniquement présents dans les documents sur lesquels votre module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="da214-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="da214-143">Cela est différent des onglets principaux personnalisés qui sont ajoutés au ruban de l’application Office lorsque le module est installé et restent présents à l’ouverture d’un autre document.</span><span class="sxs-lookup"><span data-stu-id="da214-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="da214-144">En outre, `requestCreateControls` la méthode ne peut être exécuté qu’une seule fois dans une session de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="da214-144">Also, the `requestCreateControls` method may be run only once in a session of your add-in.</span></span> <span data-ttu-id="da214-145">Si elle est appelée à nouveau, une erreur est lancée.</span><span class="sxs-lookup"><span data-stu-id="da214-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="da214-146">La structure des propriétés et sous-propriétés de l’objet blob JSON (et les noms clés) est à peu près parallèle à la structure de l’élément [CustomTab](../reference/manifest/customtab.md) et de ses éléments descendants dans le manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="da214-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="da214-147">Nous allons créer un exemple d’objet blob JSON onglets contextuel pas à pas.</span><span class="sxs-lookup"><span data-stu-id="da214-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="da214-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="da214-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="da214-149">Si vous travaillez dans Visual Studio Code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON.</span><span class="sxs-lookup"><span data-stu-id="da214-149">If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="da214-150">Pour plus d’informations, voir [Modification de JSON avec Visual Studio Code - Schémas et paramètres JSON.](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)</span><span class="sxs-lookup"><span data-stu-id="da214-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="da214-151">Commencez par créer une chaîne JSON avec deux propriétés de tableau `actions` nommées et `tabs` .</span><span class="sxs-lookup"><span data-stu-id="da214-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="da214-152">Le tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des `actions` contrôles sous l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 20*.</span><span class="sxs-lookup"><span data-stu-id="da214-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="da214-153">Cet exemple simple d’onglet contextuel n’aura qu’un seul bouton et, par conséquent, une seule action.</span><span class="sxs-lookup"><span data-stu-id="da214-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="da214-154">Ajoutez ce qui suit en tant que seul membre du `actions` tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="da214-155">À propos de ce markup, notez :</span><span class="sxs-lookup"><span data-stu-id="da214-155">About this markup, note:</span></span>

    - <span data-ttu-id="da214-156">Les `id` `type` propriétés et les propriétés sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="da214-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="da214-157">La valeur `type` peut être « ExecuteFunction » ou « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="da214-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="da214-158">La `functionName` propriété est utilisée uniquement lorsque la valeur est `type` `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="da214-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="da214-159">Il s’agit du nom d’une fonction définie dans functionFile.</span><span class="sxs-lookup"><span data-stu-id="da214-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="da214-160">Pour plus d’informations sur FunctionFile, voir [Concepts de base pour les commandes de module complémentaire.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="da214-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="da214-161">Dans une étape ultérieure, vous allez ma cartographier cette action sur un bouton de l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="da214-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="da214-162">Ajoutez ce qui suit en tant que seul membre du `tabs` tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="da214-163">À propos de ce markup, notez :</span><span class="sxs-lookup"><span data-stu-id="da214-163">About this markup, note:</span></span>

    - <span data-ttu-id="da214-164">La propriété `id` est requise.</span><span class="sxs-lookup"><span data-stu-id="da214-164">The `id` property is required.</span></span> <span data-ttu-id="da214-165">Utilisez un bref ID descriptif unique parmi tous les onglets contextuels de votre module.</span><span class="sxs-lookup"><span data-stu-id="da214-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="da214-166">La propriété `label` est requise.</span><span class="sxs-lookup"><span data-stu-id="da214-166">The `label` property is required.</span></span> <span data-ttu-id="da214-167">Il s’agit d’une chaîne conviviale qui sert d’étiquette à l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="da214-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="da214-168">La propriété `groups` est requise.</span><span class="sxs-lookup"><span data-stu-id="da214-168">The `groups` property is required.</span></span> <span data-ttu-id="da214-169">Il définit les groupes de contrôles qui apparaîtront sous l’onglet. Elle doit avoir au moins un *membre et pas plus de 20*.</span><span class="sxs-lookup"><span data-stu-id="da214-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="da214-170">(Il existe également des limites au nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui limitent également le nombre de groupes que vous avez.</span><span class="sxs-lookup"><span data-stu-id="da214-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="da214-171">Pour plus d’informations, voir l’étape suivante.)</span><span class="sxs-lookup"><span data-stu-id="da214-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="da214-172">L’objet tabulation peut également avoir une propriété facultative qui spécifie si l’onglet est visible immédiatement au démarrage `visible` du module.</span><span class="sxs-lookup"><span data-stu-id="da214-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="da214-173">Étant donné que les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche leur visibilité (par exemple, lorsque l’utilisateur sélectionne une entité d’un type dans le document), la propriété se présente par défaut lorsqu’elle n’est pas `visible` `false` présente.</span><span class="sxs-lookup"><span data-stu-id="da214-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="da214-174">Dans une section ultérieure, nous montrons comment définir la propriété en réponse `true` à un événement.</span><span class="sxs-lookup"><span data-stu-id="da214-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="da214-175">Dans l’exemple continu simple, l’onglet contextuel ne possède qu’un seul groupe.</span><span class="sxs-lookup"><span data-stu-id="da214-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="da214-176">Ajoutez ce qui suit en tant que seul membre du `groups` tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="da214-177">À propos de ce markup, notez :</span><span class="sxs-lookup"><span data-stu-id="da214-177">About this markup, note:</span></span>

    - <span data-ttu-id="da214-178">Toutes les propriétés sont requises.</span><span class="sxs-lookup"><span data-stu-id="da214-178">All the properties are required.</span></span>
    - <span data-ttu-id="da214-179">La `id` propriété doit être unique parmi tous les groupes de l’onglet. Utilisez un ID bref et descriptif.</span><span class="sxs-lookup"><span data-stu-id="da214-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="da214-180">Il `label` s’agit d’une chaîne conviviale qui sert d’étiquette au groupe.</span><span class="sxs-lookup"><span data-stu-id="da214-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="da214-181">La valeur de la propriété est un tableau d’objets qui spécifient les icônes que le groupe aura sur le ruban en fonction de la taille du ruban et de la fenêtre `icon` d’application Office’application.</span><span class="sxs-lookup"><span data-stu-id="da214-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="da214-182">La valeur de la propriété est un tableau d’objets qui spécifient les boutons et `controls` les menus du groupe.</span><span class="sxs-lookup"><span data-stu-id="da214-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="da214-183">Il doit y en avoir au moins un.</span><span class="sxs-lookup"><span data-stu-id="da214-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="da214-184">*Le nombre total de contrôles sur l’onglet entier ne peut pas être supérieur à 20.*</span><span class="sxs-lookup"><span data-stu-id="da214-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="da214-185">Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.</span><span class="sxs-lookup"><span data-stu-id="da214-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="da214-186">Chaque groupe doit avoir une icône d’au moins deux tailles, 32 x 32 px et 80 x 80 px.</span><span class="sxs-lookup"><span data-stu-id="da214-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="da214-187">Si vous le souhaitez, vous pouvez également avoir des icônes de tailles 16 x 16 px, 20 x 20 px, 24 x 24 px, 40 x 40 px, 48 x 48 px et 64 x 64 px.</span><span class="sxs-lookup"><span data-stu-id="da214-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="da214-188">Office l’icône à utiliser en fonction de la taille du ruban et de la Office’application.</span><span class="sxs-lookup"><span data-stu-id="da214-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="da214-189">Ajoutez les objets suivants au tableau d’icônes.</span><span class="sxs-lookup"><span data-stu-id="da214-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="da214-190">(Si les tailles de la fenêtre et  du ruban sont suffisamment grandes pour qu’au moins l’un des contrôles du groupe apparaisse, aucune icône de groupe ne s’affiche.</span><span class="sxs-lookup"><span data-stu-id="da214-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="da214-191">Pour obtenir un exemple, regardez le groupe **Styles** sur le ruban Word lorsque vous réduirez et développez la fenêtre Word.) À propos de ce markup, notez :</span><span class="sxs-lookup"><span data-stu-id="da214-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="da214-192">Les deux propriétés sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="da214-192">Both the properties are required.</span></span>
    - <span data-ttu-id="da214-193">`size`L’unité de mesure de propriété est pixels.</span><span class="sxs-lookup"><span data-stu-id="da214-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="da214-194">Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.</span><span class="sxs-lookup"><span data-stu-id="da214-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="da214-195">La `sourceLocation` propriété spécifie l’URL complète de l’icône.</span><span class="sxs-lookup"><span data-stu-id="da214-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="da214-196">Tout comme vous devez généralement modifier les URL dans le manifeste du add-in lorsque vous passez du développement à la production (par exemple, en modifiant le domaine localhost en contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.</span><span class="sxs-lookup"><span data-stu-id="da214-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="da214-197">Dans notre exemple simple en cours, le groupe ne possède qu’un seul bouton.</span><span class="sxs-lookup"><span data-stu-id="da214-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="da214-198">Ajoutez l’objet suivant comme seul membre du `controls` tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="da214-199">À propos de ce markup, notez :</span><span class="sxs-lookup"><span data-stu-id="da214-199">About this markup, note:</span></span>

    - <span data-ttu-id="da214-200">Toutes les propriétés, à l’exception `enabled` de , sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="da214-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="da214-201">`type` spécifie le type de contrôle.</span><span class="sxs-lookup"><span data-stu-id="da214-201">`type` specifies the type of control.</span></span> <span data-ttu-id="da214-202">Les valeurs peuvent être « Button », « Menu » ou « MobileButton ».</span><span class="sxs-lookup"><span data-stu-id="da214-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="da214-203">`id` peut prendre jusqu’à 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="da214-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="da214-204">`actionId` doit être l’ID d’une action définie dans le `actions` tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="da214-205">(Voir l’étape 1 de cette section.)</span><span class="sxs-lookup"><span data-stu-id="da214-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="da214-206">`label` est une chaîne conviviale qui sert d’étiquette au bouton.</span><span class="sxs-lookup"><span data-stu-id="da214-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="da214-207">`superTip` représente une forme enrichie d’info-conseil.</span><span class="sxs-lookup"><span data-stu-id="da214-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="da214-208">Les `title` propriétés et les `description` propriétés sont requises.</span><span class="sxs-lookup"><span data-stu-id="da214-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="da214-209">`icon` spécifie les icônes du bouton.</span><span class="sxs-lookup"><span data-stu-id="da214-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="da214-210">Les remarques précédentes sur l’icône de groupe s’appliquent également ici.</span><span class="sxs-lookup"><span data-stu-id="da214-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="da214-211">`enabled` (facultatif) indique si le bouton est activé au démarrage de l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="da214-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="da214-212">La valeur par défaut, si elle n’est pas présente, est `true` .</span><span class="sxs-lookup"><span data-stu-id="da214-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="da214-213">Voici l’exemple complet du blob JSON.</span><span class="sxs-lookup"><span data-stu-id="da214-213">The following is the complete example of the JSON blob.</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="da214-214">Inscrire l’onglet contextuel Office avec requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="da214-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="da214-215">L’onglet contextuel est inscrit auprès Office en appelant [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="da214-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="da214-216">Cette tâche est généralement effectuée dans la fonction affectée à la méthode ou `Office.initialize` avec `Office.onReady` celle-ci.</span><span class="sxs-lookup"><span data-stu-id="da214-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="da214-217">Pour plus d’informations sur ces méthodes et l’initialisation du Office, voir [Initialiser votre Office.](../develop/initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="da214-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="da214-218">Vous pouvez toutefois appeler la méthode à tout moment après l’initialisation.</span><span class="sxs-lookup"><span data-stu-id="da214-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="da214-219">La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un add-in.</span><span class="sxs-lookup"><span data-stu-id="da214-219">The `requestCreateControls` method may be called only once in a given session of an add-in.</span></span> <span data-ttu-id="da214-220">Une erreur est lancée si elle est appelée à nouveau.</span><span class="sxs-lookup"><span data-stu-id="da214-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="da214-221">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="da214-221">The following is an example.</span></span> <span data-ttu-id="da214-222">Notez que la chaîne JSON doit être convertie en objet JavaScript avec la méthode pour pouvoir être transmise `JSON.parse` à une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="da214-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="da214-223">Spécifier les contextes où l’onglet sera visible avec requestUpdate</span><span class="sxs-lookup"><span data-stu-id="da214-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="da214-224">En règle générale, un onglet contextuel personnalisé doit apparaître lorsqu’un événement initié par l’utilisateur modifie le contexte du add-in.</span><span class="sxs-lookup"><span data-stu-id="da214-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="da214-225">Envisagez un scénario dans lequel l’onglet doit être visible lorsque, et uniquement quand, un graphique (dans la feuille de calcul par défaut d’un Excel)) est activé.</span><span class="sxs-lookup"><span data-stu-id="da214-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="da214-226">Commencez par affecter des handlers.</span><span class="sxs-lookup"><span data-stu-id="da214-226">Begin by assigning handlers.</span></span> <span data-ttu-id="da214-227">Cela est généralement effectué dans la méthode comme dans l’exemple suivant qui affecte des handlers (créés à une étape ultérieure) aux événements et aux graphiques de la feuille `Office.onReady` `onActivated` de `onDeactivated` calcul.</span><span class="sxs-lookup"><span data-stu-id="da214-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="da214-228">Ensuite, définissez les handlers.</span><span class="sxs-lookup"><span data-stu-id="da214-228">Next, define the handlers.</span></span> <span data-ttu-id="da214-229">Voici un exemple simple d’une erreur, mais voir Gestion de l’erreur `showDataTab` [HostRestartNeeded](#handle-the-hostrestartneeded-error) plus loin dans cet article pour obtenir une version plus robuste de la fonction.</span><span class="sxs-lookup"><span data-stu-id="da214-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="da214-230">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="da214-230">About this code, note:</span></span>

- <span data-ttu-id="da214-231">Office effectue un contrôle lorsqu’il met à jour l’état du ruban.</span><span class="sxs-lookup"><span data-stu-id="da214-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="da214-232">La [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) met en file d’attente une demande de mise à jour.</span><span class="sxs-lookup"><span data-stu-id="da214-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="da214-233">La méthode résout l’objet dès qu’il a mis la demande en file d’attente, et non lorsque `Promise` le ruban est réellement mis à jour.</span><span class="sxs-lookup"><span data-stu-id="da214-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="da214-234">Le paramètre de la méthode est un objet `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID exactement comme spécifié dans le *JSON* et (2) spécifie la visibilité de l’onglet.</span><span class="sxs-lookup"><span data-stu-id="da214-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="da214-235">Si vous avez plusieurs onglets contextuels personnalisés qui doivent être visibles dans le même contexte, il vous suffit d’ajouter des objets onglet supplémentaires au `tabs` tableau.</span><span class="sxs-lookup"><span data-stu-id="da214-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="da214-236">Le handler pour masquer l’onglet est presque identique, sauf qu’il définit à `visible` nouveau la propriété sur `false` .</span><span class="sxs-lookup"><span data-stu-id="da214-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="da214-237">La Office JavaScript fournit également plusieurs interfaces (types) pour faciliter la construction de `RibbonUpdateData` l’objet.</span><span class="sxs-lookup"><span data-stu-id="da214-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="da214-238">Voici la fonction `showDataTab` dans TypeScript qui utilise ces types.</span><span class="sxs-lookup"><span data-stu-id="da214-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="da214-239">Activer la visibilité de l’onglet et l’état activé d’un bouton en même temps</span><span class="sxs-lookup"><span data-stu-id="da214-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="da214-240">La méthode est également utilisée pour activer ou désactiver l’état d’un bouton personnalisé sur un onglet contextuel personnalisé ou un `requestUpdate` onglet principal personnalisé. Pour plus d’informations à ce sujet, voir [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="da214-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="da214-241">Il peut y avoir des scénarios dans lesquels vous souhaitez modifier la visibilité d’un onglet et l’état activé d’un bouton en même temps.</span><span class="sxs-lookup"><span data-stu-id="da214-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="da214-242">Vous le faites avec un seul appel de `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="da214-242">You do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="da214-243">Voici un exemple dans lequel un bouton d’un onglet principal est activé en même temps qu’un onglet contextuel est rendu visible.</span><span class="sxs-lookup"><span data-stu-id="da214-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="da214-244">Dans l’exemple suivant, le bouton activé se trouve sur le même onglet contextuel que celui qui est rendu visible.</span><span class="sxs-lookup"><span data-stu-id="da214-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="open-a-task-pane-from-contextual-tabs"></a><span data-ttu-id="da214-245">Ouvrir un volet Des tâches à partir d’onglets contextuels</span><span class="sxs-lookup"><span data-stu-id="da214-245">Open a task pane from contextual tabs</span></span>

<span data-ttu-id="da214-246">Pour ouvrir votre volet Des tâches à partir d’un bouton d’un onglet contextuel personnalisé, créez une action dans le JSON avec une `type` des `ShowTaskpane` touches .</span><span class="sxs-lookup"><span data-stu-id="da214-246">To open your task pane from a button on a custom contextual tab, create an action in the JSON with a `type` of `ShowTaskpane`.</span></span> <span data-ttu-id="da214-247">Définissez ensuite un bouton dont `actionId` la propriété est définie sur la valeur de `id` l’action.</span><span class="sxs-lookup"><span data-stu-id="da214-247">Then define a button with the `actionId` property set to the `id` of the action.</span></span> <span data-ttu-id="da214-248">Cela ouvre le volet Des tâches par défaut spécifié par `<Runtime>` l’élément dans votre manifeste.</span><span class="sxs-lookup"><span data-stu-id="da214-248">This opens the default task pane specified by the `<Runtime>` element in your manifest.</span></span>

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

<span data-ttu-id="da214-249">Pour ouvrir un volet De tâches qui n’est pas le volet Des tâches par défaut, spécifiez une `sourceLocation` propriété dans la définition de l’action.</span><span class="sxs-lookup"><span data-stu-id="da214-249">To open any task pane that is not the default task pane, specify a `sourceLocation` property in the definition of the action.</span></span> <span data-ttu-id="da214-250">Dans l’exemple suivant, un deuxième volet Des tâches est ouvert à partir d’un autre bouton.</span><span class="sxs-lookup"><span data-stu-id="da214-250">In the following example, a second task pane is opened from a different button.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="da214-251">`sourceLocation`Lorsqu’une valeur est spécifiée pour l’action, le volet Des tâches *n’utilise* pas le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="da214-251">When a `sourceLocation` is specified for the action, then the task pane does *not* use the shared runtime.</span></span> <span data-ttu-id="da214-252">Il s’exécute dans un nouveau runtime JavaScript.</span><span class="sxs-lookup"><span data-stu-id="da214-252">It runs in a new JavaScript runtime.</span></span>
> - <span data-ttu-id="da214-253">Un seul volet De tâches ne peut pas utiliser le runtime partagé, de sorte qu’une seule action de type ne peut `ShowTaskpane` pas omettre la `sourceLocation` propriété.</span><span class="sxs-lookup"><span data-stu-id="da214-253">No more than one task pane can use the shared runtime, so no more than one action of type `ShowTaskpane` can omit the `sourceLocation` property.</span></span>

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a><span data-ttu-id="da214-254">Localiser le texte JSON</span><span class="sxs-lookup"><span data-stu-id="da214-254">Localize the JSON text</span></span>

<span data-ttu-id="da214-255">Le blob JSON passé à n’est pas localisée de la même façon que le marques de manifeste pour les onglets principaux personnalisés est localisée (ce qui est décrit lors de la localisation du contrôle à partir du `requestCreateControls` [manifeste).](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="da214-255">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="da214-256">Au lieu de cela, la localisation doit se produire lors de l’runtime à l’aide de blobs JSON distincts pour chaque paramètre régional.</span><span class="sxs-lookup"><span data-stu-id="da214-256">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="da214-257">Nous vous suggérons d’utiliser `switch` une instruction qui teste la [Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage)</span><span class="sxs-lookup"><span data-stu-id="da214-257">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="da214-258">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="da214-258">The following is an example.</span></span>

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

<span data-ttu-id="da214-259">Ensuite, votre code appelle la fonction pour obtenir l’objet blob local qui est transmis `requestCreateControls` à , comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="da214-259">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example.</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="da214-260">Meilleures pratiques pour les onglets contextuels personnalisés</span><span class="sxs-lookup"><span data-stu-id="da214-260">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="da214-261">Implémenter une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge</span><span class="sxs-lookup"><span data-stu-id="da214-261">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="da214-262">Certaines combinaisons de plateforme, Office application et de Office build ne sont pas prise en `requestCreateControls` charge.</span><span class="sxs-lookup"><span data-stu-id="da214-262">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="da214-263">Votre add-in doit être conçu pour offrir une expérience de remplacement aux utilisateurs qui exécutent le module sur l’une de ces combinaisons.</span><span class="sxs-lookup"><span data-stu-id="da214-263">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="da214-264">Les sections suivantes décrivent deux façons de fournir une expérience de retour.</span><span class="sxs-lookup"><span data-stu-id="da214-264">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="da214-265">Utiliser des onglets ou des contrôles nontexte</span><span class="sxs-lookup"><span data-stu-id="da214-265">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="da214-266">Il existe un élément manifeste, [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)conçu pour créer une expérience de base dans un application qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés.</span><span class="sxs-lookup"><span data-stu-id="da214-266">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="da214-267">La stratégie la plus simple pour utiliser cet élément est que vous définissez  dans le manifeste un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés nontexte) qui dupliquent les personnalisations du ruban des onglets contextuels personnalisés dans votre application.</span><span class="sxs-lookup"><span data-stu-id="da214-267">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="da214-268">Mais vous `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` ajoutez en tant que premier élément enfant de [CustomTab](../reference/manifest/customtab.md).</span><span class="sxs-lookup"><span data-stu-id="da214-268">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="da214-269">L’effet de cette utilisation est le suivant :</span><span class="sxs-lookup"><span data-stu-id="da214-269">The effect of doing so is the following:</span></span>

- <span data-ttu-id="da214-270">Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, l’onglet principal personnalisé n’apparaît pas sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="da214-270">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="da214-271">Au lieu de cela, l’onglet contextuel personnalisé est créé lorsque le add-in appelle la `requestCreateControls` méthode.</span><span class="sxs-lookup"><span data-stu-id="da214-271">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="da214-272">Si le add-in *s’exécute* sur une application ou une plateforme qui ne prend pas en charge, l’onglet principal personnalisé `requestCreateControls` apparaît sur le ruban.</span><span class="sxs-lookup"><span data-stu-id="da214-272">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="da214-273">Voici un exemple de cette stratégie simple.</span><span class="sxs-lookup"><span data-stu-id="da214-273">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="da214-274">Cette stratégie simple utilise un onglet principal personnalisé qui met en miroir un onglet contextuel personnalisé avec ses groupes et contrôles enfants, mais vous pouvez utiliser une stratégie plus complexe.</span><span class="sxs-lookup"><span data-stu-id="da214-274">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="da214-275">L’élément peut également être ajouté en tant que (le premier) élément enfant aux éléments Group et Control (type de bouton et type de `<OverriddenByRibbonApi>` [menu)](../reference/manifest/control.md#menu-dropdown-button-controls)et [](../reference/manifest/group.md) [](../reference/manifest/control.md) aux éléments de [](../reference/manifest/control.md#button-control) `<Item>` menu.</span><span class="sxs-lookup"><span data-stu-id="da214-275">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="da214-276">Cela vous permet de distribuer les groupes et les contrôles qui apparaîtraient dans l’onglet contextuel entre différents groupes, boutons et menus dans différents onglets principaux personnalisés.</span><span class="sxs-lookup"><span data-stu-id="da214-276">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="da214-277">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="da214-277">The following is an example.</span></span> <span data-ttu-id="da214-278">Notez que « MyButton » apparaît sur l’onglet principal personnalisé uniquement lorsque les onglets contextuels personnalisés ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="da214-278">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="da214-279">Toutefois, le groupe parent et l’onglet principal personnalisé apparaissent, que les onglets contextuels personnalisés soient pris en charge ou non.</span><span class="sxs-lookup"><span data-stu-id="da214-279">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="da214-280">Pour plus d’exemples, [voir OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="da214-280">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="da214-281">Lorsqu’un onglet, un groupe ou un menu parent est marqué avec, il n’est pas visible et tout son markup enfant est ignoré, lorsque les onglets contextuels personnalisés ne sont pas pris en `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` charge.</span><span class="sxs-lookup"><span data-stu-id="da214-281">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="da214-282">Ainsi, peu importe si l’un de ces éléments enfants a l’élément ou `<OverriddenByRibbonApi>` sa valeur.</span><span class="sxs-lookup"><span data-stu-id="da214-282">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="da214-283">En conséquence, si un élément de menu, un contrôle ou un groupe doit être visible dans tous les contextes, non seulement il ne doit pas être marqué avec, mais son `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *ancêtre menu,* groupe et onglet ne doit pas non plus être marqué de cette façon.</span><span class="sxs-lookup"><span data-stu-id="da214-283">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="da214-284">Ne marquez pas *tous les* éléments enfants d’un onglet, d’un groupe ou d’un menu avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="da214-284">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="da214-285">Cela est inutile si l’élément parent est marqué pour `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` des raisons indiquées dans le paragraphe précédent.</span><span class="sxs-lookup"><span data-stu-id="da214-285">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="da214-286">En outre, si vous ne le faites pas sur le parent (ou si vous le définissez sur ), le parent apparaît, que les onglets contextuels personnalisés soient pris en charge ou non, mais qu’ils soient vides lorsqu’ils sont pris en `<OverriddenByRibbonApi>` `false` charge.</span><span class="sxs-lookup"><span data-stu-id="da214-286">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="da214-287">Ainsi, si tous les éléments enfants ne doivent pas apparaître lorsque les onglets contextuels personnalisés sont pris en charge, marquez le parent et uniquement le parent, avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="da214-287">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="da214-288">Utiliser des API qui montrent ou masquent un volet Des tâches dans des contextes spécifiés</span><span class="sxs-lookup"><span data-stu-id="da214-288">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="da214-289">En remplacement, votre add-in peut définir un volet Des tâches avec des contrôles d’interface utilisateur qui dupliquent la fonctionnalité des contrôles sur un `<OverriddenByRibbonApi>` onglet contextuel personnalisé. Utilisez ensuite les [méthodes Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) et [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) pour afficher le volet Des tâches quand et uniquement quand l’onglet contextuel aurait été affiché s’il était pris en charge.</span><span class="sxs-lookup"><span data-stu-id="da214-289">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="da214-290">Pour plus d’informations sur l’utilisation de ces méthodes, voir Afficher ou masquer le volet Des tâches de [votre Office.](../develop/show-hide-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="da214-290">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="da214-291">Gérer l’erreur HostRestartNeeded</span><span class="sxs-lookup"><span data-stu-id="da214-291">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="da214-292">Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="da214-292">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="da214-293">Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau.</span><span class="sxs-lookup"><span data-stu-id="da214-293">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="da214-294">La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué.</span><span class="sxs-lookup"><span data-stu-id="da214-294">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="da214-295">Votre code doit gérer cette erreur.</span><span class="sxs-lookup"><span data-stu-id="da214-295">Your code should handle this error.</span></span> <span data-ttu-id="da214-296">Voici un exemple de comment.</span><span class="sxs-lookup"><span data-stu-id="da214-296">The following is an example of how.</span></span> <span data-ttu-id="da214-297">Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="da214-297">In this case, the `reportError` method displays the error to the user.</span></span>

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
