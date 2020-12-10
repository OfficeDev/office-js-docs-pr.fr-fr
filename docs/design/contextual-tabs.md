---
title: Créer des onglets contextuels personnalisés dans les compléments Office
description: Découvrez comment ajouter des onglets contextuels personnalisés à votre complément Office.
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: d8617c7dd8748d15393c0e38c527062e5894e791
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612735"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="e943b-103">Créer des onglets contextuels personnalisés dans les compléments Office (aperçu)</span><span class="sxs-lookup"><span data-stu-id="e943b-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="e943b-104">Un onglet contextuel est un contrôle d’onglet masqué dans le ruban Office qui s’affiche dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="e943b-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="e943b-105">Par exemple, l’onglet **création de table** qui s’affiche dans le ruban Excel lorsqu’un tableau est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="e943b-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="e943b-106">Vous pouvez inclure des onglets contextuels personnalisés dans votre complément Office et spécifier lorsqu’ils sont visibles ou masqués, en créant des gestionnaires d’événements qui modifient la visibilité.</span><span class="sxs-lookup"><span data-stu-id="e943b-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="e943b-107">(Toutefois, les onglets contextuels personnalisés ne répondent pas aux changements de focus.)</span><span class="sxs-lookup"><span data-stu-id="e943b-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="e943b-108">Cet article suppose que vous connaissez la documentation décrite ci-après.</span><span class="sxs-lookup"><span data-stu-id="e943b-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="e943b-109">Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).</span><span class="sxs-lookup"><span data-stu-id="e943b-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="e943b-110">Concepts basiques pour les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="e943b-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="e943b-111">Les onglets contextuels personnalisés sont en aperçu.</span><span class="sxs-lookup"><span data-stu-id="e943b-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="e943b-112">Essayez de les tester dans un environnement de développement ou de test, mais ne les ajoutez pas à un complément de production.</span><span class="sxs-lookup"><span data-stu-id="e943b-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="e943b-113">Les onglets contextuels personnalisés sont actuellement uniquement pris en charge sur Excel et uniquement sur ces plateformes et génèrent les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="e943b-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="e943b-114">Excel sur Windows (Microsoft 365 uniquement, pas une licence perpétuelle) : version 2011 (Build 13426,20274).</span><span class="sxs-lookup"><span data-stu-id="e943b-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="e943b-115">Votre abonnement Microsoft 365 doit peut-être être sur le [canal actuel (](https://insider.office.com/join/windows) préversion) précédemment appelé « canal mensuel (ciblé) » ou « Insider Slower ».</span><span class="sxs-lookup"><span data-stu-id="e943b-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="e943b-116">Les onglets contextuels personnalisés fonctionnent uniquement sur les plateformes qui prennent en charge les ensembles de conditions requises suivants.</span><span class="sxs-lookup"><span data-stu-id="e943b-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="e943b-117">Pour plus d’informations sur les ensembles de conditions requises et la façon de les utiliser, voir [spécifier les applications Office et les conditions requises](../develop/specify-office-hosts-and-api-requirements.md)pour les API.</span><span class="sxs-lookup"><span data-stu-id="e943b-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="e943b-118">SharedRuntime 1,1</span><span class="sxs-lookup"><span data-stu-id="e943b-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="e943b-119">Comportement des onglets contextuels personnalisés</span><span class="sxs-lookup"><span data-stu-id="e943b-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="e943b-120">L’expérience utilisateur pour les onglets contextuels personnalisés suit le modèle des onglets contextuels Office intégrés.</span><span class="sxs-lookup"><span data-stu-id="e943b-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="e943b-121">Les éléments suivants sont les principes de base pour les onglets contextuels personnalisés de placement :</span><span class="sxs-lookup"><span data-stu-id="e943b-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="e943b-122">Lorsqu’un onglet contextuel personnalisé est visible, il s’affiche à l’extrémité droite du ruban.</span><span class="sxs-lookup"><span data-stu-id="e943b-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="e943b-123">Si un ou plusieurs onglets contextuels prédéfinis et un ou plusieurs onglets contextuels personnalisés provenant de compléments sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.</span><span class="sxs-lookup"><span data-stu-id="e943b-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="e943b-124">Si votre complément comporte plusieurs onglets contextuels et des contextes dans lesquels plusieurs sont visibles, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="e943b-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="e943b-125">(Le sens est le même que celui de la langue d’Office ; autrement dit, de gauche à droite dans les langues se lisant de gauche à droite, mais de droite à gauche dans les langues se lisant de droite à gauche.) Pour plus d’informations sur la définition des [groupes et des contrôles de l’onglet, voir define the Groups and Controls](#define-the-groups-and-controls-that-appear-on-the-tab) .</span><span class="sxs-lookup"><span data-stu-id="e943b-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="e943b-126">Si plusieurs compléments disposent d’un onglet contextuel visible dans un contexte spécifique, ils apparaissent dans l’ordre dans lequel les compléments ont été lancés.</span><span class="sxs-lookup"><span data-stu-id="e943b-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="e943b-127">Les onglets *contextuels* personnalisés, contrairement aux onglets principaux personnalisés, ne sont pas ajoutés de façon permanente au ruban de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="e943b-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="e943b-128">Ils sont présents uniquement dans les documents Office sur lesquels votre complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e943b-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="e943b-129">Étapes principales pour l’inclusion d’un onglet contextuel dans un complément</span><span class="sxs-lookup"><span data-stu-id="e943b-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="e943b-130">Voici les principales étapes à suivre pour inclure un onglet contextuel personnalisé dans un complément :</span><span class="sxs-lookup"><span data-stu-id="e943b-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="e943b-131">Configurez le complément pour qu’il utilise un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="e943b-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="e943b-132">Définissez l’onglet, ainsi que les groupes et les contrôles qui s’y trouvent.</span><span class="sxs-lookup"><span data-stu-id="e943b-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="e943b-133">Enregistrer l’onglet contextuel avec Office.</span><span class="sxs-lookup"><span data-stu-id="e943b-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="e943b-134">Spécifier les circonstances dans lesquelles l’onglet sera visible.</span><span class="sxs-lookup"><span data-stu-id="e943b-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="e943b-135">Configurer le complément pour utiliser un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="e943b-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="e943b-136">L’ajout d’onglets contextuels personnalisés nécessite que votre complément utilise le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="e943b-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="e943b-137">Pour plus d’informations, consultez [la rubrique Configure an Add-in use a Shared Runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="e943b-137">For more information, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="e943b-138">Définir les groupes et les contrôles qui apparaissent sur l’onglet</span><span class="sxs-lookup"><span data-stu-id="e943b-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="e943b-139">Contrairement aux onglets principaux personnalisés, qui sont définis avec XML dans le manifeste, les onglets contextuels personnalisés sont définis au moment de l’exécution avec un blob JSON.</span><span class="sxs-lookup"><span data-stu-id="e943b-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="e943b-140">Votre code analyse le BLOB dans un objet JavaScript, puis passe l’objet à la méthode [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) .</span><span class="sxs-lookup"><span data-stu-id="e943b-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="e943b-141">Les onglets contextuels personnalisés sont présents uniquement dans les documents sur lesquels votre complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e943b-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="e943b-142">Cette fonction est différente des onglets principaux personnalisés qui sont ajoutés au ruban de l’application Office lorsque le complément est installé et reste présent lors de l’ouverture d’un autre document.</span><span class="sxs-lookup"><span data-stu-id="e943b-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="e943b-143">De plus, la `requestCreateControls` méthode ne peut être exécutée qu’une seule fois dans une session de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e943b-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="e943b-144">Si elle est encore appelée, une erreur est générée.</span><span class="sxs-lookup"><span data-stu-id="e943b-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="e943b-145">La structure des propriétés et sous-propriétés du BLOB JSON (et les noms de clés) est à peu près parallèle à la structure de l’élément [CustomTab](../reference/manifest/customtab.md) et ses éléments descendants dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="e943b-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="e943b-146">Nous allons construire un exemple d’objet de BLOB JSON d’onglets contextuels pas à pas.</span><span class="sxs-lookup"><span data-stu-id="e943b-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="e943b-147">(Le schéma complet pour l’onglet contextuel JSON se trouve [ surdynamic-ribbon.schema.js](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="e943b-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="e943b-148">Ce lien peut ne pas fonctionner dans la période d’aperçu anticipé des onglets contextuels.</span><span class="sxs-lookup"><span data-stu-id="e943b-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="e943b-149">Si le lien ne fonctionne pas, vous pouvez trouver le dernier brouillon du schéma [sur brouillon dynamic-ribbon.schema.jssur](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Si vous travaillez dans Visual Studio code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON.</span><span class="sxs-lookup"><span data-stu-id="e943b-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="e943b-150">Pour plus d’informations, consultez la rubrique [Editing JSON with Visual Studio code-JSON schemas and Settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="e943b-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="e943b-151">Commencez par créer une chaîne JSON avec deux propriétés de tableau nommées `actions` et `tabs` .</span><span class="sxs-lookup"><span data-stu-id="e943b-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="e943b-152">Le `actions` tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des contrôles dans l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 10*.</span><span class="sxs-lookup"><span data-stu-id="e943b-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="e943b-153">Cet exemple simple d’onglet contextuel ne comporte qu’un seul bouton et, par conséquent, une seule action.</span><span class="sxs-lookup"><span data-stu-id="e943b-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="e943b-154">Ajoutez ce qui suit en tant que membre unique du `actions` tableau.</span><span class="sxs-lookup"><span data-stu-id="e943b-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="e943b-155">À propos de ce balisage, notez les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="e943b-155">About this markup, note:</span></span>

    - <span data-ttu-id="e943b-156">Les `id` `type` Propriétés et sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="e943b-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="e943b-157">La valeur de `type` peut être « ExecuteFunction » ou « ShowTaskpane ».</span><span class="sxs-lookup"><span data-stu-id="e943b-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="e943b-158">La `functionName` propriété est utilisée uniquement lorsque la valeur de `type` est `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="e943b-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="e943b-159">Il s’agit du nom d’une fonction définie dans FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="e943b-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="e943b-160">Pour plus d’informations sur les FunctionFile, consultez la rubrique [concepts de base pour les commandes de complément](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="e943b-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="e943b-161">Dans une étape ultérieure, vous mapperez cette action sur un bouton de l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="e943b-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="e943b-162">Ajoutez ce qui suit en tant que membre unique du `tabs` tableau.</span><span class="sxs-lookup"><span data-stu-id="e943b-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="e943b-163">À propos de ce balisage, notez les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="e943b-163">About this markup, note:</span></span>

    - <span data-ttu-id="e943b-164">La propriété `id` est requise.</span><span class="sxs-lookup"><span data-stu-id="e943b-164">The `id` property is required.</span></span> <span data-ttu-id="e943b-165">Utilisez un bref ID descriptif qui est unique parmi tous les onglets contextuels de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e943b-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="e943b-166">La propriété `label` est requise.</span><span class="sxs-lookup"><span data-stu-id="e943b-166">The `label` property is required.</span></span> <span data-ttu-id="e943b-167">Il s’agit d’une chaîne conviviale qui servira d’étiquette de l’onglet contextuel.</span><span class="sxs-lookup"><span data-stu-id="e943b-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="e943b-168">La propriété `groups` est requise.</span><span class="sxs-lookup"><span data-stu-id="e943b-168">The `groups` property is required.</span></span> <span data-ttu-id="e943b-169">Il définit les groupes de contrôles qui s’affichent sous l’onglet. Il doit comporter au moins un membre *et pas plus de 20*.</span><span class="sxs-lookup"><span data-stu-id="e943b-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="e943b-170">(Il existe également des limites quant au nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui contraignent également le nombre de groupes dont vous disposez.</span><span class="sxs-lookup"><span data-stu-id="e943b-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="e943b-171">Pour plus d’informations, reportez-vous à l’étape suivante.)</span><span class="sxs-lookup"><span data-stu-id="e943b-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="e943b-172">L’objet Tab peut également avoir une `visible` propriété facultative qui spécifie si l’onglet est visible immédiatement au démarrage du complément.</span><span class="sxs-lookup"><span data-stu-id="e943b-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="e943b-173">Étant donné que les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche sa visibilité (par exemple, l’utilisateur sélectionnant une entité d’un certain type dans le document), la `visible` propriété est définie par défaut `false` lorsqu’il n’est pas présent.</span><span class="sxs-lookup"><span data-stu-id="e943b-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="e943b-174">Dans une section ultérieure, nous montrons comment définir la propriété sur `true` en réponse à un événement.</span><span class="sxs-lookup"><span data-stu-id="e943b-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="e943b-175">Dans l’exemple simple en cours, l’onglet contextuel ne comporte qu’un seul groupe.</span><span class="sxs-lookup"><span data-stu-id="e943b-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="e943b-176">Ajoutez ce qui suit en tant que membre unique du `groups` tableau.</span><span class="sxs-lookup"><span data-stu-id="e943b-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="e943b-177">À propos de ce balisage, notez les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="e943b-177">About this markup, note:</span></span>

    - <span data-ttu-id="e943b-178">Toutes les propriétés sont requises.</span><span class="sxs-lookup"><span data-stu-id="e943b-178">All the properties are required.</span></span>
    - <span data-ttu-id="e943b-179">La `id` propriété doit être unique parmi tous les groupes de l’onglet. Utilisez un bref ID descriptif.</span><span class="sxs-lookup"><span data-stu-id="e943b-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="e943b-180">`label`Est une chaîne conviviale qui sert d’étiquette au groupe.</span><span class="sxs-lookup"><span data-stu-id="e943b-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="e943b-181">La `icon` valeur de la propriété est un tableau d’objets qui spécifient les icônes du groupe sur le ruban en fonction de la taille du ruban et de la fenêtre de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="e943b-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="e943b-182">La `controls` valeur de la propriété est un tableau d’objets qui spécifient les boutons et les menus du groupe.</span><span class="sxs-lookup"><span data-stu-id="e943b-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="e943b-183">Il doit y avoir au moins un et *six pour un groupe*.</span><span class="sxs-lookup"><span data-stu-id="e943b-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="e943b-184">*Le nombre total de contrôles sur l’onglet entier ne peut pas être supérieur à 20.*</span><span class="sxs-lookup"><span data-stu-id="e943b-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="e943b-185">Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun, et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.</span><span class="sxs-lookup"><span data-stu-id="e943b-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="e943b-186">Chaque groupe doit avoir une icône d’au moins deux tailles, 32x32 PX et 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="e943b-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="e943b-187">Si vous le souhaitez, vous pouvez également avoir des icônes de taille de 16x16 PX, 20x20 PX, 24x24 PX, 40x40 PX, 48 x 48 px et 64 x 64 px.</span><span class="sxs-lookup"><span data-stu-id="e943b-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="e943b-188">Office décide de l’icône à utiliser en fonction de la taille du ruban et de la fenêtre de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="e943b-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="e943b-189">Ajoutez les objets suivants au tableau d’icônes.</span><span class="sxs-lookup"><span data-stu-id="e943b-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="e943b-190">(Si la taille de la fenêtre et du ruban est suffisante pour qu’au moins un des *contrôles* du groupe s’affiche, aucune icône de groupe n’apparaît.</span><span class="sxs-lookup"><span data-stu-id="e943b-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="e943b-191">Pour obtenir un exemple, Regardez le groupe **styles** sur le ruban Word lorsque vous réduisez et développez la fenêtre Word.) À propos de ce balisage, notez les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="e943b-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="e943b-192">Les deux propriétés sont requises.</span><span class="sxs-lookup"><span data-stu-id="e943b-192">Both the properties are required.</span></span>
    - <span data-ttu-id="e943b-193">L' `size` unité de mesure de la propriété est exprimée en pixels.</span><span class="sxs-lookup"><span data-stu-id="e943b-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="e943b-194">Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.</span><span class="sxs-lookup"><span data-stu-id="e943b-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="e943b-195">La `sourceLocation` propriété spécifie l’URL complète de l’icône.</span><span class="sxs-lookup"><span data-stu-id="e943b-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="e943b-196">Tout comme vous devez généralement modifier les URL dans le manifeste du complément lorsque vous passez du développement à la production (par exemple, en modifiant le domaine de localhost vers contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.</span><span class="sxs-lookup"><span data-stu-id="e943b-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="e943b-197">Dans notre exemple simple en cours, le groupe ne possède qu’un seul bouton.</span><span class="sxs-lookup"><span data-stu-id="e943b-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="e943b-198">Ajoutez l’objet suivant en tant que membre unique du `controls` tableau.</span><span class="sxs-lookup"><span data-stu-id="e943b-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="e943b-199">À propos de ce balisage, notez les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="e943b-199">About this markup, note:</span></span>

    - <span data-ttu-id="e943b-200">Toutes les propriétés, à l’exception `enabled` de, sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="e943b-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="e943b-201">`type` Spécifie le type de contrôle.</span><span class="sxs-lookup"><span data-stu-id="e943b-201">`type` specifies the type of control.</span></span> <span data-ttu-id="e943b-202">Les valeurs peuvent être « Button », « menu » ou « MobileButton ».</span><span class="sxs-lookup"><span data-stu-id="e943b-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="e943b-203">`id` peut contenir jusqu’à 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="e943b-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="e943b-204">`actionId` doit être l’ID d’une action définie dans le `actions` tableau.</span><span class="sxs-lookup"><span data-stu-id="e943b-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="e943b-205">(Voir l’étape 1 de cette section.)</span><span class="sxs-lookup"><span data-stu-id="e943b-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="e943b-206">`label` est une chaîne conviviale à utiliser comme étiquette du bouton.</span><span class="sxs-lookup"><span data-stu-id="e943b-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="e943b-207">`superTip` représente une forme enrichie d’info-bulle.</span><span class="sxs-lookup"><span data-stu-id="e943b-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="e943b-208">Les `title` Propriétés et `description` sont toutes deux requises.</span><span class="sxs-lookup"><span data-stu-id="e943b-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="e943b-209">`icon` spécifie les icônes du bouton.</span><span class="sxs-lookup"><span data-stu-id="e943b-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="e943b-210">Les remarques précédentes sur l’icône de groupe s’appliquent également ici.</span><span class="sxs-lookup"><span data-stu-id="e943b-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="e943b-211">`enabled` (facultatif) indique si le bouton est activé lorsque l’onglet contextuel s’affiche.</span><span class="sxs-lookup"><span data-stu-id="e943b-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="e943b-212">La valeur par défaut est `true` .</span><span class="sxs-lookup"><span data-stu-id="e943b-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="e943b-213">Voici l’exemple complet de l’objet BLOB JSON :</span><span class="sxs-lookup"><span data-stu-id="e943b-213">The following is the complete example of the JSON blob:</span></span>

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
      "label": "Data",
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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="e943b-214">Enregistrer l’onglet contextuel avec Office avec requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="e943b-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="e943b-215">L’onglet contextuel est inscrit avec Office en appelant la méthode [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) .</span><span class="sxs-lookup"><span data-stu-id="e943b-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="e943b-216">Cette opération s’effectue généralement dans la fonction qui est affectée à `Office.initialize` ou avec la `Office.onReady` méthode.</span><span class="sxs-lookup"><span data-stu-id="e943b-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="e943b-217">Pour plus d’informations sur ces méthodes et l’initialisation du complément, reportez-vous à la rubrique [initialiser votre complément Office](../develop/initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="e943b-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="e943b-218">Toutefois, vous pouvez appeler la méthode à tout moment après l’initialisation.</span><span class="sxs-lookup"><span data-stu-id="e943b-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e943b-219">La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un complément.</span><span class="sxs-lookup"><span data-stu-id="e943b-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="e943b-220">Une erreur est générée si elle est encore appelée.</span><span class="sxs-lookup"><span data-stu-id="e943b-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="e943b-221">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="e943b-221">The following is an example.</span></span> <span data-ttu-id="e943b-222">Notez que la chaîne JSON doit être convertie en objet JavaScript avec la `JSON.parse` méthode avant de pouvoir être transmise à une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e943b-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="e943b-223">Spécifier les contextes lorsque l’onglet est visible avec requestUpdate</span><span class="sxs-lookup"><span data-stu-id="e943b-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="e943b-224">En règle générale, un onglet contextuel personnalisé doit s’afficher lorsqu’un événement initié par l’utilisateur modifie le contexte du complément.</span><span class="sxs-lookup"><span data-stu-id="e943b-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="e943b-225">Imaginez un scénario dans lequel l’onglet doit être visible lorsque, et uniquement lorsqu’un graphique (de la feuille de calcul par défaut d’un classeur Excel) est activé.</span><span class="sxs-lookup"><span data-stu-id="e943b-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="e943b-226">Commencez par attribuer des gestionnaires.</span><span class="sxs-lookup"><span data-stu-id="e943b-226">Begin by assigning handlers.</span></span> <span data-ttu-id="e943b-227">Cette opération est généralement exécutée dans la `Office.onReady` méthode, comme dans l’exemple suivant, qui affecte des gestionnaires (créés à une étape ultérieure) aux `onActivated` `onDeactivated` événements et de tous les graphiques de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e943b-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

<span data-ttu-id="e943b-228">Ensuite, définissez les gestionnaires.</span><span class="sxs-lookup"><span data-stu-id="e943b-228">Next, define the handlers.</span></span> <span data-ttu-id="e943b-229">Voici un exemple simple de a `showDataTab` , mais voir [la gestion de l’erreur HostRestartNeeded](#handling-the-hostrestartneeded-error) plus loin dans cet article pour obtenir une version plus robuste de la fonction.</span><span class="sxs-lookup"><span data-stu-id="e943b-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handling-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="e943b-230">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="e943b-230">About this code, note:</span></span>

- <span data-ttu-id="e943b-231">Office effectue un contrôle lorsqu’il met à jour l’état du ruban.</span><span class="sxs-lookup"><span data-stu-id="e943b-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="e943b-232">La méthode  [Office. Ribbon. requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) met en file d’attente une requête à mettre à jour.</span><span class="sxs-lookup"><span data-stu-id="e943b-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="e943b-233">La méthode permet de résoudre l’objet dès qu' `Promise` il a mis en file d’attente la demande, et non lors de la mise à jour du ruban.</span><span class="sxs-lookup"><span data-stu-id="e943b-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="e943b-234">Le paramètre de la `requestUpdate` méthode est un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID *exactement comme spécifié dans JSON* et (2) indique la visibilité de l’onglet.</span><span class="sxs-lookup"><span data-stu-id="e943b-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="e943b-235">Si vous avez plusieurs onglets contextuels personnalisés qui doivent être visibles dans le même contexte, il vous suffit d’ajouter des objets Tab supplémentaires au `tabs` tableau.</span><span class="sxs-lookup"><span data-stu-id="e943b-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="e943b-236">Le gestionnaire permettant de masquer l’onglet est quasiment identique, à la différence qu’il rétablit la `visible` propriété sur `false` .</span><span class="sxs-lookup"><span data-stu-id="e943b-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="e943b-237">La bibliothèque JavaScript Office fournit également plusieurs interfaces (types) pour faciliter la création de l' `RibbonUpdateData` objet.</span><span class="sxs-lookup"><span data-stu-id="e943b-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="e943b-238">Voici la `showDataTab` fonction dans la machine à écrire et elle utilise ces types.</span><span class="sxs-lookup"><span data-stu-id="e943b-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="e943b-239">Activer/désactiver la visibilité de l’onglet et l’état activé d’un bouton en même temps</span><span class="sxs-lookup"><span data-stu-id="e943b-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="e943b-240">La `requestUpdate` méthode est également utilisée pour faire basculer l’état activé ou désactivé d’un bouton personnalisé dans un onglet contextuel personnalisé ou un onglet de base personnalisé. Pour plus d’informations à ce sujet, consultez la rubrique [activer et désactiver des commandes de complément](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="e943b-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="e943b-241">Il peut y avoir des scénarios dans lesquels vous souhaitez modifier à la fois la visibilité d’un onglet et l’état activé d’un bouton en même temps.</span><span class="sxs-lookup"><span data-stu-id="e943b-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="e943b-242">Vous pouvez effectuer cette opération à l’aide d’un seul appel de `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="e943b-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="e943b-243">Voici un exemple dans lequel un bouton d’un onglet principal est activé en même temps qu’un onglet contextuel est rendu visible.</span><span class="sxs-lookup"><span data-stu-id="e943b-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

<span data-ttu-id="e943b-244">Dans l’exemple suivant, le bouton activé se trouve sur le même onglet contextuel qui est rendu visible.</span><span class="sxs-lookup"><span data-stu-id="e943b-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="localizing-the-json-blob"></a><span data-ttu-id="e943b-245">Localisation de l’objet BLOB JSON</span><span class="sxs-lookup"><span data-stu-id="e943b-245">Localizing the JSON blob</span></span>

<span data-ttu-id="e943b-246">L’objet BLOB JSON transmis à `requestCreateControls` n’est pas localisé de la même façon que le balisage de manifeste pour les onglets principaux personnalisés est localisé (ce qui est décrit lors [de la localisation du contrôle à partir du manifeste](../develop/localization.md#control-localization-from-the-manifest)).</span><span class="sxs-lookup"><span data-stu-id="e943b-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="e943b-247">Au lieu de cela, la localisation doit être effectuée au moment de l’exécution à l’aide d’objets BLOB JSON distincts pour chaque paramètre régional.</span><span class="sxs-lookup"><span data-stu-id="e943b-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="e943b-248">Nous vous suggérons d’utiliser une `switch` instruction qui teste la propriété [Office. Context. displayLanguage](/javascript/api/office/office.context#displayLanguage) .</span><span class="sxs-lookup"><span data-stu-id="e943b-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="e943b-249">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="e943b-249">The following is an example:</span></span>

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
                          "label": "Data",
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
                          "label": "Données",
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

<span data-ttu-id="e943b-250">Votre code appelle ensuite la fonction pour obtenir l’objet BLOB localisé qui est transmis `requestCreateControls` , comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="e943b-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a><span data-ttu-id="e943b-251">Gestion de l’erreur HostRestartNeeded</span><span class="sxs-lookup"><span data-stu-id="e943b-251">Handling the HostRestartNeeded error</span></span>

<span data-ttu-id="e943b-252">Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="e943b-252">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="e943b-253">Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau.</span><span class="sxs-lookup"><span data-stu-id="e943b-253">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="e943b-254">La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué.</span><span class="sxs-lookup"><span data-stu-id="e943b-254">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="e943b-255">Voici comment vous pouvez gérer cette erreur.</span><span class="sxs-lookup"><span data-stu-id="e943b-255">The following is an example of how to handle this error.</span></span> <span data-ttu-id="e943b-256">Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e943b-256">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
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
