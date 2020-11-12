---
title: Commandes Activé et Désactivé pour les compléments
description: Découvrez la modification de l'état Activé ou Désactivé des boutons de rubans et des éléments de menu personnalisés dans votre complément web Office.
ms.date: 11/07/2020
localization_priority: Normal
ms.openlocfilehash: 7a9994ae25285c876236879e65861ee3cc59f7e5
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996388"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="77a61-103">Commandes Activé et Désactivé pour les compléments</span><span class="sxs-lookup"><span data-stu-id="77a61-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="77a61-104">Lorsque seulement quelques fonctionnalités de votre complément doivent être disponibles dans certains contextes, vous pouvez activer ou désactiver vos commandes de complément personnalisées par programme.</span><span class="sxs-lookup"><span data-stu-id="77a61-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="77a61-105">Par exemple, une fonction qui modifie l’en-tête d’un tableau doit être uniquement activée lorsque le curseur se trouve dans un tableau.</span><span class="sxs-lookup"><span data-stu-id="77a61-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="77a61-106">Vous pouvez également spécifier si la commande est activée ou désactivée lorsque l’application cliente Office s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="77a61-106">You can also specify whether the command is enabled or disabled when the Office client application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="77a61-107">Cet article suppose que vous connaissez la documentation décrite ci-après.</span><span class="sxs-lookup"><span data-stu-id="77a61-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="77a61-108">Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).</span><span class="sxs-lookup"><span data-stu-id="77a61-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="77a61-109">Concepts basiques pour les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="77a61-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a><span data-ttu-id="77a61-110">Prise en charge de l’application et de la plateforme Office uniquement</span><span class="sxs-lookup"><span data-stu-id="77a61-110">Office application and platform support only</span></span>

<span data-ttu-id="77a61-111">Les API décrites dans cet article sont disponibles uniquement dans Excel, et uniquement dans Office sous Windows, Office sur Mac et Office sur le Web.</span><span class="sxs-lookup"><span data-stu-id="77a61-111">The APIs described in this article are only available in Excel, and only in Office on Windows, Office on Mac, and Office on the web.</span></span>

### <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="77a61-112">Effectuez un test pour la prise en charge des plateformes avec les ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="77a61-112">Test for platform support with requirement sets</span></span>

<span data-ttu-id="77a61-113">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="77a61-113">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="77a61-114">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une combinaison d’applications Office et de plateformes prend en charge les API dont un complément a besoin.</span><span class="sxs-lookup"><span data-stu-id="77a61-114">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application and platform combination supports APIs that an add-in needs.</span></span> <span data-ttu-id="77a61-115">Pour plus d’informations, consultez la rubrique [versions d’Office et ensembles de conditions requises](../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="77a61-115">For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="77a61-116">Les API d’activation/de désactivation appartiennent à l’ensemble de conditions requises [RibbonApi 1,1](../reference/requirement-sets/ribbon-api-requirement-sets.md) .</span><span class="sxs-lookup"><span data-stu-id="77a61-116">The enable/disable APIs belong to the [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) requirement set.</span></span>

> [!NOTE]
> <span data-ttu-id="77a61-117">L’ensemble de conditions requises **RibbonApi 1,1** n’étant pas encore pris en charge dans le manifeste, vous ne pouvez pas le spécifier dans la section du manifeste `<Requirements>` .</span><span class="sxs-lookup"><span data-stu-id="77a61-117">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span> <span data-ttu-id="77a61-118">Pour tester la prise en charge, votre code doit appeler `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` .</span><span class="sxs-lookup"><span data-stu-id="77a61-118">To test for support, your code should call `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`.</span></span> <span data-ttu-id="77a61-119">Si, *et seulement si* , cet appel `true` est renvoyé, votre code peut appeler les API activer/désactiver.</span><span class="sxs-lookup"><span data-stu-id="77a61-119">If, *and only if* , that call returns `true`, your code can call the enable/disable APIs.</span></span> <span data-ttu-id="77a61-120">Si l’appel de la `isSetSupported` méthode retournée `false` est activé, toutes les commandes de complément personnalisées sont activées en totalité.</span><span class="sxs-lookup"><span data-stu-id="77a61-120">If the call of `isSetSupported` returns `false`, then all custom add-in commands are enabled all of the time.</span></span> <span data-ttu-id="77a61-121">Vous devez concevoir votre complément de production, ainsi que toutes les instructions dans l’application, pour tenir compte de la façon dont il fonctionnera lorsque l’ensemble de conditions requises **RibbonApi 1,1** n’est pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="77a61-121">You must design your production add-in, and any in-app instructions, to take account of how it will work when the **RibbonApi 1.1** requirement set is not supported.</span></span> <span data-ttu-id="77a61-122">Pour plus d’informations et des exemples d’utilisation `isSetSupported` , reportez-vous à la rubrique [spécifier les applications Office et les conditions requises](../develop/specify-office-hosts-and-api-requirements.md)de l’API, notamment [utiliser les vérifications d’exécution dans votre code JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="77a61-122">For more information and examples of using `isSetSupported`, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md), especially [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="77a61-123">(La section [définir l’élément Requirements dans le manifeste](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) de cet article ne s’applique pas au ruban 1,1.)</span><span class="sxs-lookup"><span data-stu-id="77a61-123">(The section [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)</span></span>

## <a name="shared-runtime-required"></a><span data-ttu-id="77a61-124">Runtime partagé requis</span><span class="sxs-lookup"><span data-stu-id="77a61-124">Shared runtime required</span></span>

<span data-ttu-id="77a61-125">Les API et balisages de manifeste décrits dans cet article exigent que le manifeste du complément spécifie la nécessité d’utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="77a61-125">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="77a61-126">Pour ce faire, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="77a61-126">To do this take the following steps.</span></span>

1. <span data-ttu-id="77a61-127">Dans l'élément [Runtimes du manifeste](../reference/manifest/runtimes.md), ajoutez l’élément enfant suivant : `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span><span class="sxs-lookup"><span data-stu-id="77a61-127">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="77a61-128">(s’il n’y a pas encore d’élément `<Runtimes>` dans le manifeste, créez-le en tant que premier enfant sous l’élément `<Host>` dans la section `VersionOverrides`.)</span><span class="sxs-lookup"><span data-stu-id="77a61-128">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="77a61-129">Dans la section [Resources.Urls](../reference/manifest/resources.md) du manifeste, ajoutez l’élément enfant suivant : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, où `{MyDomain}` est le domaine du complément et `{path-to-start-page}` le chemin d’accès de la page de démarrage du complément. par exemple : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span><span class="sxs-lookup"><span data-stu-id="77a61-129">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="77a61-130">En fonction du contenu de votre complément, à savoir un volet des tâches, un fichier de fonctions ou une fonction Excel personnalisée, vous devez effectuer au moins une des trois étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="77a61-130">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="77a61-131">Si le complément contient un volet Office, définissez l'attribut `resid` de l’élément [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) sur la même chaîne que celle que vous avez utilisée pour le `resid` de l’élément `<Runtime>` à l’étape 1 ; par exemple, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="77a61-131">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="77a61-132">Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="77a61-132">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="77a61-133">Si le complément contient une fonction personnalisée Excel, définissez l'attribut `resid` de l’élément [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) sur la même chaîne que celle que vous avez utilisée pour le `resid` de l’élément `<Runtime>` à l’étape 1 ; par exemple, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="77a61-133">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="77a61-134">Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="77a61-134">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="77a61-135">Si le complément contient un fichier de fonctions, définissez l'attribut `resid` de l’élément [FunctionFile](../reference/manifest/functionfile.md) sur la même chaîne que celle que vous avez utilisée pour le `resid` de l’élément `<Runtime>` à l’étape 1 ; par exemple, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="77a61-135">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="77a61-136">Le fichier doit ressembler à ceci : `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="77a61-136">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="77a61-137">Configurer l'état par défaut sur désactivé</span><span class="sxs-lookup"><span data-stu-id="77a61-137">Set the default state to disabled</span></span>

<span data-ttu-id="77a61-138">Les commandes de complément sont activées par défaut au démarrage de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="77a61-138">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="77a61-139">Si vous souhaitez qu’un bouton ou un élément de menu personnalisé soit désactivé au démarrage de l’application Office, vous devez le spécifier dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="77a61-139">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="77a61-140">Il vous suffit d’ajouter un élément [activé](../reference/manifest/enabled.md) (avec la valeur `false`) juste *au-dessous* (non à l’intérieur) de l'élément [Action](../reference/manifest/action.md) dans la déclaration du contrôle.</span><span class="sxs-lookup"><span data-stu-id="77a61-140">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="77a61-141">La structure basique est la suivante :</span><span class="sxs-lookup"><span data-stu-id="77a61-141">The following shows the basic structure:</span></span>

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
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a><span data-ttu-id="77a61-142">Modifier l’état par programme</span><span class="sxs-lookup"><span data-stu-id="77a61-142">Change the state programmatically</span></span>

<span data-ttu-id="77a61-143">Les principales étapes pour modifier l’état activé d’une commande de complément sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="77a61-143">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="77a61-144">Créez un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie la commande et son onglet parent, selon leur ID, comme spécifié dans le manifeste. et (2) indique l’état activé ou désactivé de la commande.</span><span class="sxs-lookup"><span data-stu-id="77a61-144">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="77a61-145">Transmettez l’objet **RibbonUpdaterData** à la méthode [Office.ribbon.requestUpdate ()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-).</span><span class="sxs-lookup"><span data-stu-id="77a61-145">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method.</span></span>

<span data-ttu-id="77a61-146">Voici un exemple simple.</span><span class="sxs-lookup"><span data-stu-id="77a61-146">The following is a simple example.</span></span> <span data-ttu-id="77a61-147">Veuillez noter que « MyButton » et « OfficeAddinTab1 » sont copiés à partir du manifeste.</span><span class="sxs-lookup"><span data-stu-id="77a61-147">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
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

<span data-ttu-id="77a61-148">Nous proposons également plusieurs interfaces (types) pour faciliter la construction de l’objet **RibbonUpdateData**.</span><span class="sxs-lookup"><span data-stu-id="77a61-148">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="77a61-149">L’exemple suivant est l’équivalent de TypeScript et il utilise ces types.</span><span class="sxs-lookup"><span data-stu-id="77a61-149">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="77a61-150">Office effectue un contrôle lorsqu’il met à jour l’état du ruban.</span><span class="sxs-lookup"><span data-stu-id="77a61-150">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="77a61-151">La méthode **requestUpdate()** met en file d’attente une demande de mise à jour.</span><span class="sxs-lookup"><span data-stu-id="77a61-151">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="77a61-152">La méthode permet de résoudre l’objet Promesse dès la mise en file d’attente de la demande, et non lors la prochaine mise à jour du ruban.</span><span class="sxs-lookup"><span data-stu-id="77a61-152">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="77a61-153">Modifier l’état en réponse à un événement</span><span class="sxs-lookup"><span data-stu-id="77a61-153">Change the state in response to an event</span></span>

<span data-ttu-id="77a61-154">Un scénario courant est celui lors duquel l’état du ruban peut être modifié lorsqu’un événement initié par l’utilisateur modifie le contexte du complément.</span><span class="sxs-lookup"><span data-stu-id="77a61-154">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="77a61-155">Imaginez un scénario dans lequel un bouton doit être activé lorsque, et seulement lorsqu'un graphique est activé.</span><span class="sxs-lookup"><span data-stu-id="77a61-155">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="77a61-156">La première étape consiste à définir l'élément [Activé](../reference/manifest/enabled.md) pour le bouton dans le manifeste `false`.</span><span class="sxs-lookup"><span data-stu-id="77a61-156">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="77a61-157">Voir l'exemple ci-dessus.</span><span class="sxs-lookup"><span data-stu-id="77a61-157">See above for an example.</span></span>

<span data-ttu-id="77a61-158">Deuxièmement, assignez des gestionnaires.</span><span class="sxs-lookup"><span data-stu-id="77a61-158">Second, assign handlers.</span></span> <span data-ttu-id="77a61-159">Cette procédure est généralement effectuée dans la méthode **Office.onReady** comme illustré dans l’exemple suivant qui assigne des gestionnaires (créés dans une étape ultérieure) aux évènements **onActivated** et **onDeactivated** de tous les graphiques de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="77a61-159">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

<span data-ttu-id="77a61-160">Troisièmement, définissez le gestionnaire `enableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="77a61-160">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="77a61-161">Voici un exemple simple, mais consultez les [Pratiques recommandées : test pour les erreurs de contrôle d’état](#best-practice-test-for-control-status-errors) ci-dessous pour modifier l’état d’un contrôle de façon plus efficace.</span><span class="sxs-lookup"><span data-stu-id="77a61-161">The following is a simple example, but see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors) below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="77a61-162">Quatrièmement, définissez le gestionnaire `disableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="77a61-162">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="77a61-163">Il est identique à `enableChartFormat`, sauf que la propriété **activé** de l’objet bouton a la valeur `false`.</span><span class="sxs-lookup"><span data-stu-id="77a61-163">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="77a61-164">Pratiques recommandées : test pour les erreurs de contrôle d'état</span><span class="sxs-lookup"><span data-stu-id="77a61-164">Best practice: Test for control status errors</span></span>

<span data-ttu-id="77a61-165">Le ruban ne se redessine pas, dans certains cas, une fois que `requestUpdate` est appelé, de sorte que l’état du contrôle cliquable ne change pas.</span><span class="sxs-lookup"><span data-stu-id="77a61-165">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="77a61-166">Il est pour cette raison recommandé de suivre l'état des contrôles du complément.</span><span class="sxs-lookup"><span data-stu-id="77a61-166">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="77a61-167">Le complément doit respecter les règles suivantes :</span><span class="sxs-lookup"><span data-stu-id="77a61-167">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="77a61-168">Lorsque `requestUpdate` est appelé, le code doit enregistrer l’état prévu des boutons et éléments de menu personnalisés.</span><span class="sxs-lookup"><span data-stu-id="77a61-168">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="77a61-169">Lorsque l’utilisateur clique sur un contrôle personnalisé, le premier code dans le gestionnaire doit vérifier si le bouton aurait dû être cliquable.</span><span class="sxs-lookup"><span data-stu-id="77a61-169">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="77a61-170">Dans la négative, le code doit signaler une erreur ou consigner une erreur et réessayer de définir les boutons de l'état prévu.</span><span class="sxs-lookup"><span data-stu-id="77a61-170">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="77a61-171">L’exemple suivant présente une fonction qui désactive un bouton et enregistre l’état du bouton.</span><span class="sxs-lookup"><span data-stu-id="77a61-171">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="77a61-172">Veuillez noter que `chartFormatButtonEnabled` est une variable Boolean globale qui est initialisée sur la même valeur que l'élément [Activé](../reference/manifest/enabled.md) pour le bouton dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="77a61-172">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="77a61-173">L’exemple suivant présente la façon dont le gestionnaire du bouton vérifie l’état d’un bouton incorrect.</span><span class="sxs-lookup"><span data-stu-id="77a61-173">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="77a61-174">Veuillez noter que `reportError` est une fonction qui affiche ou consigne une erreur.</span><span class="sxs-lookup"><span data-stu-id="77a61-174">Note that `reportError` is a function that shows or logs an error.</span></span>

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a><span data-ttu-id="77a61-175">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="77a61-175">Error handling</span></span>

<span data-ttu-id="77a61-176">Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="77a61-176">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="77a61-177">Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau.</span><span class="sxs-lookup"><span data-stu-id="77a61-177">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="77a61-178">La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué.</span><span class="sxs-lookup"><span data-stu-id="77a61-178">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="77a61-179">Voici comment vous pouvez gérer cette erreur.</span><span class="sxs-lookup"><span data-stu-id="77a61-179">The following is an example of how to handle this error.</span></span> <span data-ttu-id="77a61-180">Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="77a61-180">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function disableChartFormat() {
    try {
        var button = {id: "ChartFormatButton", enabled: false};
        var parentTab = {id: "CustomChartTab", controls: [button]};
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
