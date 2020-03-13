---
title: Commandes Activé et Désactivé pour les compléments
description: Découvrez la modification de l'état Activé ou Désactivé des boutons de rubans et des éléments de menu personnalisés dans votre complément web Office.
ms.date: 03/09/2020
localization_priority: Priority
ms.openlocfilehash: dbe895a121a5d10d687c9a599b85234ae62919f5
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596682"
---
# <a name="enable-and-disable-add-in-commands-preview"></a><span data-ttu-id="fc4e8-103">Commandes Activé et Désactivé pour les compléments (préversion)</span><span class="sxs-lookup"><span data-stu-id="fc4e8-103">Enable and Disable Add-in Commands (preview)</span></span>

<span data-ttu-id="fc4e8-104">Lorsque seulement quelques fonctionnalités de votre complément doivent être disponibles dans certains contextes, vous pouvez activer ou désactiver vos commandes de complément personnalisées par programme.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="fc4e8-105">Par exemple, une fonction qui modifie l’en-tête d’un tableau doit être uniquement activée lorsque le curseur se trouve dans un tableau.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="fc4e8-106">Vous pouvez également préciser si la commande est activée ou désactivée lorsque l’application hôte Office s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-106">You can also specify whether the command is enabled or disabled when the Office host application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="fc4e8-107">Cet article suppose que vous connaissez la documentation décrite ci-après.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="fc4e8-108">Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).</span><span class="sxs-lookup"><span data-stu-id="fc4e8-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> [<span data-ttu-id="fc4e8-109">Concepts basiques pour les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="fc4e8-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="preview-status"></a><span data-ttu-id="fc4e8-110">État de préversion</span><span class="sxs-lookup"><span data-stu-id="fc4e8-110">Preview status</span></span>

<span data-ttu-id="fc4e8-111">Les API décrites dans cet article sont en préversion et ne sont actuellement disponibles que dans Excel.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-111">The APIs described in this article are in preview and are currently only available in Excel.</span></span>

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="rules-and-gotchas"></a><span data-ttu-id="fc4e8-112">Règles et pièges</span><span class="sxs-lookup"><span data-stu-id="fc4e8-112">Rules and gotchas</span></span>

### <a name="single-line-ribbon-in-office-on-the-web"></a><span data-ttu-id="fc4e8-113">Ruban d'une seule ligne dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="fc4e8-113">Single-line ribbon in Office on the web</span></span>

<span data-ttu-id="fc4e8-114">Les API et balisages de manifeste décrits dans cet article s’appliquent uniquement au ruban d'une seule ligne dans Office sur le web.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-114">In Office on the web, the APIs and manifest markup described in this article only affect the single-line ribbon.</span></span> <span data-ttu-id="fc4e8-115">Ils n’ont aucun effet sur le ruban multi-ligne.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-115">They have no effect on the multiline ribbon.</span></span> <span data-ttu-id="fc4e8-116">Ils ont un effet sur les deux rubans dans la version de bureau d’Office.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-116">They affect both ribbons for desktop Office.</span></span> <span data-ttu-id="fc4e8-117">Pour plus d’informations sur les deux rubans, voir [Utiliser le ruban simplifié](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span><span class="sxs-lookup"><span data-stu-id="fc4e8-117">For more information about the two ribbons, see [Use the simplified ribbon](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span></span>

### <a name="shared-runtime-required"></a><span data-ttu-id="fc4e8-118">Runtime partagé requis</span><span class="sxs-lookup"><span data-stu-id="fc4e8-118">Shared runtime required</span></span>

<span data-ttu-id="fc4e8-119">Les API et balisages de manifeste décrits dans cet article pour lesquels le complément de manifeste spécifie qu’il doit utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-119">The APIs and manifest markup described in this article that the add-in's manifest specifies that it should use a shared runtime.</span></span> <span data-ttu-id="fc4e8-120">Pour ce faire, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-120">To do this take the following steps.</span></span>

1. <span data-ttu-id="fc4e8-121">Dans l'élément [Runtimes du manifeste](../reference/manifest/runtimes.md), ajoutez l’élément enfant suivant : `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-121">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="fc4e8-122">(s’il n’y a pas encore d’élément `<Runtimes>` dans le manifeste, créez-le en tant que premier enfant sous l’élément `<Host>` dans la section `VersionOverrides`.)</span><span class="sxs-lookup"><span data-stu-id="fc4e8-122">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="fc4e8-123">Dans la section [Resources.Urls](../reference/manifest/resources.md) du manifeste, ajoutez l’élément enfant suivant : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, où `{MyDomain}` est le domaine du complément et `{path-to-start-page}` le chemin d’accès de la page de démarrage du complément. par exemple : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-123">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="fc4e8-124">En fonction du contenu de votre complément, à savoir un volet des tâches, un fichier de fonctions ou une fonction Excel personnalisée, vous devez effectuer au moins une des trois étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-124">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="fc4e8-125">Si le complément contient un volet des tâches, configurez l'attribut `resid` de l’[Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md), élément de `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-125">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="fc4e8-126">Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-126">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="fc4e8-127">Si le complément contient une fonction Excel personnalisée, configurez l'attribut `resid` de [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md), élément de `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-127">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="fc4e8-128">Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-128">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="fc4e8-129">Si le complément contient un fichier de fonction, configurez l’attribut `resid` de [FunctionFile](../reference/manifest/functionfile.md), élément de `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-129">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="fc4e8-130">Le fichier doit ressembler à ceci : `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-130">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="fc4e8-131">Configurer l'état par défaut sur désactivé</span><span class="sxs-lookup"><span data-stu-id="fc4e8-131">Set the default state to disabled</span></span>

<span data-ttu-id="fc4e8-132">Les commandes de complément sont activées par défaut au démarrage de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-132">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="fc4e8-133">Si vous souhaitez qu’un bouton ou un élément de menu personnalisé soit désactivé au démarrage de l’application Office, vous devez le spécifier dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-133">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="fc4e8-134">Il vous suffit d’ajouter l'élément [activé](../reference/manifest/enabled.md) (avec la valeur `false`) juste en dessous de l'élément [action](../reference/manifest/action.md) dans la déclaration du contrôle.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-134">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately below the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="fc4e8-135">La structure basique est la suivante :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-135">The following shows the basic structure:</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a><span data-ttu-id="fc4e8-136">Modifier l’état par programme</span><span class="sxs-lookup"><span data-stu-id="fc4e8-136">Change the state programmatically</span></span>

<span data-ttu-id="fc4e8-137">Les principales étapes pour modifier l’état activé d’une commande de complément sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-137">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="fc4e8-138">Créez un objet [RibbonUpdaterData](/javascript/api/office-runtime/officeruntime.ribbonupdaterdata) qui (1) spécifie la commande et son onglet parent, selon leur ID, comme spécifié dans le manifeste. et (2) indique l’état activé ou désactivé de la commande.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-138">Create a [RibbonUpdaterData](/javascript/api/office-runtime/officeruntime.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="fc4e8-139">Passez l’objet **RibbonUpdaterData** à la méthode [OfficeRuntime.Ribbon.requestUpdate ()](/javascript/api/office-runtime/officeruntime.ribbon#requestupdate-input-).</span><span class="sxs-lookup"><span data-stu-id="fc4e8-139">Pass the **RibbonUpdaterData** object to the [OfficeRuntime.Ribbon.requestUpdate()](/javascript/api/office-runtime/officeruntime.ribbon#requestupdate-input-) method.</span></span>

<span data-ttu-id="fc4e8-140">Voici un exemple simple.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-140">The following is a simple example.</span></span> <span data-ttu-id="fc4e8-141">Veuillez noter que « MyButton » et « OfficeAddinTab1 » sont copiés à partir du manifeste.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-141">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

```javascript
function enableButton() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) {
            ribbon.requestUpdate({
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
        });
}
```

> [!NOTE]
> <span data-ttu-id="fc4e8-142">Nous envisageons provisoirement de simplifier les API en avril 2020, de deux façons :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-142">We tentatively plan to simplify the APIs in April, 2020, in two ways:</span></span>
>
> - <span data-ttu-id="fc4e8-143">Les API sont déplacées de l’espace de noms `OfficeRuntime` vers l’espace de noms `Office`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-143">The APIs will move from the `OfficeRuntime` namespace to the `Office` namespace.</span></span>
> - <span data-ttu-id="fc4e8-144">Vous n’avez pas besoin d’appeler une méthode `getRibbon()`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-144">You will not need to call a `getRibbon()` method.</span></span> <span data-ttu-id="fc4e8-145">L’objet `Ribbon` sera une propriété singleton de l’objet `Office`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-145">The `Ribbon` object will be a singleton property of the `Office` object.</span></span>
>
> <span data-ttu-id="fc4e8-146">Par exemple, le code antérieur sera réécrit comme suit :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-146">For example, the preceding code would be rewritten as follows:</span></span>
>
> ```javascript
> function enableButton() {
>    Office.ribbon.requestUpdate({
>        tabs: [
>            {
>                id: "OfficeAppTab1", 
>                controls: [
>                {
>                    id: "MyButton", 
>                    enabled: true
>                }
>            ]}
>        ]});
> }
> ```

<span data-ttu-id="fc4e8-147">Nous proposons également plusieurs interfaces (types) pour faciliter la construction de l’objet **RibbonUpdateData**.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-147">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="fc4e8-148">L’exemple suivant est l’équivalent de TypeScript et il utilise ces types.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-148">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    const ribbon: Ribbon = await OfficeRuntime.ui.getRibbon();
    await ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="fc4e8-149">Office effectue un contrôle lorsqu’il met à jour l’état du ruban.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-149">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="fc4e8-150">La méthode **requestUpdate()** met en file d’attente une demande de mise à jour.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-150">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="fc4e8-151">La méthode permet de résoudre l’objet Promesse dès la mise en file d’attente de la demande, et non lors la prochaine mise à jour du ruban.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-151">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="fc4e8-152">Modifier l’état en réponse à un événement</span><span class="sxs-lookup"><span data-stu-id="fc4e8-152">Change the state in response to an event</span></span>

<span data-ttu-id="fc4e8-153">Un scénario courant est celui lors duquel l’état du ruban peut être modifié lorsqu’un événement initié par l’utilisateur modifie le contexte du complément.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-153">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="fc4e8-154">Imaginez un scénario dans lequel un bouton doit être activé lorsque, et seulement lorsqu'un graphique est activé.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-154">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="fc4e8-155">La première étape consiste à définir l'élément [Activé](../reference/manifest/enabled.md) pour le bouton dans le manifeste `false`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-155">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="fc4e8-156">Voir l'exemple ci-dessus.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-156">See above for an example.</span></span>

<span data-ttu-id="fc4e8-157">Deuxièmement, assignez des gestionnaires.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-157">Second, assign handlers.</span></span> <span data-ttu-id="fc4e8-158">Cette procédure est généralement effectuée dans la méthode **Office.onReady** comme illustré dans l’exemple suivant qui assigne des gestionnaires (créés dans une étape ultérieure) aux évènements **onActivated** et **onDeactivated** de tous les graphiques de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-158">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

<span data-ttu-id="fc4e8-159">Troisièmement, définissez le gestionnaire `enableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-159">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="fc4e8-160">Voici un exemple simple, mais consultez les **Pratiques recommandées : test pour les erreurs de contrôle d’état** ci-dessous pour modifier l’état d’un contrôle de façon plus efficace.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-160">The following is a simple example, but see **Best practice: Test for control status errors** below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) {
            var button = {id: "ChartFormatButton", enabled: true};
            var parentTab = {id: "CustomChartTab", controls: [button]};
            var ribbonUpdater = {tabs: [parentTab]};
            await ribbon.requestUpdate(ribbonUpdater);
        });
}
```

<span data-ttu-id="fc4e8-161">Quatrièmement, définissez le gestionnaire `disableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-161">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="fc4e8-162">Il est identique à `enableChartFormat`, sauf que la propriété **activé** de l’objet bouton a la valeur `false`.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-162">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="fc4e8-163">Pratiques recommandées : test pour les erreurs de contrôle d'état</span><span class="sxs-lookup"><span data-stu-id="fc4e8-163">Best practice: Test for control status errors</span></span>

<span data-ttu-id="fc4e8-164">Le ruban ne se redessine pas, dans certains cas, une fois que `requestUpdate` est appelé, de sorte que l’état du contrôle cliquable ne change pas.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-164">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="fc4e8-165">Il est pour cette raison recommandé de suivre l'état des contrôles du complément.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-165">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="fc4e8-166">Le complément doit respecter les règles suivantes :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-166">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="fc4e8-167">Lorsque `requestUpdate` est appelé, le code doit enregistrer l’état prévu des boutons et éléments de menu personnalisés.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-167">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="fc4e8-168">Lorsque l’utilisateur clique sur un contrôle personnalisé, le premier code dans le gestionnaire doit vérifier si le bouton aurait dû être cliquable.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-168">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="fc4e8-169">Dans la négative, le code doit signaler une erreur ou consigner une erreur et réessayer de définir les boutons de l'état prévu.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-169">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="fc4e8-170">L’exemple suivant présente une fonction qui désactive un bouton et enregistre l’état du bouton.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-170">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="fc4e8-171">Veuillez noter que `chartFormatButtonEnabled` est une variable Boolean globale qui est initialisée sur la même valeur que l'élément [Activé](../reference/manifest/enabled.md) pour le bouton dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-171">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) {
            var button = {id: "ChartFormatButton", enabled: false};
            var parentTab = {id: "CustomChartTab", controls: [button]};
            var ribbonUpdater = {tabs: [parentTab]};
            await ribbon.requestUpdate(ribbonUpdater);

            chartFormatButtonEnabled = false;
        });
}
```

<span data-ttu-id="fc4e8-172">L’exemple suivant présente la façon dont le gestionnaire du bouton vérifie l’état d’un bouton incorrect.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-172">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="fc4e8-173">Veuillez noter que `reportError` est une fonction qui affiche ou consigne une erreur.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-173">Note that `reportError` is a function that shows or logs an error.</span></span>

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a><span data-ttu-id="fc4e8-174">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="fc4e8-174">Error handling</span></span>

<span data-ttu-id="fc4e8-175">Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-175">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="fc4e8-176">Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-176">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="fc4e8-177">La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-177">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="fc4e8-178">Voici comment vous pouvez gérer cette erreur.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-178">The following is an example of how to handle this error.</span></span> <span data-ttu-id="fc4e8-179">Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fc4e8-179">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function disableChartFormat() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) {
            var button = {id: "ChartFormatButton", enabled: false};
            var parentTab = {id: "CustomChartTab", controls: [button]};
            var ribbonUpdater = {tabs: [parentTab]};
            await ribbon.requestUpdate(ribbonUpdater);

            chartFormatButtonEnabled = false;
        })
        .catch(function (error){
            if (error.code == "HostRestartNeeded"){
                reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
            }
        });
}
```

## <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="fc4e8-180">Effectuez un test pour la prise en charge des plateformes avec les ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="fc4e8-180">Test for platform support with requirement sets</span></span>

<span data-ttu-id="fc4e8-p123">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="fc4e8-p123">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="fc4e8-184">Les API activer/désactiver nécessitent la prise en charge des ensembles de configuration suivants :</span><span class="sxs-lookup"><span data-stu-id="fc4e8-184">The enable/disable APIs require support of the following requirement sets:</span></span>

- [<span data-ttu-id="fc4e8-185">AddinCommands 1.1</span><span class="sxs-lookup"><span data-stu-id="fc4e8-185">AddinCommands 1.1</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
