---
title: Ajout d’une prise en charge mobile pour un complément Outlook
description: L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 2e4ff53b371fdf50ddca43142cb5a036cfc96b25
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166144"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a><span data-ttu-id="dacd0-103">Ajouter la prise en charge des commandes de complément pour Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="dacd0-103">Add support for add-in commands for Outlook Mobile</span></span>

<span data-ttu-id="dacd0-104">L’utilisation de commandes de complément dans Outlook Mobile permet à vos utilisateurs d’accéder aux mêmes fonctionnalités (avec certaines [limitations](#code-considerations)) dont ils disposent déjà dans Outlook sur le Web, Windows et Mac.</span><span class="sxs-lookup"><span data-stu-id="dacd0-104">Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="dacd0-105">L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.</span><span class="sxs-lookup"><span data-stu-id="dacd0-105">Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.</span></span>

## <a name="updating-the-manifest"></a><span data-ttu-id="dacd0-106">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="dacd0-106">Updating the manifest</span></span>

<span data-ttu-id="dacd0-p102">La première étape de l’activation des commandes de complément dans Outlook Mobile est de les définir dans le manifeste du complément. Le schéma [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 définit un nouveau facteur de forme pour les versions mobiles, [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="dacd0-p102">The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span></span>

<span data-ttu-id="dacd0-p103">Cet élément contient toutes les informations pour charger le complément dans des clients mobiles. Cela vous permet de définir entièrement différents éléments de l’interface utilisateur et fichiers JavaScript pour l’expérience mobile.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p103">This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.</span></span>

<span data-ttu-id="dacd0-111">L’exemple suivant montre un bouton de volet Office unique dans `MobileFormFactor` un élément.</span><span class="sxs-lookup"><span data-stu-id="dacd0-111">The following example shows a single task pane button in a `MobileFormFactor` element.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

<span data-ttu-id="dacd0-112">Cet exemple est semblable aux éléments qui apparaissent dans un élément [DesktopFormFactor](../reference/manifest/desktopformfactor.md), avec toutefois quelques différences importantes.</span><span class="sxs-lookup"><span data-stu-id="dacd0-112">This is very similar to the elements that appear in a [DesktopFormFactor](../reference/manifest/desktopformfactor.md) element, with some notable differences.</span></span>

- <span data-ttu-id="dacd0-113">L’élément [OfficeTab](../reference/manifest/officetab.md) n’est pas utilisé.</span><span class="sxs-lookup"><span data-stu-id="dacd0-113">The [OfficeTab](../reference/manifest/officetab.md) element is not used.</span></span>
- <span data-ttu-id="dacd0-p104">L’élément [ExtensionPoint](../reference/manifest/extensionpoint.md) doit avoir un seul élément enfant. Si le complément ajoute uniquement un bouton, l’élément enfant doit être un élément [Control](../reference/manifest/control.md). Si le complément ajoute plusieurs boutons, l’élément enfant doit être un élément [Group](../reference/manifest/group.md) qui contient plusieurs éléments `Control`.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p104">The [ExtensionPoint](../reference/manifest/extensionpoint.md) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](../reference/manifest/control.md) element. If the add-in adds more than one button, the child element should be a [Group](../reference/manifest/group.md) element that contains multiple `Control` elements.</span></span>
- <span data-ttu-id="dacd0-117">Il n’existe aucun équivalent de type `Menu` pour l’élément `Control`.</span><span class="sxs-lookup"><span data-stu-id="dacd0-117">There is no `Menu` type equivalent for the `Control` element.</span></span>
- <span data-ttu-id="dacd0-118">L’élément [Supertip](../reference/manifest/supertip.md) n’est pas utilisé.</span><span class="sxs-lookup"><span data-stu-id="dacd0-118">The [Supertip](../reference/manifest/supertip.md) element is not used.</span></span>
- <span data-ttu-id="dacd0-p105">Les tailles d’icône requises sont différentes. Au minimum, les compléments mobiles doivent prendre en charge les icônes 25 x 25, 32 x 32 et 48 x 48 pixels.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p105">The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.</span></span>

## <a name="code-considerations"></a><span data-ttu-id="dacd0-121">Éléments à prendre en compte pour le code</span><span class="sxs-lookup"><span data-stu-id="dacd0-121">Code considerations</span></span>

<span data-ttu-id="dacd0-122">La conception d’un complément pour mobile implique certaines considérations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dacd0-122">Designing an add-in for mobile introduces some additional considerations.</span></span>

### <a name="use-rest-instead-of-exchange-web-services"></a><span data-ttu-id="dacd0-123">Utiliser REST plutôt que les services web Exchange</span><span class="sxs-lookup"><span data-stu-id="dacd0-123">Use REST instead of Exchange Web Services</span></span>

<span data-ttu-id="dacd0-p106">La méthode [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) n’est pas prise en charge dans Outlook Mobile. Les compléments doivent privilégier l’obtention d’informations auprès de l’API Office.js lorsque cela est possible. Si les compléments requièrent des informations non exposées par l’API Office.js, ils doivent utiliser les [API REST Outlook](/outlook/rest/) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p106">The [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.</span></span>

<span data-ttu-id="dacd0-127">L’ensemble de conditions requises pour les boîtes aux lettres 1,5 a introduit une nouvelle version d' [Office. Context. Mailbox. getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) qui peut demander un jeton d’accès compatible avec les API REST, ainsi qu’une nouvelle propriété [Office. Context. Mailbox. restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) qui peut être utilisée pour rechercher le point de terminaison de l’API REST pour l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dacd0-127">Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property that can be used to find the REST API endpoint for the user.</span></span>

### <a name="pinch-zoom"></a><span data-ttu-id="dacd0-128">Pincer pour zoomer</span><span class="sxs-lookup"><span data-stu-id="dacd0-128">Pinch zoom</span></span>

<span data-ttu-id="dacd0-p107">Par défaut les utilisateurs peuvent utiliser le mouvement pincer pour zoomer sur les volets Office. Si ce mouvement n’est pas pertinent pour votre scénario, veillez à désactiver la fonction « pincer pour zoomer » dans votre code HTML.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p107">By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.</span></span>

### <a name="close-task-panes"></a><span data-ttu-id="dacd0-131">Fermeture des volets Office</span><span class="sxs-lookup"><span data-stu-id="dacd0-131">Close task panes</span></span>

<span data-ttu-id="dacd0-p108">Dans Outlook Mobile, les volets Office occupent la totalité de l’écran et exigent par défaut que l’utilisateur les ferme pour revenir au message. Envisagez d’utiliser la méthode [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) pour fermer le volet Office lorsque votre scénario est terminé.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p108">In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) method to close the task pane when your scenario is complete.</span></span>

### <a name="compose-mode-and-appointments"></a><span data-ttu-id="dacd0-134">Mode composition et rendez-vous</span><span class="sxs-lookup"><span data-stu-id="dacd0-134">Compose mode and appointments</span></span>

<span data-ttu-id="dacd0-p109">Actuellement, les compléments dans Outlook Mobile ne prennent en charge l’activation que lors de la lecture des messages. Les compléments ne sont pas activés lors de la composition des messages, ou lors de l’affichage ou de la rédaction des rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dacd0-p109">Currently add-ins in Outlook Mobile only support activation when reading messages. Add-ins are not activated when composing messages or when viewing or composing appointments.</span></span>

### <a name="unsupported-apis"></a><span data-ttu-id="dacd0-137">API non prises en charge</span><span class="sxs-lookup"><span data-stu-id="dacd0-137">Unsupported APIs</span></span>

<span data-ttu-id="dacd0-138">Les API introduites dans l’ensemble de conditions requises 1,6 ou version ultérieure ne sont pas prises en charge par Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="dacd0-138">APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile.</span></span> <span data-ttu-id="dacd0-139">Les API suivantes des ensembles de conditions requises précédents ne sont pas non plus prises en charge.</span><span class="sxs-lookup"><span data-stu-id="dacd0-139">The following APIs from earlier requirement sets are also not supported.</span></span>

  - [<span data-ttu-id="dacd0-140">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="dacd0-140">Office.context.officeTheme</span></span>](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [<span data-ttu-id="dacd0-141">Office.context.mailbox.ewsUrl</span><span class="sxs-lookup"><span data-stu-id="dacd0-141">Office.context.mailbox.ewsUrl</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [<span data-ttu-id="dacd0-142">Office.context.mailbox.convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="dacd0-142">Office.context.mailbox.convertToEwsId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="dacd0-143">Office.context.mailbox.convertToRestId</span><span class="sxs-lookup"><span data-stu-id="dacd0-143">Office.context.mailbox.convertToRestId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="dacd0-144">Office.context.mailbox.displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dacd0-144">Office.context.mailbox.displayAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="dacd0-145">Office.context.mailbox.displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="dacd0-145">Office.context.mailbox.displayMessageForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="dacd0-146">Office.context.mailbox.displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dacd0-146">Office.context.mailbox.displayNewAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="dacd0-147">Office.context.mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="dacd0-147">Office.context.mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="dacd0-148">Office.context.mailbox.item.dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="dacd0-148">Office.context.mailbox.item.dateTimeModified</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [<span data-ttu-id="dacd0-149">Office.context.mailbox.item.displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="dacd0-149">Office.context.mailbox.item.displayReplyAllForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="dacd0-150">Office.context.mailbox.item.displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="dacd0-150">Office.context.mailbox.item.displayReplyForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="dacd0-151">Office.context.mailbox.item.getEntities</span><span class="sxs-lookup"><span data-stu-id="dacd0-151">Office.context.mailbox.item.getEntities</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="dacd0-152">Office.context.mailbox.item.getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="dacd0-152">Office.context.mailbox.item.getEntitiesByType</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="dacd0-153">Office.context.mailbox.item.getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="dacd0-153">Office.context.mailbox.item.getFilteredEntitiesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="dacd0-154">Office.context.mailbox.item.getRegexMatches</span><span class="sxs-lookup"><span data-stu-id="dacd0-154">Office.context.mailbox.item.getRegexMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="dacd0-155">Office.context.mailbox.item.getRegexMatchesByName</span><span class="sxs-lookup"><span data-stu-id="dacd0-155">Office.context.mailbox.item.getRegexMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a><span data-ttu-id="dacd0-156">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dacd0-156">See also</span></span>

[<span data-ttu-id="dacd0-157">Prise en charge de l’ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dacd0-157">Requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)