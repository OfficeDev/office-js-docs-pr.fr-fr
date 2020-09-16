---
title: Élémznr VersionOverrides dans le fichier manifest
description: Documentation de référence de l’élément VersionOverrides pour les fichiers manifeste des compléments Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: a744772c01c57c41a9dc20ee0accea5f070c3ff3
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819825"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="89042-103">Élément VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="89042-103">VersionOverrides element</span></span>

<span data-ttu-id="89042-p101">Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="89042-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="89042-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="89042-107">Attributes</span></span>

|  <span data-ttu-id="89042-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="89042-108">Attribute</span></span>  |  <span data-ttu-id="89042-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="89042-109">Required</span></span>  |  <span data-ttu-id="89042-110">Description</span><span class="sxs-lookup"><span data-stu-id="89042-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="89042-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="89042-111">**xmlns**</span></span>       |  <span data-ttu-id="89042-112">Oui</span><span class="sxs-lookup"><span data-stu-id="89042-112">Yes</span></span>  |  <span data-ttu-id="89042-113">Espace de noms du schéma VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="89042-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="89042-114">Les valeurs autorisées varient en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément et de la valeur **xsi : type** de l' `<OfficeApp>` élément parent.</span><span class="sxs-lookup"><span data-stu-id="89042-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="89042-115">Voir les [valeurs d’espace de noms](#namespace-values) ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="89042-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="89042-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="89042-116">**xsi:type**</span></span>  |  <span data-ttu-id="89042-117">Oui</span><span class="sxs-lookup"><span data-stu-id="89042-117">Yes</span></span>  | <span data-ttu-id="89042-p103">Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="89042-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="89042-120">Valeurs d’espace de noms</span><span class="sxs-lookup"><span data-stu-id="89042-120">Namespace values</span></span>

<span data-ttu-id="89042-121">Le code suivant répertorie la valeur requise de la valeur **xmlns** en fonction de la valeur **xsi : type** de l' `<OfficeApp>` élément parent.</span><span class="sxs-lookup"><span data-stu-id="89042-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="89042-122">**Taskpaneapp,** prend en charge uniquement la version 1,0 de VersionOverrides et le **xmlns** doit être `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="89042-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="89042-123">**ContentApp** prend en charge uniquement la version 1,0 de VersionOverrides et le **xmlns** doit être `http://schemas.microsoft.com/office/contentappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="89042-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="89042-124">**MailApp** prend en charge les versions 1,0 et 1,1 de VersionOverrides, de sorte que la valeur de **xmlns** varie en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément :</span><span class="sxs-lookup"><span data-stu-id="89042-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="89042-125">Lorsque **xsi : type** est `VersionOverridesV1_0` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="89042-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="89042-126">Lorsque **xsi : type** est `VersionOverridesV1_1` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .</span><span class="sxs-lookup"><span data-stu-id="89042-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="89042-127">Actuellement, seul Outlook 2016 ou version ultérieure prend en charge le schéma VersionOverrides v 1.1 et le `VersionOverridesV1_1` type.</span><span class="sxs-lookup"><span data-stu-id="89042-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="89042-128">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="89042-128">Child elements</span></span>

|  <span data-ttu-id="89042-129">Élément</span><span class="sxs-lookup"><span data-stu-id="89042-129">Element</span></span> |  <span data-ttu-id="89042-130">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="89042-130">Required</span></span>  |  <span data-ttu-id="89042-131">Description</span><span class="sxs-lookup"><span data-stu-id="89042-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="89042-132">**Description**</span><span class="sxs-lookup"><span data-stu-id="89042-132">**Description**</span></span>    |  <span data-ttu-id="89042-133">Non</span><span class="sxs-lookup"><span data-stu-id="89042-133">No</span></span>   |  <span data-ttu-id="89042-p104">Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md). L’attribut `resid` de l’élément **Description** est défini sur la valeur de l’attribut `id` de l’élément `String` qui contient le texte.</span><span class="sxs-lookup"><span data-stu-id="89042-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="89042-138">**Configuration requise**</span><span class="sxs-lookup"><span data-stu-id="89042-138">**Requirements**</span></span>  |  <span data-ttu-id="89042-139">Non</span><span class="sxs-lookup"><span data-stu-id="89042-139">No</span></span>   |  <span data-ttu-id="89042-p105">Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="89042-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="89042-142">Hôtes</span><span class="sxs-lookup"><span data-stu-id="89042-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="89042-143">Oui</span><span class="sxs-lookup"><span data-stu-id="89042-143">Yes</span></span>  |  <span data-ttu-id="89042-144">Spécifie une collection d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="89042-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="89042-145">L’élément hosts enfant remplace l’élément hosts dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="89042-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="89042-146">Ressources</span><span class="sxs-lookup"><span data-stu-id="89042-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="89042-147">Oui</span><span class="sxs-lookup"><span data-stu-id="89042-147">Yes</span></span>  | <span data-ttu-id="89042-148">Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.</span><span class="sxs-lookup"><span data-stu-id="89042-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="89042-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="89042-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="89042-150">Non</span><span class="sxs-lookup"><span data-stu-id="89042-150">No</span></span>  | <span data-ttu-id="89042-151">Spécifie les compléments natifs (COM/XLL) équivalents au complément Web.</span><span class="sxs-lookup"><span data-stu-id="89042-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="89042-152">Le complément Web n’est pas activé si un complément natif équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="89042-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="89042-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="89042-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="89042-154">Non</span><span class="sxs-lookup"><span data-stu-id="89042-154">No</span></span>  | <span data-ttu-id="89042-p108">Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="89042-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="89042-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="89042-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="89042-158">Non</span><span class="sxs-lookup"><span data-stu-id="89042-158">No</span></span>  | <span data-ttu-id="89042-159">Fournit des détails sur l’inscription du complément avec des émetteurs de jetons sécurisés, tels qu’Azure Active Directory V 2.0.</span><span class="sxs-lookup"><span data-stu-id="89042-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="89042-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="89042-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="89042-161">Non</span><span class="sxs-lookup"><span data-stu-id="89042-161">No</span></span>  |  <span data-ttu-id="89042-162">Spécifie une collection d’autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="89042-162">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="89042-163">**Important**: étant donné que l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) est actuellement en préversion, les compléments qui utilisent l' `ExtendedPermissions` élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé.</span><span class="sxs-lookup"><span data-stu-id="89042-163">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="89042-164">Exemple VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="89042-164">VersionOverrides example</span></span>

<span data-ttu-id="89042-165">Voici un exemple d’un `<VersionOverrides>` élément typique, y compris des éléments enfants qui ne sont pas obligatoires, mais qui sont généralement utilisés.</span><span class="sxs-lookup"><span data-stu-id="89042-165">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="89042-166">Mise en œuvre de plusieurs versions</span><span class="sxs-lookup"><span data-stu-id="89042-166">Implementing multiple versions</span></span>

<span data-ttu-id="89042-p109">Un manifeste peut implémenter plusieurs versions de l’élément `VersionOverrides` qui prennent en charge différentes versions du schéma VersionOverrides. Cette opération permet éventuellement la prise en charge de nouvelles fonctionnalités dans un schéma plus récent tout en prenant en charge des clients plus anciens qui ne prennent pas en charge les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="89042-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="89042-169">Pour mettre en œuvre plusieurs versions, l’élément `VersionOverrides` de la nouvelle version doit être un enfant de l’élément `VersionOverrides` de l’ancienne version.</span><span class="sxs-lookup"><span data-stu-id="89042-169">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="89042-170">L’élément enfant `VersionOverrides` n’hérite pas des valeurs du parent.</span><span class="sxs-lookup"><span data-stu-id="89042-170">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="89042-171">Pour mettre en œuvre à la fois les schémas VersionOverrides v1.0 et v1.1, le manifeste devrait ressembler à l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="89042-171">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
