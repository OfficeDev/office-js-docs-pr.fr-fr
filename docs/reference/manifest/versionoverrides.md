---
title: Élémznr VersionOverrides dans le fichier manifest
description: Documentation de référence de l’élément VersionOverrides pour les fichiers manifeste des compléments Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 055a796d269ffb230293639a7f69b2dde4e36eee
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717851"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="7175a-103">Élément VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="7175a-103">VersionOverrides element</span></span>

<span data-ttu-id="7175a-p101">Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](./officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="7175a-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="7175a-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="7175a-107">Attributes</span></span>

|  <span data-ttu-id="7175a-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="7175a-108">Attribute</span></span>  |  <span data-ttu-id="7175a-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7175a-109">Required</span></span>  |  <span data-ttu-id="7175a-110">Description</span><span class="sxs-lookup"><span data-stu-id="7175a-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="7175a-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="7175a-111">**xmlns**</span></span>       |  <span data-ttu-id="7175a-112">Oui</span><span class="sxs-lookup"><span data-stu-id="7175a-112">Yes</span></span>  |  <span data-ttu-id="7175a-113">Espace de noms du schéma VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="7175a-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="7175a-114">Les valeurs autorisées varient en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément et de la valeur **xsi : type** de `<OfficeApp>` l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="7175a-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="7175a-115">Voir les [valeurs d’espace de noms](#namespace-values) ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="7175a-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="7175a-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="7175a-116">**xsi:type**</span></span>  |  <span data-ttu-id="7175a-117">Oui</span><span class="sxs-lookup"><span data-stu-id="7175a-117">Yes</span></span>  | <span data-ttu-id="7175a-p103">Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="7175a-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="7175a-120">Valeurs d’espace de noms</span><span class="sxs-lookup"><span data-stu-id="7175a-120">Namespace values</span></span>

<span data-ttu-id="7175a-121">Le code suivant répertorie la valeur requise de la valeur **xmlns** en fonction de la valeur **xsi : type** de `<OfficeApp>` l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="7175a-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="7175a-122">**Taskpaneapp,** prend en charge uniquement la version 1,0 de VersionOverrides **xmlns** et le xmlns `http://schemas.microsoft.com/office/taskpaneappversionoverrides`doit être.</span><span class="sxs-lookup"><span data-stu-id="7175a-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="7175a-123">**ContentApp** prend en charge uniquement la version 1,0 de VersionOverrides **xmlns** et le xmlns `http://schemas.microsoft.com/office/contentappversionoverrides`doit être.</span><span class="sxs-lookup"><span data-stu-id="7175a-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="7175a-124">**MailApp** prend en charge les versions 1,0 et 1,1 de VersionOverrides, de **xmlns** sorte que la valeur de xmlns `<VersionOverrides>` varie en fonction de la valeur **xsi : type** de cet élément :</span><span class="sxs-lookup"><span data-stu-id="7175a-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="7175a-125">Lorsque **xsi : type** est `VersionOverridesV1_0`, **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides`.</span><span class="sxs-lookup"><span data-stu-id="7175a-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="7175a-126">Lorsque **xsi : type** est `VersionOverridesV1_1`, **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span><span class="sxs-lookup"><span data-stu-id="7175a-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="7175a-127">Actuellement, seul Outlook 2016 ou version ultérieure prend en charge le schéma VersionOverrides `VersionOverridesV1_1` v 1.1 et le type.</span><span class="sxs-lookup"><span data-stu-id="7175a-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7175a-128">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="7175a-128">Child elements</span></span>

|  <span data-ttu-id="7175a-129">Élément</span><span class="sxs-lookup"><span data-stu-id="7175a-129">Element</span></span> |  <span data-ttu-id="7175a-130">Requis</span><span class="sxs-lookup"><span data-stu-id="7175a-130">Required</span></span>  |  <span data-ttu-id="7175a-131">Description</span><span class="sxs-lookup"><span data-stu-id="7175a-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="7175a-132">**Description**</span><span class="sxs-lookup"><span data-stu-id="7175a-132">**Description**</span></span>    |  <span data-ttu-id="7175a-133">Non</span><span class="sxs-lookup"><span data-stu-id="7175a-133">No</span></span>   |  <span data-ttu-id="7175a-p104">Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md). L’attribut `resid` de l’élément **Description** est défini sur la valeur de l’attribut `id` de l’élément `String` qui contient le texte.</span><span class="sxs-lookup"><span data-stu-id="7175a-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="7175a-138">**Configuration requise**</span><span class="sxs-lookup"><span data-stu-id="7175a-138">**Requirements**</span></span>  |  <span data-ttu-id="7175a-139">Non</span><span class="sxs-lookup"><span data-stu-id="7175a-139">No</span></span>   |  <span data-ttu-id="7175a-p105">Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="7175a-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="7175a-142">Hôtes</span><span class="sxs-lookup"><span data-stu-id="7175a-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="7175a-143">Oui</span><span class="sxs-lookup"><span data-stu-id="7175a-143">Yes</span></span>  |  <span data-ttu-id="7175a-p106">Spécifie une collection d’hôtes d’Office. L’élément Hosts enfant remplace l’élément Hosts dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="7175a-p106">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="7175a-146">Ressources</span><span class="sxs-lookup"><span data-stu-id="7175a-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="7175a-147">Oui</span><span class="sxs-lookup"><span data-stu-id="7175a-147">Yes</span></span>  | <span data-ttu-id="7175a-148">Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.</span><span class="sxs-lookup"><span data-stu-id="7175a-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="7175a-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="7175a-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="7175a-150">Non</span><span class="sxs-lookup"><span data-stu-id="7175a-150">No</span></span>  | <span data-ttu-id="7175a-151">Spécifie les compléments natifs (COM/XLL) équivalents au complément Web.</span><span class="sxs-lookup"><span data-stu-id="7175a-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="7175a-152">Le complément Web n’est pas activé si un complément natif équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="7175a-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="7175a-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="7175a-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="7175a-154">Non</span><span class="sxs-lookup"><span data-stu-id="7175a-154">No</span></span>  | <span data-ttu-id="7175a-p108">Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="7175a-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="7175a-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="7175a-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="7175a-158">Non</span><span class="sxs-lookup"><span data-stu-id="7175a-158">No</span></span>  | <span data-ttu-id="7175a-159">Fournit des détails sur l’inscription du complément avec des émetteurs de jetons sécurisés, tels qu’Azure Active Directory V 2.0.</span><span class="sxs-lookup"><span data-stu-id="7175a-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="7175a-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="7175a-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="7175a-161">Non</span><span class="sxs-lookup"><span data-stu-id="7175a-161">No</span></span>  |  <span data-ttu-id="7175a-162">Spécifie une collection d’autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="7175a-162">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="7175a-163">**Important**: étant donné que l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) est actuellement en préversion, les compléments qui utilisent l' `ExtendedPermissions` élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé.</span><span class="sxs-lookup"><span data-stu-id="7175a-163">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="7175a-164">Exemple VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="7175a-164">VersionOverrides example</span></span>

<span data-ttu-id="7175a-165">Voici un exemple d’un élément typique `<VersionOverrides>` , y compris des éléments enfants qui ne sont pas obligatoires, mais qui sont généralement utilisés.</span><span class="sxs-lookup"><span data-stu-id="7175a-165">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="7175a-166">Mise en œuvre de plusieurs versions</span><span class="sxs-lookup"><span data-stu-id="7175a-166">Implementing multiple versions</span></span>

<span data-ttu-id="7175a-p109">Un manifeste peut implémenter plusieurs versions de l’élément `VersionOverrides` qui prennent en charge différentes versions du schéma VersionOverrides. Cette opération permet éventuellement la prise en charge de nouvelles fonctionnalités dans un schéma plus récent tout en prenant en charge des clients plus anciens qui ne prennent pas en charge les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="7175a-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="7175a-169">Pour mettre en œuvre plusieurs versions, l’élément `VersionOverrides` de la nouvelle version doit être un enfant de l’élément `VersionOverrides` de l’ancienne version.</span><span class="sxs-lookup"><span data-stu-id="7175a-169">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="7175a-170">L’élément enfant `VersionOverrides` n’hérite pas des valeurs du parent.</span><span class="sxs-lookup"><span data-stu-id="7175a-170">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="7175a-171">Pour mettre en œuvre à la fois les schémas VersionOverrides v1.0 et v1.1, le manifeste devrait ressembler à l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="7175a-171">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
