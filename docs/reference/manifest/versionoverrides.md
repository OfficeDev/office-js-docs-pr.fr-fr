---
title: Élémznr VersionOverrides dans le fichier manifest
description: Documentation de référence de l’élément VersionOverrides Office fichiers manifestes add-ins (XML).
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 0a70ded82b4603b1ac70698947a4710a4a44b5b6
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555149"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="e9d2a-103">Élément VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="e9d2a-103">VersionOverrides element</span></span>

<span data-ttu-id="e9d2a-p101">Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="e9d2a-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="e9d2a-107">Attributes</span></span>

|  <span data-ttu-id="e9d2a-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="e9d2a-108">Attribute</span></span>  |  <span data-ttu-id="e9d2a-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e9d2a-109">Required</span></span>  |  <span data-ttu-id="e9d2a-110">Description</span><span class="sxs-lookup"><span data-stu-id="e9d2a-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e9d2a-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="e9d2a-111">**xmlns**</span></span>       |  <span data-ttu-id="e9d2a-112">Oui</span><span class="sxs-lookup"><span data-stu-id="e9d2a-112">Yes</span></span>  |  <span data-ttu-id="e9d2a-113">L’espace nominatif versionOverrides schéma.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="e9d2a-114">Les valeurs autorisées varient en fonction de `<VersionOverrides>` la valeur **xsi:type de cet** élément et de la valeur **xsi:type** de l’élément `<OfficeApp>` parent.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="e9d2a-115">Voir les [valeurs namespace](#namespace-values) ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="e9d2a-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="e9d2a-116">**xsi:type**</span></span>  |  <span data-ttu-id="e9d2a-117">Oui</span><span class="sxs-lookup"><span data-stu-id="e9d2a-117">Yes</span></span>  | <span data-ttu-id="e9d2a-p103">Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="e9d2a-120">Valeurs de l’espace nominatif</span><span class="sxs-lookup"><span data-stu-id="e9d2a-120">Namespace values</span></span>

<span data-ttu-id="e9d2a-121">Ce qui suit énumère la valeur requise de la **valeur xmlns** en fonction de la **valeur xsi:type** de l’élément `<OfficeApp>` parent.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="e9d2a-122">**TaskPaneApp prend** en charge uniquement la version 1.0 de VersionOverrides, et **les xmlns** doivent être `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="e9d2a-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="e9d2a-123">**ContentApp** prend en charge uniquement la version 1.0 de VersionOverrides, et **les xmlns** doivent être `http://schemas.microsoft.com/office/contentappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="e9d2a-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="e9d2a-124">**MailApp** prend en charge les versions 1.0 et 1.1 de VersionOverrides, de sorte que la valeur **des xmlns** varie en fonction de la `<VersionOverrides>` valeur **xsi:type de cet** élément :</span><span class="sxs-lookup"><span data-stu-id="e9d2a-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="e9d2a-125">Lorsque **xsi:type** est `VersionOverridesV1_0` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="e9d2a-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="e9d2a-126">Lorsque **xsi:type** est `VersionOverridesV1_1` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .</span><span class="sxs-lookup"><span data-stu-id="e9d2a-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="e9d2a-127">Actuellement, Outlook 2016 ou plus tard prend en charge le schéma VersionOverrides v1.1 et le `VersionOverridesV1_1` type.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e9d2a-128">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="e9d2a-128">Child elements</span></span>

|  <span data-ttu-id="e9d2a-129">Élément</span><span class="sxs-lookup"><span data-stu-id="e9d2a-129">Element</span></span> |  <span data-ttu-id="e9d2a-130">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e9d2a-130">Required</span></span>  |  <span data-ttu-id="e9d2a-131">Description</span><span class="sxs-lookup"><span data-stu-id="e9d2a-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e9d2a-132">**Description**</span><span class="sxs-lookup"><span data-stu-id="e9d2a-132">**Description**</span></span>    |  <span data-ttu-id="e9d2a-133">Non</span><span class="sxs-lookup"><span data-stu-id="e9d2a-133">No</span></span>   |  <span data-ttu-id="e9d2a-134">Décrit le complément.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-134">Describes the add-in.</span></span> <span data-ttu-id="e9d2a-135">Cela remplace l’élément `Description` dans une partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-135">This overrides the `Description` element in any parent portion of the manifest.</span></span> <span data-ttu-id="e9d2a-136">Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="e9d2a-136">The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element.</span></span> <span data-ttu-id="e9d2a-137">`resid`L’attribut de **l’élément Description** ne peut pas être supérieur à 32 caractères et est défini sur la valeur de `id` `String` l’attribut de l’élément qui contient le texte.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-137">The `resid` attribute of the **Description** element can be no more than 32 characters and is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="e9d2a-138">**Configuration requise**</span><span class="sxs-lookup"><span data-stu-id="e9d2a-138">**Requirements**</span></span>  |  <span data-ttu-id="e9d2a-139">Non</span><span class="sxs-lookup"><span data-stu-id="e9d2a-139">No</span></span>   |  <span data-ttu-id="e9d2a-p105">Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="e9d2a-142">Hôtes</span><span class="sxs-lookup"><span data-stu-id="e9d2a-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="e9d2a-143">Oui</span><span class="sxs-lookup"><span data-stu-id="e9d2a-143">Yes</span></span>  |  <span data-ttu-id="e9d2a-144">Spécifie une collection d’applications Office’utilisation.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="e9d2a-145">L’élément hôtes de l’enfant l’emporte sur l’élément Hôtes dans la partie parente du manifeste.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="e9d2a-146">Ressources</span><span class="sxs-lookup"><span data-stu-id="e9d2a-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="e9d2a-147">Oui</span><span class="sxs-lookup"><span data-stu-id="e9d2a-147">Yes</span></span>  | <span data-ttu-id="e9d2a-148">Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="e9d2a-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="e9d2a-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="e9d2a-150">Non</span><span class="sxs-lookup"><span data-stu-id="e9d2a-150">No</span></span>  | <span data-ttu-id="e9d2a-151">Spécifie les modules d’ajout natifs (COM/XLL) équivalents à l’add-in web.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="e9d2a-152">L’add-in web n’est pas activé si un module d’ajout natif équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="e9d2a-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="e9d2a-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="e9d2a-154">Non</span><span class="sxs-lookup"><span data-stu-id="e9d2a-154">No</span></span>  | <span data-ttu-id="e9d2a-p108">Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="e9d2a-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="e9d2a-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="e9d2a-158">Non</span><span class="sxs-lookup"><span data-stu-id="e9d2a-158">No</span></span>  | <span data-ttu-id="e9d2a-159">Spécifie les détails de l’enregistrement de l’add-in auprès d’émetteurs de jetons sécurisés, tels que Azure Active Directory V2.0.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="e9d2a-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="e9d2a-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="e9d2a-161">Non</span><span class="sxs-lookup"><span data-stu-id="e9d2a-161">No</span></span>  |  <span data-ttu-id="e9d2a-162">Spécifie une collection d’autorisations étendues.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-162">Specifies a collection of extended permissions.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="e9d2a-163">Exemple VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="e9d2a-163">VersionOverrides example</span></span>

<span data-ttu-id="e9d2a-164">Ce qui suit est un exemple d’élément `<VersionOverrides>` typique, y compris certains éléments enfant qui ne sont pas nécessaires, mais sont généralement utilisés.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-164">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="e9d2a-165">Mise en œuvre de plusieurs versions</span><span class="sxs-lookup"><span data-stu-id="e9d2a-165">Implementing multiple versions</span></span>

<span data-ttu-id="e9d2a-p109">Un manifeste peut implémenter plusieurs versions de l’élément `VersionOverrides` qui prennent en charge différentes versions du schéma VersionOverrides. Cette opération permet éventuellement la prise en charge de nouvelles fonctionnalités dans un schéma plus récent tout en prenant en charge des clients plus anciens qui ne prennent pas en charge les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="e9d2a-168">Pour mettre en œuvre plusieurs versions, l’élément `VersionOverrides` de la nouvelle version doit être un enfant de l’élément `VersionOverrides` de l’ancienne version.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-168">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="e9d2a-169">L’élément enfant `VersionOverrides` n’hérite pas des valeurs du parent.</span><span class="sxs-lookup"><span data-stu-id="e9d2a-169">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="e9d2a-170">Pour mettre en œuvre à la fois les schémas VersionOverrides v1.0 et v1.1, le manifeste devrait ressembler à l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="e9d2a-170">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
