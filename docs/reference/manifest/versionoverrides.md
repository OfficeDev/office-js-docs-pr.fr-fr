---
title: Élémznr VersionOverrides dans le fichier manifest
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 0afa3183e34a736a878217c079b7b8d0259be5b1
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324798"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="2e7e9-102">Élément VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="2e7e9-102">VersionOverrides element</span></span>

<span data-ttu-id="2e7e9-p101">Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](./officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="2e7e9-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="2e7e9-106">Attributes</span></span>

|  <span data-ttu-id="2e7e9-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="2e7e9-107">Attribute</span></span>  |  <span data-ttu-id="2e7e9-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2e7e9-108">Required</span></span>  |  <span data-ttu-id="2e7e9-109">Description</span><span class="sxs-lookup"><span data-stu-id="2e7e9-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2e7e9-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="2e7e9-110">**xmlns**</span></span>       |  <span data-ttu-id="2e7e9-111">Oui</span><span class="sxs-lookup"><span data-stu-id="2e7e9-111">Yes</span></span>  |  <span data-ttu-id="2e7e9-112">Espace de noms du schéma VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-112">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="2e7e9-113">Les valeurs autorisées varient en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément et de la valeur **xsi : type** de `<OfficeApp>` l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-113">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="2e7e9-114">Voir les [valeurs d’espace de noms](#namespace-values) ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-114">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="2e7e9-115">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="2e7e9-115">**xsi:type**</span></span>  |  <span data-ttu-id="2e7e9-116">Oui</span><span class="sxs-lookup"><span data-stu-id="2e7e9-116">Yes</span></span>  | <span data-ttu-id="2e7e9-p103">Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="2e7e9-119">Valeurs d’espace de noms</span><span class="sxs-lookup"><span data-stu-id="2e7e9-119">Namespace values</span></span>

<span data-ttu-id="2e7e9-120">Le code suivant répertorie la valeur requise de la valeur **xmlns** en fonction de la valeur **xsi : type** de `<OfficeApp>` l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-120">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="2e7e9-121">**Taskpaneapp,** prend en charge uniquement la version 1,0 de VersionOverrides \*\*\*\* et le xmlns `http://schemas.microsoft.com/office/taskpaneappversionoverrides`doit être.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-121">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="2e7e9-122">**ContentApp** prend en charge uniquement la version 1,0 de VersionOverrides \*\*\*\* et le xmlns `http://schemas.microsoft.com/office/contentappversionoverrides`doit être.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-122">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="2e7e9-123">**MailApp** prend en charge les versions 1,0 et 1,1 de VersionOverrides, de \*\*\*\* sorte que la valeur de xmlns `<VersionOverrides>` varie en fonction de la valeur **xsi : type** de cet élément :</span><span class="sxs-lookup"><span data-stu-id="2e7e9-123">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="2e7e9-124">Lorsque **xsi : type** est `VersionOverridesV1_0`, **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides`.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-124">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="2e7e9-125">Lorsque **xsi : type** est `VersionOverridesV1_1`, **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-125">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="2e7e9-126">Actuellement, seul Outlook 2016 ou version ultérieure prend en charge le schéma VersionOverrides `VersionOverridesV1_1` v 1.1 et le type.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-126">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="2e7e9-127">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="2e7e9-127">Child elements</span></span>

|  <span data-ttu-id="2e7e9-128">Élément</span><span class="sxs-lookup"><span data-stu-id="2e7e9-128">Element</span></span> |  <span data-ttu-id="2e7e9-129">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2e7e9-129">Required</span></span>  |  <span data-ttu-id="2e7e9-130">Description</span><span class="sxs-lookup"><span data-stu-id="2e7e9-130">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2e7e9-131">**Description**</span><span class="sxs-lookup"><span data-stu-id="2e7e9-131">**Description**</span></span>    |  <span data-ttu-id="2e7e9-132">Non</span><span class="sxs-lookup"><span data-stu-id="2e7e9-132">No</span></span>   |  <span data-ttu-id="2e7e9-p104">Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](./resources.md). L’attribut `resid` de l’élément **Description** est défini sur la valeur de l’attribut `id` de l’élément `String` qui contient le texte.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="2e7e9-137">**Configuration requise**</span><span class="sxs-lookup"><span data-stu-id="2e7e9-137">**Requirements**</span></span>  |  <span data-ttu-id="2e7e9-138">Non</span><span class="sxs-lookup"><span data-stu-id="2e7e9-138">No</span></span>   |  <span data-ttu-id="2e7e9-p105">Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="2e7e9-141">Hôtes</span><span class="sxs-lookup"><span data-stu-id="2e7e9-141">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="2e7e9-142">Oui</span><span class="sxs-lookup"><span data-stu-id="2e7e9-142">Yes</span></span>  |  <span data-ttu-id="2e7e9-p106">Spécifie une collection d’hôtes d’Office. L’élément Hosts enfant remplace l’élément Hosts dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p106">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="2e7e9-145">Ressources</span><span class="sxs-lookup"><span data-stu-id="2e7e9-145">Resources</span></span>](./resources.md)    |  <span data-ttu-id="2e7e9-146">Oui</span><span class="sxs-lookup"><span data-stu-id="2e7e9-146">Yes</span></span>  | <span data-ttu-id="2e7e9-147">Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-147">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="2e7e9-148">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="2e7e9-148">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="2e7e9-149">Non</span><span class="sxs-lookup"><span data-stu-id="2e7e9-149">No</span></span>  | <span data-ttu-id="2e7e9-150">Spécifie les compléments natifs (COM/XLL) équivalents au complément Web.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-150">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="2e7e9-151">Le complément Web n’est pas activé si un complément natif équivalent est installé.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-151">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="2e7e9-152">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="2e7e9-152">**VersionOverrides**</span></span>    |  <span data-ttu-id="2e7e9-153">Non</span><span class="sxs-lookup"><span data-stu-id="2e7e9-153">No</span></span>  | <span data-ttu-id="2e7e9-p108">Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="2e7e9-156">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="2e7e9-156">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="2e7e9-157">Non</span><span class="sxs-lookup"><span data-stu-id="2e7e9-157">No</span></span>  | <span data-ttu-id="2e7e9-158">Fournit des détails sur l’inscription du complément avec des émetteurs de jetons sécurisés, tels qu’Azure Active Directory V 2.0.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-158">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="2e7e9-159">Exemple VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="2e7e9-159">VersionOverrides example</span></span>

<span data-ttu-id="2e7e9-160">Voici un exemple d’un élément typique `<VersionOverrides>` , y compris des éléments enfants qui ne sont pas obligatoires, mais qui sont généralement utilisés.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-160">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="2e7e9-161">Mise en œuvre de plusieurs versions</span><span class="sxs-lookup"><span data-stu-id="2e7e9-161">Implementing multiple versions</span></span>

<span data-ttu-id="2e7e9-p109">Un manifeste peut implémenter plusieurs versions de l’élément `VersionOverrides` qui prennent en charge différentes versions du schéma VersionOverrides. Cette opération permet éventuellement la prise en charge de nouvelles fonctionnalités dans un schéma plus récent tout en prenant en charge des clients plus anciens qui ne prennent pas en charge les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="2e7e9-164">Pour mettre en œuvre plusieurs versions, l’élément `VersionOverrides` de la nouvelle version doit être un enfant de l’élément `VersionOverrides` de l’ancienne version.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-164">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="2e7e9-165">L’élément enfant `VersionOverrides` n’hérite pas des valeurs du parent.</span><span class="sxs-lookup"><span data-stu-id="2e7e9-165">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="2e7e9-166">Pour mettre en œuvre à la fois les schémas VersionOverrides v1.0 et v1.1, le manifeste devrait ressembler à l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="2e7e9-166">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
