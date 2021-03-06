---
title: Élément ExtendedOverrides dans le fichier manifeste
description: Spécifie les URL d’une extension au format JSON du manifeste.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f433c9c5604f3fae35580ba20780ea6fe91401c7
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505471"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="7c767-103">Élément ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="7c767-103">ExtendedOverrides element</span></span>

<span data-ttu-id="7c767-104">Spécifie les URL complètes pour les fichiers au format JSON qui étendent le manifeste.</span><span class="sxs-lookup"><span data-stu-id="7c767-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span> <span data-ttu-id="7c767-105">Pour plus d’informations sur l’utilisation de cet élément et de ses éléments descendants, voir [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="7c767-105">For detailed information about the use of this element and its descendent elements, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="7c767-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="7c767-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="7c767-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7c767-107">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="7c767-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="7c767-108">Contained in</span></span>

[<span data-ttu-id="7c767-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7c767-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="7c767-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="7c767-110">Can contain</span></span>

|<span data-ttu-id="7c767-111">Élément</span><span class="sxs-lookup"><span data-stu-id="7c767-111">Element</span></span>|<span data-ttu-id="7c767-112">Contenu</span><span class="sxs-lookup"><span data-stu-id="7c767-112">Content</span></span>|<span data-ttu-id="7c767-113">Courrier</span><span class="sxs-lookup"><span data-stu-id="7c767-113">Mail</span></span>|<span data-ttu-id="7c767-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="7c767-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="7c767-115">Jetons</span><span class="sxs-lookup"><span data-stu-id="7c767-115">Tokens</span></span>](tokens.md)|||<span data-ttu-id="7c767-116">x</span><span class="sxs-lookup"><span data-stu-id="7c767-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="7c767-117">Attributs</span><span class="sxs-lookup"><span data-stu-id="7c767-117">Attributes</span></span>

|<span data-ttu-id="7c767-118">Attribut</span><span class="sxs-lookup"><span data-stu-id="7c767-118">Attribute</span></span>|<span data-ttu-id="7c767-119">Description</span><span class="sxs-lookup"><span data-stu-id="7c767-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="7c767-120">Url (obligatoire)</span><span class="sxs-lookup"><span data-stu-id="7c767-120">Url (required)</span></span>| <span data-ttu-id="7c767-121">URL complète du fichier JSON de remplacements étendu.</span><span class="sxs-lookup"><span data-stu-id="7c767-121">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="7c767-122">À l’avenir, cette valeur pourrait être un modèle d’URL qui utilise des jetons définis par [l’élément Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="7c767-122">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="7c767-123">Voir [exemples.](#examples)</span><span class="sxs-lookup"><span data-stu-id="7c767-123">See [Examples](#examples).</span></span>|
|<span data-ttu-id="7c767-124">ResourcesUrl (facultatif)</span><span class="sxs-lookup"><span data-stu-id="7c767-124">ResourcesUrl (optional)</span></span> | <span data-ttu-id="7c767-125">URL complète d’un fichier qui fournit des ressources supplémentaires, telles que des chaînes localisées, pour le fichier spécifié dans `Url` l’attribut.</span><span class="sxs-lookup"><span data-stu-id="7c767-125">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="7c767-126">Il peut s’agit d’un modèle d’URL qui utilise des jetons définis par [l’élément Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="7c767-126">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="examples"></a><span data-ttu-id="7c767-127">範例</span><span class="sxs-lookup"><span data-stu-id="7c767-127">Examples</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="7c767-128">À l’avenir, cette valeur pourrait être un modèle d’URL qui utilise des jetons définis par [l’élément Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="7c767-128">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="7c767-129">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="7c767-129">The following is an example.</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```
