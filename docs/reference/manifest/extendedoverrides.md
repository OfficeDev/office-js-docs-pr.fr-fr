---
title: Élément ExtendedOverrides dans le fichier manifeste
description: Spécifie les URL pour une extension au format JSON du manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996687"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="95dc1-103">Élément ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="95dc1-103">ExtendedOverrides element</span></span>

<span data-ttu-id="95dc1-104">Spécifie les URL complètes des fichiers au format JSON qui étendent le manifeste.</span><span class="sxs-lookup"><span data-stu-id="95dc1-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span>

<span data-ttu-id="95dc1-105">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="95dc1-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="95dc1-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="95dc1-106">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="95dc1-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="95dc1-107">Contained in</span></span>

[<span data-ttu-id="95dc1-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="95dc1-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="95dc1-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="95dc1-109">Can contain</span></span>

|<span data-ttu-id="95dc1-110">Élément</span><span class="sxs-lookup"><span data-stu-id="95dc1-110">Element</span></span>|<span data-ttu-id="95dc1-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="95dc1-111">Content</span></span>|<span data-ttu-id="95dc1-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="95dc1-112">Mail</span></span>|<span data-ttu-id="95dc1-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="95dc1-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="95dc1-114">Jetons</span><span class="sxs-lookup"><span data-stu-id="95dc1-114">Tokens</span></span>](tokens.md)|||<span data-ttu-id="95dc1-115">x</span><span class="sxs-lookup"><span data-stu-id="95dc1-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="95dc1-116">Attributs</span><span class="sxs-lookup"><span data-stu-id="95dc1-116">Attributes</span></span>

|<span data-ttu-id="95dc1-117">Attribut</span><span class="sxs-lookup"><span data-stu-id="95dc1-117">Attribute</span></span>|<span data-ttu-id="95dc1-118">Description</span><span class="sxs-lookup"><span data-stu-id="95dc1-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="95dc1-119">URL (obligatoire)</span><span class="sxs-lookup"><span data-stu-id="95dc1-119">Url (required)</span></span>| <span data-ttu-id="95dc1-120">URL complète du fichier JSON des substitutions étendues.</span><span class="sxs-lookup"><span data-stu-id="95dc1-120">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="95dc1-121">Il peut s’agir d’un modèle d’URL qui utilise des jetons définis par l’élément [tokens](tokens.md) .</span><span class="sxs-lookup"><span data-stu-id="95dc1-121">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|
|<span data-ttu-id="95dc1-122">ResourcesUrl (facultatif)</span><span class="sxs-lookup"><span data-stu-id="95dc1-122">ResourcesUrl (optional)</span></span> | <span data-ttu-id="95dc1-123">URL complète d’un fichier qui fournit des ressources supplémentaires, telles que des chaînes localisées, pour le fichier spécifié dans l' `Url` attribut.</span><span class="sxs-lookup"><span data-stu-id="95dc1-123">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="95dc1-124">Il peut s’agir d’un modèle d’URL qui utilise des jetons définis par l’élément [tokens](tokens.md) .</span><span class="sxs-lookup"><span data-stu-id="95dc1-124">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="example"></a><span data-ttu-id="95dc1-125">Exemple</span><span class="sxs-lookup"><span data-stu-id="95dc1-125">Example</span></span>

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
