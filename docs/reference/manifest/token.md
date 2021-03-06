---
title: Élément Token dans le fichier manifeste
description: Spécifie un jeton ou un caractère générique qui peut être utilisé avec des modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 48078f8211a8fd3f0e3f9d7c3f3aabd1d31b0a6d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505366"
---
# <a name="token-element"></a><span data-ttu-id="18e92-103">Élément Token</span><span class="sxs-lookup"><span data-stu-id="18e92-103">Token element</span></span>

<span data-ttu-id="18e92-104">Définit un jeton d’URL individuel.</span><span class="sxs-lookup"><span data-stu-id="18e92-104">Defines an individual URL token.</span></span> <span data-ttu-id="18e92-105">Pour plus d’informations sur l’utilisation de cet élément, voir [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="18e92-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="18e92-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="18e92-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="18e92-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="18e92-107">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="18e92-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="18e92-108">Contained in</span></span>

[<span data-ttu-id="18e92-109">Jetons</span><span class="sxs-lookup"><span data-stu-id="18e92-109">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="18e92-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="18e92-110">Can contain</span></span>

|<span data-ttu-id="18e92-111">Élément</span><span class="sxs-lookup"><span data-stu-id="18e92-111">Element</span></span>|<span data-ttu-id="18e92-112">Contenu</span><span class="sxs-lookup"><span data-stu-id="18e92-112">Content</span></span>|<span data-ttu-id="18e92-113">Courrier</span><span class="sxs-lookup"><span data-stu-id="18e92-113">Mail</span></span>|<span data-ttu-id="18e92-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="18e92-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="18e92-115">Override</span><span class="sxs-lookup"><span data-stu-id="18e92-115">Override</span></span>](override.md)|||<span data-ttu-id="18e92-116">x</span><span class="sxs-lookup"><span data-stu-id="18e92-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="18e92-117">Attributs</span><span class="sxs-lookup"><span data-stu-id="18e92-117">Attributes</span></span>

|<span data-ttu-id="18e92-118">Attribut</span><span class="sxs-lookup"><span data-stu-id="18e92-118">Attribute</span></span>|<span data-ttu-id="18e92-119">Description</span><span class="sxs-lookup"><span data-stu-id="18e92-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="18e92-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="18e92-120">DefaultValue</span></span>|<span data-ttu-id="18e92-121">Valeur par défaut de ce jeton si aucune condition dans un élément `<Override>` enfant ne correspond.</span><span class="sxs-lookup"><span data-stu-id="18e92-121">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="18e92-122">Nom</span><span class="sxs-lookup"><span data-stu-id="18e92-122">Name</span></span>|<span data-ttu-id="18e92-123">Nom du jeton.</span><span class="sxs-lookup"><span data-stu-id="18e92-123">Token name.</span></span> <span data-ttu-id="18e92-124">Ce nom est défini par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="18e92-124">This name is user-defined.</span></span> <span data-ttu-id="18e92-125">Le type du jeton est déterminé par l’attribut de type.</span><span class="sxs-lookup"><span data-stu-id="18e92-125">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="18e92-126">xsi:type</span><span class="sxs-lookup"><span data-stu-id="18e92-126">xsi:type</span></span>|<span data-ttu-id="18e92-127">Définit le type de jeton.</span><span class="sxs-lookup"><span data-stu-id="18e92-127">Defines the kind of Token.</span></span> <span data-ttu-id="18e92-128">Cet attribut doit être définie sur l’une des valeurs  `"RequirementsToken"` : ou  `"LocaleToken"` .</span><span class="sxs-lookup"><span data-stu-id="18e92-128">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="18e92-129">Exemple</span><span class="sxs-lookup"><span data-stu-id="18e92-129">Example</span></span>

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