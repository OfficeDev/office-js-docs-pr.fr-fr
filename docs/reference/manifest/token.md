---
title: Élément de jeton dans le fichier manifeste
description: Spécifie un jeton ou un caractère générique qui peut être utilisé avec les modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996685"
---
# <a name="token-element"></a><span data-ttu-id="ccd0d-103">Élément Token</span><span class="sxs-lookup"><span data-stu-id="ccd0d-103">Token element</span></span>

<span data-ttu-id="ccd0d-104">Définit un jeton d’URL individuel.</span><span class="sxs-lookup"><span data-stu-id="ccd0d-104">Defines an individual URL token.</span></span>

<span data-ttu-id="ccd0d-105">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="ccd0d-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="ccd0d-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ccd0d-106">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="ccd0d-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ccd0d-107">Contained in</span></span>

[<span data-ttu-id="ccd0d-108">Jetons</span><span class="sxs-lookup"><span data-stu-id="ccd0d-108">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="ccd0d-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="ccd0d-109">Can contain</span></span>

|<span data-ttu-id="ccd0d-110">Élément</span><span class="sxs-lookup"><span data-stu-id="ccd0d-110">Element</span></span>|<span data-ttu-id="ccd0d-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="ccd0d-111">Content</span></span>|<span data-ttu-id="ccd0d-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="ccd0d-112">Mail</span></span>|<span data-ttu-id="ccd0d-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="ccd0d-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="ccd0d-114">Override</span><span class="sxs-lookup"><span data-stu-id="ccd0d-114">Override</span></span>](override.md)|||<span data-ttu-id="ccd0d-115">x</span><span class="sxs-lookup"><span data-stu-id="ccd0d-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="ccd0d-116">Attributs</span><span class="sxs-lookup"><span data-stu-id="ccd0d-116">Attributes</span></span>

|<span data-ttu-id="ccd0d-117">Attribut</span><span class="sxs-lookup"><span data-stu-id="ccd0d-117">Attribute</span></span>|<span data-ttu-id="ccd0d-118">Description</span><span class="sxs-lookup"><span data-stu-id="ccd0d-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="ccd0d-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="ccd0d-119">DefaultValue</span></span>|<span data-ttu-id="ccd0d-120">Valeur par défaut de ce jeton si aucune condition n’est `<Override>` correspondante dans un élément enfant.</span><span class="sxs-lookup"><span data-stu-id="ccd0d-120">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="ccd0d-121">Nom</span><span class="sxs-lookup"><span data-stu-id="ccd0d-121">Name</span></span>|<span data-ttu-id="ccd0d-122">Nom du jeton.</span><span class="sxs-lookup"><span data-stu-id="ccd0d-122">Token name.</span></span> <span data-ttu-id="ccd0d-123">Ce nom est défini par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ccd0d-123">This name is user-defined.</span></span> <span data-ttu-id="ccd0d-124">Le type du jeton est déterminé par l’attribut type.</span><span class="sxs-lookup"><span data-stu-id="ccd0d-124">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="ccd0d-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ccd0d-125">xsi:type</span></span>|<span data-ttu-id="ccd0d-126">Définit le type de jeton.</span><span class="sxs-lookup"><span data-stu-id="ccd0d-126">Defines the kind of Token.</span></span> <span data-ttu-id="ccd0d-127">Cet attribut doit être défini sur l’un des éléments suivants :  `"RequirementsToken"` , ou  `"LocaleToken"` .</span><span class="sxs-lookup"><span data-stu-id="ccd0d-127">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="ccd0d-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="ccd0d-128">Example</span></span>

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