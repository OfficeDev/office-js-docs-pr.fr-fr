---
title: Élément Tokens dans le fichier manifeste
description: Spécifie les jetons ou les caractères génériques qui peuvent être utilisés avec les modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996681"
---
# <a name="tokens-element"></a><span data-ttu-id="00de0-103">Élément Tokens</span><span class="sxs-lookup"><span data-stu-id="00de0-103">Tokens element</span></span>

<span data-ttu-id="00de0-104">Définit les jetons qui peuvent être utilisés dans les URL de modèles.</span><span class="sxs-lookup"><span data-stu-id="00de0-104">Defines tokens that could be used in template URLs.</span></span>

<span data-ttu-id="00de0-105">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="00de0-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="00de0-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="00de0-106">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="00de0-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="00de0-107">Contained in</span></span>

[<span data-ttu-id="00de0-108">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="00de0-108">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="00de0-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="00de0-109">Must contain</span></span>

|<span data-ttu-id="00de0-110">Élément</span><span class="sxs-lookup"><span data-stu-id="00de0-110">Element</span></span>|<span data-ttu-id="00de0-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="00de0-111">Content</span></span>|<span data-ttu-id="00de0-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="00de0-112">Mail</span></span>|<span data-ttu-id="00de0-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="00de0-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="00de0-114">Jeton</span><span class="sxs-lookup"><span data-stu-id="00de0-114">Token</span></span>](token.md)|||<span data-ttu-id="00de0-115">x</span><span class="sxs-lookup"><span data-stu-id="00de0-115">x</span></span>|

## <a name="example"></a><span data-ttu-id="00de0-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="00de0-116">Example</span></span>

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