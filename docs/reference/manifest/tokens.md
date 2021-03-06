---
title: Élément Tokens dans le fichier manifeste
description: Spécifie les jetons ou les caractères génériques qui peuvent être utilisés avec des modèles d’URL dans le manifeste.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505324"
---
# <a name="tokens-element"></a><span data-ttu-id="45323-103">Élément Tokens</span><span class="sxs-lookup"><span data-stu-id="45323-103">Tokens element</span></span>

<span data-ttu-id="45323-104">Définit les jetons qui peuvent être utilisés dans les URL de modèle.</span><span class="sxs-lookup"><span data-stu-id="45323-104">Defines tokens that could be used in template URLs.</span></span> <span data-ttu-id="45323-105">Pour plus d’informations sur l’utilisation de cet élément, voir [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="45323-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="45323-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="45323-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="45323-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="45323-107">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="45323-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="45323-108">Contained in</span></span>

[<span data-ttu-id="45323-109">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="45323-109">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="45323-110">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="45323-110">Must contain</span></span>

|<span data-ttu-id="45323-111">Élément</span><span class="sxs-lookup"><span data-stu-id="45323-111">Element</span></span>|<span data-ttu-id="45323-112">Contenu</span><span class="sxs-lookup"><span data-stu-id="45323-112">Content</span></span>|<span data-ttu-id="45323-113">Courrier</span><span class="sxs-lookup"><span data-stu-id="45323-113">Mail</span></span>|<span data-ttu-id="45323-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="45323-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="45323-115">Jeton</span><span class="sxs-lookup"><span data-stu-id="45323-115">Token</span></span>](token.md)|||<span data-ttu-id="45323-116">x</span><span class="sxs-lookup"><span data-stu-id="45323-116">x</span></span>|

## <a name="example"></a><span data-ttu-id="45323-117">Exemple</span><span class="sxs-lookup"><span data-stu-id="45323-117">Example</span></span>

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