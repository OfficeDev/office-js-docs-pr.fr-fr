---
title: Élément EquivalentAddin dans le fichier manifeste
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059922"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="ae237-102">Élément EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="ae237-102">EquivalentAddin element</span></span>

<span data-ttu-id="ae237-103">Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.</span><span class="sxs-lookup"><span data-stu-id="ae237-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="ae237-104">**Type de complément:** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="ae237-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="ae237-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ae237-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="ae237-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ae237-106">Contained in</span></span>

[<span data-ttu-id="ae237-107">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="ae237-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="ae237-108">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="ae237-108">Must contain</span></span>

[<span data-ttu-id="ae237-109">Type</span><span class="sxs-lookup"><span data-stu-id="ae237-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="ae237-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="ae237-110">Can contain</span></span>

<span data-ttu-id="ae237-111">[ProgID](progid.md)
[nom de fichier](filename.md)</span><span class="sxs-lookup"><span data-stu-id="ae237-111">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="ae237-112">Remarques</span><span class="sxs-lookup"><span data-stu-id="ae237-112">Remarks</span></span>

<span data-ttu-id="ae237-113">Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgId` éléments et. `Type`</span><span class="sxs-lookup"><span data-stu-id="ae237-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="ae237-114">Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .</span><span class="sxs-lookup"><span data-stu-id="ae237-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="ae237-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ae237-115">See also</span></span>

- [<span data-ttu-id="ae237-116">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="ae237-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="ae237-117">Faire en sorte que votre complément Excel soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="ae237-117">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)