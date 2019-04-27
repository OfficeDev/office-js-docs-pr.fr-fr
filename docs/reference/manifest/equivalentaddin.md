---
title: Élément EquivalentAddin dans le fichier manifeste
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356864"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="4c23d-102">Élément EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="4c23d-102">EquivalentAddin element</span></span>

<span data-ttu-id="4c23d-103">Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.</span><span class="sxs-lookup"><span data-stu-id="4c23d-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="4c23d-104">**Type de complément:** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="4c23d-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="4c23d-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="4c23d-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="4c23d-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="4c23d-106">Contained in</span></span>

[<span data-ttu-id="4c23d-107">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="4c23d-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="4c23d-108">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="4c23d-108">Must contain</span></span>

[<span data-ttu-id="4c23d-109">Type</span><span class="sxs-lookup"><span data-stu-id="4c23d-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="4c23d-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="4c23d-110">Can contain</span></span>

<span data-ttu-id="4c23d-111">[ProgID](progid.md)
[nom de fichier](filename.md)</span><span class="sxs-lookup"><span data-stu-id="4c23d-111">[ProgID](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="4c23d-112">Remarques</span><span class="sxs-lookup"><span data-stu-id="4c23d-112">Remarks</span></span>

<span data-ttu-id="4c23d-113">Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgID` éléments et. `Type`</span><span class="sxs-lookup"><span data-stu-id="4c23d-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgID` and `Type` elements.</span></span> <span data-ttu-id="4c23d-114">Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .</span><span class="sxs-lookup"><span data-stu-id="4c23d-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="4c23d-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4c23d-115">See also</span></span>

- [<span data-ttu-id="4c23d-116">Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur</span><span class="sxs-lookup"><span data-stu-id="4c23d-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="4c23d-117">Faire en sorte que votre complément Office soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="4c23d-117">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)