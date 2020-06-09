---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611560"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="1ad6a-103">Élément EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="1ad6a-103">EquivalentAddin element</span></span>

<span data-ttu-id="1ad6a-104">Spécifie la compatibilité descendante pour un complément COM équivalent ou une XLL.</span><span class="sxs-lookup"><span data-stu-id="1ad6a-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="1ad6a-105">**Type de complément :** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="1ad6a-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="1ad6a-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1ad6a-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="1ad6a-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="1ad6a-107">Contained in</span></span>

[<span data-ttu-id="1ad6a-108">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="1ad6a-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="1ad6a-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="1ad6a-109">Must contain</span></span>

[<span data-ttu-id="1ad6a-110">Type</span><span class="sxs-lookup"><span data-stu-id="1ad6a-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="1ad6a-111">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="1ad6a-111">Can contain</span></span>

<span data-ttu-id="1ad6a-112">[ProgID](progid.md) 
 [Nom de fichier](filename.md)</span><span class="sxs-lookup"><span data-stu-id="1ad6a-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="1ad6a-113">Remarques</span><span class="sxs-lookup"><span data-stu-id="1ad6a-113">Remarks</span></span>

<span data-ttu-id="1ad6a-114">Pour spécifier un complément COM en tant que complément équivalent, fournissez les `ProgId` `Type` éléments et.</span><span class="sxs-lookup"><span data-stu-id="1ad6a-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="1ad6a-115">Pour spécifier un XLL en tant que complément équivalent, fournissez les `FileName` éléments et `Type` .</span><span class="sxs-lookup"><span data-stu-id="1ad6a-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="1ad6a-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1ad6a-116">See also</span></span>

- [<span data-ttu-id="1ad6a-117">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="1ad6a-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="1ad6a-118">Faire en sorte que votre complément Excel soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="1ad6a-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)