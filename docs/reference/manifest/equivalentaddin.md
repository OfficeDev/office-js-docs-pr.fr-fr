---
title: Élément EquivalentAddin dans le fichier manifeste
description: Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836836"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="31b44-103">Élément EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="31b44-103">EquivalentAddin element</span></span>

<span data-ttu-id="31b44-104">Spécifie la compatibilité ascendante pour un add-in COM ou une XLL équivalent.</span><span class="sxs-lookup"><span data-stu-id="31b44-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="31b44-105">**Type de add-in :** Volet Des tâches, Fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="31b44-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="31b44-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="31b44-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="31b44-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="31b44-107">Contained in</span></span>

[<span data-ttu-id="31b44-108">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="31b44-108">EquivalentAddins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="31b44-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="31b44-109">Must contain</span></span>

[<span data-ttu-id="31b44-110">Type (Type)</span><span class="sxs-lookup"><span data-stu-id="31b44-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="31b44-111">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="31b44-111">Can contain</span></span>

<span data-ttu-id="31b44-112">[ProgId](progid.md) 
 [FileName](filename.md)</span><span class="sxs-lookup"><span data-stu-id="31b44-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="31b44-113">Remarques</span><span class="sxs-lookup"><span data-stu-id="31b44-113">Remarks</span></span>

<span data-ttu-id="31b44-114">Pour spécifier un compl?ment COM en tant que compl?ment équivalent, fournissez les deux `ProgId` `Type` éléments.</span><span class="sxs-lookup"><span data-stu-id="31b44-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="31b44-115">Pour spécifier un XLL en tant que module équivalent, fournissez à la fois les `FileName` éléments et les `Type` éléments.</span><span class="sxs-lookup"><span data-stu-id="31b44-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="31b44-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="31b44-116">See also</span></span>

- [<span data-ttu-id="31b44-117">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="31b44-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="31b44-118">Rendre votre complément Office compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="31b44-118">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)