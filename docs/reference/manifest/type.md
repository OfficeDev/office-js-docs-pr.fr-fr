---
title: Élément type dans le fichier manifeste
description: L’élément type spécifie si le complément équivalent est un complément COM ou un XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720315"
---
# <a name="type-element"></a><span data-ttu-id="e7278-103">Élément Type</span><span class="sxs-lookup"><span data-stu-id="e7278-103">Type element</span></span>

<span data-ttu-id="e7278-104">Indique si le complément équivalent est un complément COM ou un XLL.</span><span class="sxs-lookup"><span data-stu-id="e7278-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="e7278-105">**Type de complément :** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="e7278-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="e7278-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="e7278-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="e7278-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="e7278-107">Contained in</span></span>

[<span data-ttu-id="e7278-108">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="e7278-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="e7278-109">Valeurs de type de complément</span><span class="sxs-lookup"><span data-stu-id="e7278-109">Add-in type values</span></span>

<span data-ttu-id="e7278-110">Vous devez spécifier l’une des valeurs suivantes pour l' `Type` élément.</span><span class="sxs-lookup"><span data-stu-id="e7278-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="e7278-111">COM : spécifie que le complément équivalent est un complément COM.</span><span class="sxs-lookup"><span data-stu-id="e7278-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="e7278-112">XLL : spécifie que le complément équivalent est une XLL Excel.</span><span class="sxs-lookup"><span data-stu-id="e7278-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="e7278-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e7278-113">See also</span></span>

- [<span data-ttu-id="e7278-114">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="e7278-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="e7278-115">Faire en sorte que votre complément Excel soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="e7278-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)