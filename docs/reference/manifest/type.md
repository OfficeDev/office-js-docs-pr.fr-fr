---
title: Élément type dans le fichier manifeste
description: L’élément type spécifie si le complément équivalent est un complément COM ou un XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604558"
---
# <a name="type-element"></a><span data-ttu-id="8b5a6-103">Élément Type</span><span class="sxs-lookup"><span data-stu-id="8b5a6-103">Type element</span></span>

<span data-ttu-id="8b5a6-104">Indique si le complément équivalent est un complément COM ou un XLL.</span><span class="sxs-lookup"><span data-stu-id="8b5a6-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="8b5a6-105">**Type de complément :** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="8b5a6-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="8b5a6-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8b5a6-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="8b5a6-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8b5a6-107">Contained in</span></span>

[<span data-ttu-id="8b5a6-108">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="8b5a6-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="8b5a6-109">Valeurs de type de complément</span><span class="sxs-lookup"><span data-stu-id="8b5a6-109">Add-in type values</span></span>

<span data-ttu-id="8b5a6-110">Vous devez spécifier l’une des valeurs suivantes pour l' `Type` élément.</span><span class="sxs-lookup"><span data-stu-id="8b5a6-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="8b5a6-111">COM : spécifie que le complément équivalent est un complément COM.</span><span class="sxs-lookup"><span data-stu-id="8b5a6-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="8b5a6-112">XLL : spécifie que le complément équivalent est une XLL Excel.</span><span class="sxs-lookup"><span data-stu-id="8b5a6-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="8b5a6-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8b5a6-113">See also</span></span>

- [<span data-ttu-id="8b5a6-114">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="8b5a6-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="8b5a6-115">Faire en sorte que votre complément Excel soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="8b5a6-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)