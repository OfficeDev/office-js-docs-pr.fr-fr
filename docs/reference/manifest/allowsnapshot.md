---
title: Élément AllowSnapshot dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f1aced0ce37b01c277ea5a8621f6c7764d2f761b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432346"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="2c7a7-102">AllowSnapshot, élément</span><span class="sxs-lookup"><span data-stu-id="2c7a7-102">AllowSnapshot element</span></span>

<span data-ttu-id="2c7a7-103">Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.</span><span class="sxs-lookup"><span data-stu-id="2c7a7-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="2c7a7-104">**Type de complément :** Contenu</span><span class="sxs-lookup"><span data-stu-id="2c7a7-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="2c7a7-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="2c7a7-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="2c7a7-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="2c7a7-106">Contained in</span></span>

[<span data-ttu-id="2c7a7-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2c7a7-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="2c7a7-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="2c7a7-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="2c7a7-109">**AllowSnapshot** est défini sur `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="2c7a7-109">Security Note:**AllowSnapshot** is true`true` by default.</span></span> <span data-ttu-id="2c7a7-110">Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="2c7a7-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="2c7a7-111">Toutefois, cela signifie également que les informations potentiellement sensibles affichées dans le complément sont accessibles directement à partir du document hébergeant le complément.</span><span class="sxs-lookup"><span data-stu-id="2c7a7-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

