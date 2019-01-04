---
title: Élément Hosts dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 59010c0f6c0d14d8721856f81def11540db28704
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433410"
---
# <a name="hosts-element"></a><span data-ttu-id="2d0ab-102">Hosts, élément</span><span class="sxs-lookup"><span data-stu-id="2d0ab-102">Hosts element</span></span>

<span data-ttu-id="2d0ab-p101">Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres.</span><span class="sxs-lookup"><span data-stu-id="2d0ab-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="2d0ab-105">Lorsqu’il est inclus dans le nœud [VersionOverrides](versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="2d0ab-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="2d0ab-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="2d0ab-106">Child elements</span></span>

|  <span data-ttu-id="2d0ab-107">Élément</span><span class="sxs-lookup"><span data-stu-id="2d0ab-107">Element</span></span> |  <span data-ttu-id="2d0ab-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2d0ab-108">Required</span></span>  |  <span data-ttu-id="2d0ab-109">Description</span><span class="sxs-lookup"><span data-stu-id="2d0ab-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="2d0ab-110">Host</span><span class="sxs-lookup"><span data-stu-id="2d0ab-110">Host</span></span>](host.md)    |  <span data-ttu-id="2d0ab-111">Oui</span><span class="sxs-lookup"><span data-stu-id="2d0ab-111">Yes</span></span>   |  <span data-ttu-id="2d0ab-112">Décrit un hôte et ses paramètres.</span><span class="sxs-lookup"><span data-stu-id="2d0ab-112">Describes a host and its settings.</span></span> |
