---
title: Élément Hosts dans le fichier manifeste
description: Spécifie l’application cliente Office dans laquelle le complément Office s’active.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611805"
---
# <a name="hosts-element"></a><span data-ttu-id="3d6a5-103">Hosts, élément</span><span class="sxs-lookup"><span data-stu-id="3d6a5-103">Hosts element</span></span>

<span data-ttu-id="3d6a5-p101">Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres.</span><span class="sxs-lookup"><span data-stu-id="3d6a5-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="3d6a5-106">Lorsqu’il est inclus dans le nœud [VersionOverrides](versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="3d6a5-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="3d6a5-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="3d6a5-107">Child elements</span></span>

|  <span data-ttu-id="3d6a5-108">Élément</span><span class="sxs-lookup"><span data-stu-id="3d6a5-108">Element</span></span> |  <span data-ttu-id="3d6a5-109">Requis</span><span class="sxs-lookup"><span data-stu-id="3d6a5-109">Required</span></span>  |  <span data-ttu-id="3d6a5-110">Description</span><span class="sxs-lookup"><span data-stu-id="3d6a5-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3d6a5-111">Host</span><span class="sxs-lookup"><span data-stu-id="3d6a5-111">Host</span></span>](host.md)    |  <span data-ttu-id="3d6a5-112">Oui</span><span class="sxs-lookup"><span data-stu-id="3d6a5-112">Yes</span></span>   |  <span data-ttu-id="3d6a5-113">Décrit un hôte et ses paramètres.</span><span class="sxs-lookup"><span data-stu-id="3d6a5-113">Describes a host and its settings.</span></span> |
