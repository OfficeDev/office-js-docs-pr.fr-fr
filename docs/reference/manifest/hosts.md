---
title: Élément Hosts dans le fichier manifeste
description: Spécifie l’application cliente Office dans laquelle le complément Office s’active.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718103"
---
# <a name="hosts-element"></a>Hosts, élément

Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres. 

Lorsqu’il est inclus dans le nœud [VersionOverrides](versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste. 

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Oui   |  Décrit un hôte et ses paramètres. |
