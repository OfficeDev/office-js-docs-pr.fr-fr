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
# <a name="hosts-element"></a>Hosts, élément

Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres. 

Lorsqu’il est inclus dans le nœud [VersionOverrides](versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste. 

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Oui   |  Décrit un hôte et ses paramètres. |
