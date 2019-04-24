---
title: Élément Hosts dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 606073977366e37ecc4419f468f01bfb25647a7d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452024"
---
# <a name="hosts-element"></a>Hosts, élément

Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres. 

Lorsqu’il est inclus dans le nœud [VersionOverrides](versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste. 

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Oui   |  Décrit un hôte et ses paramètres. |
