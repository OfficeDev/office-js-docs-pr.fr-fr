---
title: Élément Hosts dans le fichier manifeste
description: Spécifie l’application cliente Office dans laquelle le complément Office s’active.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2684753fc32a295d7e177ef3bf668c194458128e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150459"
---
# <a name="hosts-element"></a>Hosts, élément

Spécifie l’application cliente Office dans laquelle le complément Office s’active. Contient une collection d’éléments **Host** et leurs paramètres. 

Lorsqu’il est inclus dans le nœud [VersionOverrides](versionoverrides.md), cet élément remplace l’élément **Hosts** dans la partie parent du manifeste. 

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Oui   |  Décrit un hôte et ses paramètres. |
