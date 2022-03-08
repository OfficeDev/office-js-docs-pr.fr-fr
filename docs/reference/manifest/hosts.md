---
title: Élément Hosts dans le fichier manifeste
description: Spécifie l’Office applications clientes dans laquelle le Office’application sera activé.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9ea6cc9745f47b6e9b1c9bb0232b744304078053
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341071"
---
# <a name="hosts-element"></a>Hosts, élément

Spécifie l’Office applications clientes dans laquelle le Office’application sera activé. Contient une collection d’éléments **Host** et leurs paramètres. 

## <a name="as-child-of-versionoverrides-element"></a>Enfant de l’élément VersionOverrides

Les informations de cette section *s’appliquent uniquement* lorsque l’élément **Hosts** est un enfant [d’un Élément VersionOverrides](versionoverrides.md).

Cet élément remplace **l’élément Hosts** dans le manifeste de base.

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Oui   |  Décrit un hôte et ses paramètres. |
