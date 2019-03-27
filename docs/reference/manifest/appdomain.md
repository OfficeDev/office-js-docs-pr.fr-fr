---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870407"
---
# <a name="appdomain-element"></a>AppDomain, élément

Indique un domaine supplémentaire permettant de charger des pages dans la fenêtre du complément.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).
> 2. Ne placez *pas* de barre oblique («/») sur la valeur.

## <a name="contained-in"></a>Contenu dans

[AppDomains](appdomains.md)

## <a name="remarks"></a>Remarques

Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md). Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).
