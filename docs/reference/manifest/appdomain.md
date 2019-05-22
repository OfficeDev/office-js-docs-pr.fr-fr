---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337194"
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
