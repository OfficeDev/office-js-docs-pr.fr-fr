---
title: Élément AppDomain dans le fichier manifeste
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575638"
---
# <a name="appdomain-element"></a>AppDomain, élément

Spécifie les domaines supplémentaires qui chargent des pages dans la fenêtre du complément. Il répertorie également les domaines approuvés à partir desquels les appels de l’API Office. js peuvent être effectués depuis des IFrames au sein du complément.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain</AppDomain>`).
> 2. Ne placez *pas* de barre oblique (« / ») sur la valeur.

## <a name="contained-in"></a>Contenu dans

[AppDomains](appdomains.md)

## <a name="remarks"></a>Remarques

Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md). Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).
