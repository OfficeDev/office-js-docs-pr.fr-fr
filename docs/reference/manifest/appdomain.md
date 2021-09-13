---
title: Élément AppDomain dans le fichier manifeste
description: Spécifie les domaines supplémentaires qui sont utilisés par votre complément et qui doivent être Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: c17195e6d9d3f4f22465c8aa1fc626afd3eb06c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152187"
---
# <a name="appdomain-element"></a>AppDomain, élément

Spécifie un domaine supplémentaire qui doit Office, en plus de celui spécifié dans [l’élément SourceLocation](sourcelocation.md). La spécification d’un domaine a les effets suivants :

- Il permet d’ouvrir des pages, des itinéraires ou d’autres ressources dans le domaine directement dans le volet Des tâches racine du module de Office de bureau. (La spécification d’un domaine dans un **AppDomain** n’est pas nécessaire pour Office sur le Web ou pour ouvrir une ressource dans un IFrame, ni pour ouvrir une ressource dans une boîte de dialogue ouverte avec [l’API](../../develop/dialog-api-in-office-add-ins.md)de dialogue.)
- Il permet aux pages du domaine d’effectuer des Office.js API à partir d’IFrames au sein du module.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. La valeur de l’élément **AppDomain** doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. S’il existe un port explicite pour le domaine, incluez-le (par exemple, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).
> 3. Si un sous-domaine doit être approuvé, incluez-le (par exemple, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ). Le sous-domaine et `mysubdomain.mydomain.com` `mydomain.com` sont des domaines différents. Si les deux doivent être fiables, les deux doivent se trouver dans des éléments **AppDomain** distincts.
> 4. Le fait de répertorier le même domaine que celui spécifié dans l’élément [SourceLocation](sourcelocation.md) n’a aucun effet et peut être erroné. En particulier, lorsque vous développez sur , vous n’avez pas besoin de créer un élément `localhost` **AppDomain** pour `localhost` .
> 5. N’incluez aucun segment d’une URL au-delà du domaine. Par exemple, n’incluez pas l’URL complète d’une page.
> 6. Ne *placez* pas de barre oblique fermante « / » sur la valeur.

## <a name="contained-in"></a>Contenu dans

[AppDomains](appdomains.md)

## <a name="remarks"></a>Remarques

Pour plus d’informations, voir le [manifeste XML de compléments Office](../../develop/add-in-manifests.md).
