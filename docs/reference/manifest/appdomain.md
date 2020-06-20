---
title: Élément AppDomain dans le fichier manifeste
description: Spécifie les domaines supplémentaires utilisés par votre complément et doit être approuvé par Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778647"
---
# <a name="appdomain-element"></a>AppDomain, élément

Spécifie un domaine supplémentaire qu’Office doit approuver, en plus de celui spécifié dans l' [élément SourceLocation](sourcelocation.md). La spécification d’un domaine a les effets suivants :

- Elle permet l’ouverture directe des pages, des itinéraires ou d’autres ressources dans le domaine dans le volet Office racine du complément sur les plateformes de bureau. (Il n’est pas nécessaire de spécifier un domaine dans un **AppDomain** pour Office sur le Web ou d’ouvrir une ressource dans un IFRAME, et il n’est pas nécessaire d’ouvrir une ressource dans une boîte de dialogue ouverte avec l' [API Dialog](../../develop/dialog-api-in-office-add-ins.md).)
- Elle permet aux pages du domaine d’effectuer des appels d’API Office.js à partir d’IFrames dans le complément.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. S’il existe un port explicite pour le domaine, incluez-le (par exemple, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).
> 3. Si un sous-domaine doit être approuvé, incluez-le (par exemple, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ). Le sous-domaine `mysubdomain.mydomain.com` et `mydomain.com` sont des domaines différents. Si les deux doivent être approuvés, les deux doivent se trouver dans des éléments **AppDomain** distincts.
> 4. Le fait de répertorier le même domaine que celui spécifié dans l' [élément SourceLocation](sourcelocation.md) n’a aucun effet et peut être trompeur. En particulier, lorsque vous développez sur `localhost` , vous n’avez pas besoin de créer un élément **AppDomain** pour `localhost` .
> 5. N’incluez pas de segments d’URL au-delà du domaine. Par exemple, n’incluez pas l’URL complète d’une page.
> 6. Ne placez *pas* de barre oblique (« / ») sur la valeur.

## <a name="contained-in"></a>Contenu dans

[AppDomains](appdomains.md)

## <a name="remarks"></a>Remarques

Pour plus d’informations, voir le [manifeste XML de compléments Office](../../develop/add-in-manifests.md).
