---
title: Inscrire un complément Office qui utilise l’authentification unique auprès du Plateforme d'identités Microsoft
description: Découvrez comment inscrire un complément Office auprès du Plateforme d'identités Microsoft pour utiliser l’authentification unique avec Word, Excel, PowerPoint et Outlook.
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0aab7d421ac57d1436d68c659f5d820717bcb846
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68842095"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>Inscrire un complément Office qui utilise l’authentification unique (SSO) auprès du Plateforme d'identités Microsoft

Cet article explique comment inscrire un complément Office auprès du Plateforme d'identités Microsoft afin de pouvoir utiliser l’authentification unique. Inscrivez le complément lorsque vous commencez à le développer afin que, lorsque vous passez au test ou à la production, vous puissiez modifier l’inscription existante ou créer des inscriptions distinctes pour les versions de développement, de test et de production du complément.

Le tableau suivant d?taille les informations n?cessaires pour effectuer cette proc?dure, ainsi que les espaces r?serv?s correspondant tels qu'ils apparaissent dans les instructions.

|Informations  |Exemples  |Espace réservé  |
|---------|---------|---------|
|Nom lisible par l’utilisateur pour le complément. (L’unicité est recommandée, mais pas obligatoire.)|`Contoso Marketing Excel Add-in (Prod)`|S/O|
|ID d’application qu’Azure génère pour vous dans le cadre du processus d’inscription.|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|Le nom de domaine complet du compl?ment (sauf pour le protocole). *Vous devez utiliser un domaine que vous poss?dez.* C'est pourquoi il n'est pas possible d'utiliser certains domaines connus comme `azurewebsites.net` ou `cloudapp.net`. Le domaine doit être identique, y compris tous les sous-domaines, comme il est utilisé dans les URL de la **\<Resources\>** section du manifeste du complément.|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|Les autorisations d’accès au Plateforme d'identités Microsoft et à Microsoft Graph dont votre complément a besoin. (`profile` est toujours requis.)|`profile`, `Files.Read.All`|S/O|

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]