---
title: Inscrire un Office qui utilise l’sso avec le Plateforme d'identités Microsoft
description: Découvrez comment inscrire un Office avec le Plateforme d'identités Microsoft pour utiliser l’sso avec Word, Excel, PowerPoint et Outlook.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: b11ce5130e020b049038631b9ae1c0e62fdadeab
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320244"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>Inscrire un Office qui utilise l’sign-on unique (SSO) avec le Plateforme d'identités Microsoft

Cet article explique comment inscrire un Office avec le Plateforme d'identités Microsoft afin que vous pouvez utiliser l’utilisateur sso. Inscrivez le add-in lorsque vous commencez à le développer afin que lorsque vous progressez vers le test ou la production, vous pouvez modifier l’inscription existante ou créer des inscriptions distinctes pour les versions de développement, de test et de production du module.

Le tableau suivant d?taille les informations n?cessaires pour effectuer cette proc?dure, ainsi que les espaces r?serv?s correspondant tels qu'ils apparaissent dans les instructions.

|Informations  |Exemples  |Espace réservé  |
|---------|---------|---------|
|A human readable name for the add-in. (Uniqueness recommended, but not required.)|`Contoso Marketing Excel Add-in (Prod)`|S/O|
|ID d’application qu’Azure génère pour vous dans le cadre du processus d’inscription.|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|Le nom de domaine complet du compl?ment (sauf pour le protocole). *Vous devez utiliser un domaine que vous poss?dez.* C'est pourquoi il n'est pas possible d'utiliser certains domaines connus comme `azurewebsites.net` ou `cloudapp.net`. Le domaine et les sous-domaines doivent être les mêmes que ceux utilisés dans les URL dans la section `<Resources>` du manifeste du complément.|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|Les autorisations sur les Plateforme d'identités Microsoft et Microsoft Graph dont votre add-in a besoin. (`profile` est toujours requis.)|`profile`, `Files.Read.All`|S/O|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
