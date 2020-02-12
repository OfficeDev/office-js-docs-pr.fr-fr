---
title: Enregistrer un complément Office utilisant une SSO (authentification unique) au point de terminaison Azure AD v2.0
description: ''
ms.date: 04/10/2019
localization_priority: Normal
ms.openlocfilehash: 3594b1e1b22f7a4341b5fd9a5b6774f3d21d8c26
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949670"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Enregistrer un complément Office utilisant une SSO (authentification unique) au point de terminaison Azure AD v2.0

This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint. You need to register the add-in when you begin developing it. When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.

Le tableau suivant d?taille les informations n?cessaires pour effectuer cette proc?dure, ainsi que les espaces r?serv?s correspondant tels qu'ils apparaissent dans les instructions.

|Informations  |Exemples  |Espace réservé  |
|---------|---------|---------|
|A human readable name for the add-in. (Uniqueness recommended, but not required.)|`Contoso Marketing Excel Add-in (Prod)`|**$ADD-IN-NAME$**|
|Le nom de domaine complet du compl?ment (sauf pour le protocole). *Vous devez utiliser un domaine que vous poss?dez.* C'est pourquoi il n'est pas possible d'utiliser certains domaines connus comme `azurewebsites.net` ou `cloudapp.net`. Le domaine et les sous-domaines doivent être les mêmes que ceux utilisés dans les URL dans la section `<Resources>` du manifeste du complément.|`localhost:6789`, `addins.contoso.com`|**$FQDN-WITHOUT-PROTOCOL$**|
|Les autorisations dont votre compl?ment a besoin pour AAD et Microsoft Graph. (`profile` est toujours requis.)|`profile`, `Files.Read.All`|S/O|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
