---
title: Enregistrer un complément Office utilisant une SSO (authentification unique) au point de terminaison Azure AD v2.0
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 95b690e21bddf7f2754cc308c8b771e629bbc630
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437254"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Enregistrer un complément Office utilisant une SSO (authentification unique) au point de terminaison Azure AD v2.0

Cet article explique comment enregistrer un complément Office au point de terminaison Azure AD v2.0. Vous devez inscrire le complément dès le début de son développement. Lorsque vous passez à la phase de test ou de production, vous pouvez modifier l'enregistrement existant ou créer des enregistrements distincts pour chaque version de développement, de test et de production du complément. 

Le tableau suivant détaille les informations nécessaires pour effectuer cette procédure, ainsi que les espaces réservés correspondant tels qu'ils apparaissent dans les instructions. 

|Information  |Exemples  |Espace réservé  |
|---------|---------|---------|
|Un nom contrôlable de visu pour le complément. (Caractère unique recommandée, mais pas obligatoire)    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|Le nom de domaine complet du complément (sauf pour le protocole). *Vous devez utiliser un domaine que vous possédez.* C'est pourquoi il n'est pas possible d'utiliser certains domaines connus comme `azurewebsites.net` ou `cloudapp.net`.   |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|Les autorisations dont votre complément a besoin pour AAD et Microsoft Graph. (`profile` est toujours requis.)    |`profile`, `Files.Read.All`         |N/A         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]