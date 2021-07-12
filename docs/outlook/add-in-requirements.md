---
title: Conditions requises pour les compléments Outlook
description: Pour que les compléments Outlook se chargent et fonctionnent correctement, les serveurs et les clients doivent répondre à plusieurs conditions.
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 6062073d44a412d67961f806677cd60701bbdb9b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348593"
---
# <a name="outlook-add-in-requirements"></a>Conditions requises pour les compléments Outlook

Pour que les compléments Outlook se chargent et fonctionnent correctement, les serveurs et les clients doivent répondre à plusieurs conditions.

## <a name="client-requirements"></a>Configuration requise du client

- Le client doit être l’un des applications pris en charge pour les compléments Outlook. Les clients suivants prennent en charge les compléments.

  - Outlook 2013 ou version ultérieure sur Windows
  - Outlook 2016 ou version ultérieure sur Mac
  - Outlook sur iOS
  - Outlook sur Android
  - Outlook sur le web pour Exchange 2016 ou une version ultérieure
  - Outlook sur le web pour Exchange 2013
  - Outlook.com

- Le client doit être connecté à un serveur Exchange ou Microsoft 365 par une connexion directe. Lors de la configuration du client, l'utilisateur doit choisir un **Exchange**, **Office**, ou **Outlook.com** type de compte. Si le client est configuré pour se connecter avec POP3 ou IMAP, les add-ins ne se chargeront pas.

## <a name="mail-server-requirements"></a>Configuration requise pour le serveur de messagerie

Si l'utilisateur est connecté à Microsoft 365 ou Outlook.com, les besoins en matière de serveur de messagerie sont déjà pris en charge. Toutefois, pour les utilisateurs connectés à des installations sur site du serveur Exchange, les exigences suivantes s'appliquent.

- Le serveur doit être un serveur Exchange 2013 ou de version ultérieure.
- Les services web Exchange doivent être activés et exposés sur Internet. De nombreux compléments exigent que les services web Exchange fonctionnent correctement.
- Le serveur doit avoir un certificat d’authentification valide pour émettre des jetons d’identité valides. Les nouvelles installations du serveur Exchange incluent un certificat d’authentification par défaut. Pour plus d’informations, reportez-vous aux articles [Certificats numériques et chiffrement dans Exchange 2016](/Exchange/architecture/client-access/certificates) et [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).
- Pour accéder à des compléments à partir d’[AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), les serveurs d’accès au client doivent être en mesure de communiquer avec AppSource.

## <a name="add-in-server-requirements"></a>Conditions de serveur pour le complément

Les fichiers du complément (HTML, JavaScript, etc.) peuvent être hébergés sur n’importe quelle plateforme de serveur web. Les seules conditions sont que le serveur doit être configuré de manière à utiliser le protocole HTTPS et que le certificat SSL doit être approuvé par le client.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](../concepts/requirements-for-running-office-add-ins.md)
- [Application cliente Office et disponibilité de la plateforme pour les compléments Office (section Outlook)](../overview/office-add-in-availability.md#outlook)
- [Prise en charge des ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
